VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmWindowLink 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6045
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   1890
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   375
      Left            =   1380
      TabIndex        =   0
      Top             =   360
      Width           =   1035
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
      Caption         =   "frmWindowLink.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmWindowLink.frx":002E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":004E
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   14
      Left            =   120
      Top             =   4800
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
      Caption         =   "frmWindowLink.frx":006A
      BackColor       =   10485760
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":00A2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":00C2
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   13
      Left            =   120
      Top             =   4500
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
      Caption         =   "frmWindowLink.frx":00DE
      BackColor       =   10526720
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":0116
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":0136
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   12
      Left            =   120
      Top             =   4200
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
      Caption         =   "frmWindowLink.frx":0152
      BackColor       =   40960
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":018A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":01AA
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   11
      Left            =   120
      Top             =   3900
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
      Caption         =   "frmWindowLink.frx":01C6
      BackColor       =   41120
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":01FE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":021E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   10
      Left            =   120
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
      Caption         =   "frmWindowLink.frx":023A
      BackColor       =   24736
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":0272
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":0292
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   9
      Left            =   120
      Top             =   3300
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
      Caption         =   "frmWindowLink.frx":02AE
      BackColor       =   160
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":02E6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":0306
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   8
      Left            =   120
      Top             =   3000
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
      Caption         =   "frmWindowLink.frx":0322
      BackColor       =   10485920
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":035A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":037A
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   15
      Left            =   120
      Top             =   5100
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
      Caption         =   "frmWindowLink.frx":0396
      BackColor       =   6316128
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":03CE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":03EE
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   16
      Left            =   120
      Top             =   600
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
      Caption         =   "frmWindowLink.frx":040A
      BackColor       =   1
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":044A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":046A
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   7
      Left            =   120
      Top             =   2700
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
      Caption         =   "frmWindowLink.frx":0486
      BackColor       =   16711680
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":04BE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":04DE
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   6
      Left            =   120
      Top             =   2400
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
      Caption         =   "frmWindowLink.frx":04FA
      BackColor       =   16776960
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":0532
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":0552
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   5
      Left            =   120
      Top             =   2100
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
      Caption         =   "frmWindowLink.frx":056E
      BackColor       =   65280
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":05A6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":05C6
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   4
      Left            =   120
      Top             =   1800
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
      Caption         =   "frmWindowLink.frx":05E2
      BackColor       =   65535
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":061A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":063A
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   3
      Left            =   120
      Top             =   1500
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
      Caption         =   "frmWindowLink.frx":0656
      BackColor       =   33023
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":068E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":06AE
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   2
      Left            =   120
      Top             =   1200
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
      Caption         =   "frmWindowLink.frx":06CA
      BackColor       =   255
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":0702
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":0722
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   1
      Left            =   120
      Top             =   900
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
      Caption         =   "frmWindowLink.frx":073E
      BackColor       =   16711935
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":0776
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":0796
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblHdr 
      Height          =   195
      Left            =   120
      Top             =   60
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
      Caption         =   "frmWindowLink.frx":07B2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":07F6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":0816
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   240
      Index           =   0
      Left            =   120
      Top             =   300
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
      Caption         =   "frmWindowLink.frx":0832
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":0868
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":0888
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblDesc 
      Height          =   615
      Left            =   180
      Top             =   5400
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
      Caption         =   "frmWindowLink.frx":08A4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmWindowLink.frx":0944
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWindowLink.frx":0964
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmWindowLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kNumColors = 16 '8

Public Enum eLinkMode
    eLink_Symbol = 0
    eLink_Period = 1
End Enum

Private Type mPrivate
    frmCalledFrom As Form
    eMode As eLinkMode
End Type
Private m As mPrivate

Public Sub ShowMe(frmCalledFrom As Form, ByVal eMode As eLinkMode, ByVal nX&, ByVal nY&)
On Error GoTo ErrSection:

    Dim i&, iForm&, nColor&, nDiff&, bShowBlack As Boolean
    Dim pt As POINTAPI
    Dim frm As Form

    If Me.Visible Then Exit Sub

    m.eMode = eMode
    Set m.frmCalledFrom = frmCalledFrom

    nDiff = lblColor(kNumColors).Top - lblColor(0).Top
    If (eMode = eLink_Symbol) And Not IsFrmChart(frmCalledFrom) Then
        bShowBlack = True
    End If
    lblColor(kNumColors).Visible = bShowBlack
    i = 6075 '3690
    If Not bShowBlack Then
        i = i - nDiff
    End If
    Me.Height = i

    For i = 1 To kNumColors
        lblColor(i).Tag = ""
        If i <> kNumColors Then
            lblColor(i).Caption = ""
            If bShowBlack Then
                lblColor(i).Top = lblColor(kNumColors).Top + nDiff * i
            Else
                lblColor(i).Top = lblColor(0).Top + nDiff * i
            End If
        End If
    Next
    lblDesc.Top = lblColor(kNumColors - 1).Top + nDiff

    If eMode = eLink_Period Then
        lblHdr = Replace(lblHdr, "Symbol", "Period")
        lblDesc = Replace(lblDesc, "symbol", "period")
        lblDesc = "(bar period will auto- sync for all charts with the same color link)"
    End If

    For iForm = 0 To Forms.Count - 1
        Set frm = Forms(iForm)
        nColor = 0
        On Error Resume Next
        If eMode = eLink_Period Then
            nColor = frm.WindowLink.PeriodColor
        Else
            nColor = frm.WindowLink.SymbolColor
        End If
        On Error GoTo ErrSection:
        
        If nColor > 0 Then
            For i = 1 To kNumColors - 1
                If lblColor(i).BackColor = nColor Then
                    If lblColor(i).Caption = "" Then
                        If eMode = eLink_Period Then
                            lblColor(i).Tag = Str(frm.Periodicity)
                            lblColor(i).Caption = GetPeriodStr(frm.Periodicity)
                        Else
                            lblColor(i).Tag = Str(frm.SymbolID)
                            lblColor(i).Caption = GetSymbol(frm.SymbolID)
                        End If
                    End If
                End If
            Next
        End If
    Next
    Set frm = Nothing

    ' locate this form right below the link button
    pt.X = nX \ Screen.TwipsPerPixelX
    pt.Y = nY \ Screen.TwipsPerPixelY
    ClientToScreen m.frmCalledFrom.hWnd, pt
    pt.Y = pt.Y * Screen.TwipsPerPixelY
    If eMode = eLink_Period Then
        pt.X = pt.X * Screen.TwipsPerPixelX - Me.Width / 4#
    Else
        pt.X = pt.X * Screen.TwipsPerPixelX - Me.Width * 3 / 4#
    End If

    Me.Move pt.X, pt.Y
    
    DoEvents
    ShowForm Me ', eForm_Modal
    
    ' make the calling form "look active" while this form actually receives the keyboard input
    FakeActiveLook m.frmCalledFrom.hWnd, Me.hWnd
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWindowLink.ShowMe", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWindowLink.cmdCancel_Click", eGDRaiseError_Show
End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWindowLink.Form_Deactivate", eGDRaiseError_Show
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    cmdCancel.Top = -cmdCancel.Height - 500
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWindowLink.Form_Load", eGDRaiseError_Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.frmCalledFrom = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWindowLink.Form_Unload", eGDRaiseError_Show
End Sub

Private Sub lblColor_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim nID&, nColor&, nVseBackColor&, nVseForeColor&
    Dim frm As Form

    ' need to use local variables in case this form unloads during a deactivate before this routine finishes
    nColor = lblColor(Index).BackColor
    nID = Val(lblColor(Index).Tag)
    nVseBackColor = nColor
    nVseForeColor = RGB(255, 255, 255)
    If Index = 0 Then
        nID = 0
        nColor = 0
        nVseForeColor = 0
    ElseIf Index = kNumColors Then
        Set frm = ActiveChart
        If Not frm Is Nothing Then
            nID = frm.SymbolID
        End If
    End If
    Set frm = m.frmCalledFrom
    
    If m.eMode = eLink_Period Then
        If IsFrmChart(frm) Then
            frm.vsePeriodLink.BackColor = nVseBackColor
            frm.vsePeriodLink.ForeColor = nVseForeColor
        End If
        frm.WindowLink.PeriodColor = nColor
        If nID <> 0 Then
            frm.Periodicity = nID
        End If
    Else
        If IsFrmChart(frm) Then
            frm.vseSymbolLink.BackColor = nVseBackColor
            frm.vseSymbolLink.ForeColor = nVseForeColor
        End If
        frm.WindowLink.SymbolColor = nColor
        If nID <> 0 Then
            frm.SymbolID = nID
        End If
        If Index = 0 Then
            SetIniFileProperty frm.Name, 0, "LinkDefaults", g.strIniFile
        ElseIf Index = kNumColors Then
            SetIniFileProperty frm.Name, 1, "LinkDefaults", g.strIniFile
        End If
    End If
    
    Set frm = Nothing
    g.bDirtyChartPage = True
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    Set frm = Nothing
    RaiseError "frmWindowLink.lblColor_Click", eGDRaiseError_Show
End Sub

