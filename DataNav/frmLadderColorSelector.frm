VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmLadderColorSelector 
   Caption         =   "Select color"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   218
      TabIndex        =   3
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
      Caption         =   "frmLadderColorSelector.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLadderColorSelector.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLadderColorSelector.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   1230
         TabIndex        =   2
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
         Caption         =   "frmLadderColorSelector.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLadderColorSelector.frx":0092
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLadderColorSelector.frx":00B2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   60
         TabIndex        =   5
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
         Caption         =   "frmLadderColorSelector.frx":00CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLadderColorSelector.frx":00F2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLadderColorSelector.frx":0112
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClearAll 
         Height          =   375
         Left            =   2400
         TabIndex        =   4
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
         Caption         =   "frmLadderColorSelector.frx":012E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLadderColorSelector.frx":0160
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLadderColorSelector.frx":0180
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraColorSelector 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
      Caption         =   "frmLadderColorSelector.frx":019C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLadderColorSelector.frx":01BC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLadderColorSelector.frx":01DC
      RightToLeft     =   0   'False
      Begin VB.Timer tmr 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3360
         Top             =   360
      End
      Begin gdOCX.gdSelectColor gdOutlineColor 
         Height          =   375
         Left            =   2648
         TabIndex        =   1
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP lblOutline 
         Height          =   255
         Left            =   248
         Top             =   180
         Width           =   2415
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
         Caption         =   "frmLadderColorSelector.frx":01F8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLadderColorSelector.frx":0246
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLadderColorSelector.frx":0266
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmLadderColorSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    frmTDGrid As frmTickDistribution
    nGridRow As Long            'for color selector
End Type

Private m As mPrivate

Public Sub ShowMe(frmCaller As Form, Optional ByVal nRow& = -1, _
    Optional ByVal nLastUsedColor& = -1, Optional ByVal strPrice$ = "")
On Error GoTo ErrSection:

    tmr.Enabled = False

    If frmCaller Is Nothing Then Exit Sub
    If nRow <= 0 Then Exit Sub
    
    Set Me.Icon = Picture16("kBlank")
    
    Set m.frmTDGrid = frmCaller
    m.nGridRow = nRow

    If nLastUsedColor = -1 Then nLastUsedColor = vbYellow
    gdOutlineColor.Color = nLastUsedColor
    lblOutline.Caption = "Outline " & strPrice & " with color:"

    Me.Move m.frmTDGrid.Left + (m.frmTDGrid.Width - Me.Width) / 2, m.frmTDGrid.Top + (m.frmTDGrid.Height - Me.Height) / 2
    ShowForm Me, eForm_Modal, m.frmTDGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLadderColorSelector.ShowMe", eGDRaiseError_Show

End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    
    If m.nGridRow > 0 Then m.frmTDGrid.OutlineCell m.nGridRow, 0, True
    Unload Me

End Sub

Private Sub cmdClearAll_Click()
On Error Resume Next

    m.frmTDGrid.OutlineCell -1, 0, True
    Unload Me

End Sub

Private Sub cmdOK_Click()
On Error Resume Next

    Me.Hide

    If m.nGridRow > 0 Then
        m.frmTDGrid.OutlineCell m.nGridRow, gdOutlineColor.Color, False
    End If
    
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    tmr.Enabled = False
    Set m.frmTDGrid = Nothing

End Sub

Private Sub gdOutlineColor_Changed()
On Error Resume Next
    
    m.frmTDGrid.OutlineCell m.nGridRow, gdOutlineColor.Color, False
    tmr.Enabled = True

End Sub

Private Sub tmr_Timer()
On Error Resume Next

    tmr.Enabled = False
    Unload Me

End Sub

