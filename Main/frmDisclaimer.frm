VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDisclaimer 
   Caption         =   "Warnings and Disclaimer"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7200
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   4680
      TabIndex        =   1
      Top             =   5580
      Width           =   795
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "I understand and &ACCEPT these terms"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   420
      TabIndex        =   0
      Top             =   5580
      Width           =   3675
   End
   Begin RichTextLib.RichTextBox rtbInfo 
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9340
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDisclaimer.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDisclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    bAccepted As Boolean
End Type
Private m As mPrivate

Private Sub cmdCancel_Click()

    Me.Hide

End Sub

Private Sub cmdOk_Click()

    m.bAccepted = True
    Me.Hide

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim nSpace&
    
    nSpace = rtbInfo.Left
    If LimitFormSize(Me, cmdOK.Width + cmdCancel.Width + nSpace * 3, 1500) Then Exit Sub
    
    If cmdCancel.Enabled Then
        cmdOK.Move (Me.ScaleWidth - (cmdOK.Width + cmdCancel.Width)) / 3, _
                    Me.ScaleHeight - cmdOK.Height - nSpace
        cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - cmdOK.Left, cmdOK.Top
    Else
        cmdOK.Move (Me.ScaleWidth - cmdOK.Width) / 2, _
                    Me.ScaleHeight - cmdOK.Height - nSpace
    End If
    
    rtbInfo.Move nSpace, nSpace, Me.ScaleWidth - nSpace * 2, cmdOK.Top - nSpace * 2

End Sub

Public Function ShowMe(ByVal bDisplayOnly As Boolean, Optional ByVal strCaption$ = "") As Boolean

    Dim strFile$

    If Len(strCaption) > 0 Then Me.Caption = strCaption
    If bDisplayOnly Then
        cmdCancel.Enabled = False
        cmdCancel.Visible = False
    Else
        cmdCancel.Enabled = True
        cmdCancel.Visible = True
    End If
    
    CenterTheForm Me
    strFile = App.Path & "\Info\Disclaimer.rtf"
    rtbInfo.TextRTF = FileToString(strFile)
    
    m.bAccepted = False
    ShowForm Me, eForm_Modal
    
    If m.bAccepted Then
        SetIniFileProperty "Disclaimer", FileDate(strFile), "General", g.strIniFile
        ShowMe = True
    ElseIf Not bDisplayOnly Then
        SetIniFileProperty "Disclaimer", 0, "General", g.strIniFile
    End If
    
    Unload Me

End Function
