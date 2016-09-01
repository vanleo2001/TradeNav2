VERSION 5.00
Begin VB.Form frmSamePriceOrders 
   Caption         =   "Select order action"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optOrderAction 
      Caption         =   "Consolidate into one order "
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   1
      Top             =   1140
      Width           =   3555
   End
   Begin VB.OptionButton optOrderAction 
      Caption         =   "Cancel existing order and place new one"
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   2
      Top             =   1500
      Width           =   3555
   End
   Begin VB.OptionButton optOrderAction 
      Caption         =   "Do nothing (make no changes)"
      Height          =   195
      Index           =   2
      Left            =   165
      TabIndex        =   3
      Top             =   1860
      Value           =   -1  'True
      Width           =   3555
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1748
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.OptionButton optOrderAction 
      Caption         =   "Leave existing order and place new one"
      Height          =   195
      Index           =   3
      Left            =   1980
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label lblMsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   488
      TabIndex        =   0
      Top             =   180
      Width           =   3615
   End
   Begin VB.Label lblOrderAll 
      Height          =   495
      Left            =   2220
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "frmSamePriceOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    nAction As Long
End Type

Private m As mPrivate

Public Function ShowMe(ByVal strMsg$, ByVal strConsolidate$) As Long
On Error GoTo ErrSection:
        
    m.nAction = 2                   'default to make no changes
    optOrderAction(2).Value = True
        
    lblMsg.Caption = strMsg
    If Len(strConsolidate) > 0 Then
        optOrderAction(0).Caption = strConsolidate
    Else
        optOrderAction(0).Enabled = False
    End If
    
    CenterTheForm Me
    ShowForm Me, eForm_Modal
    
    ShowMe = m.nAction
    
    Exit Function

ErrSection:
    RaiseError "frmSamePriceOrders.ShowMe"
    
End Function

Private Sub cmdOk_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = Picture16(ToolbarIcon("ID_TickDistribution"))
End Sub

Private Sub optOrderAction_Click(Index As Integer)
On Error Resume Next
    m.nAction = Index
End Sub
