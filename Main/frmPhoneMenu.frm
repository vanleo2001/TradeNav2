VERSION 5.00
Begin VB.Form frmPhoneMenu 
   Caption         =   "Trade Navigator"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   12150
   Begin VB.Image Image1 
      Height          =   240
      Left            =   420
      Picture         =   "frmPhoneMenu.frx":0000
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmPhoneMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Move 0, 0

End Sub

Public Sub ShowMe()

    ShowForm Me, eForm_Nonmodal, frmMain

End Sub
