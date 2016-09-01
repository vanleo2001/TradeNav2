VERSION 5.00
Begin VB.Form frmTest3 
   Caption         =   "frmTest3"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11760
End
Attribute VB_Name = "frmTest3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    CenterTheForm Me
    
    g.Styler.StyleForm Me

End Sub

Public Sub ShowMe()

    ShowForm Me, eForm_Nonmodal, frmMain

End Sub

