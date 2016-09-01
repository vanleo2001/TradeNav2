VERSION 5.00
Begin VB.Form frmFlowChart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Designing and Testing a Strategy ..."
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "frmFlowChart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.Image img 
      Height          =   5775
      Left            =   120
      Picture         =   "frmFlowChart.frx":030A
      Top             =   360
      Width           =   10290
   End
End
Attribute VB_Name = "frmFlowChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFlowChart.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    img.Move 0, 0
    Me.Width = img.Width + (Me.Width - Me.ScaleWidth)
    Me.Height = img.Height + (Me.Height - Me.ScaleHeight)
    CenterTheForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFlowChart.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        WindowStateX(Me) = wsMinimized
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFlowChart.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub
