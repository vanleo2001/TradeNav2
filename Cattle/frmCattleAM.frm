VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Begin VB.Form frmCattleAM 
   Caption         =   "Form1"
   ClientHeight    =   960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   960
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrGridScrollPressed 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4140
      Top             =   120
   End
   Begin gdOCX.gdAppMail gdCattle 
      Left            =   120
      Top             =   120
      _ExtentX        =   953
      _ExtentY        =   847
      ControlName     =   "TNTurnkey"
   End
   Begin VB.Image imgGreen 
      Height          =   195
      Left            =   780
      Picture         =   "frmCattleAM.frx":0000
      Top             =   660
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgYellow 
      Height          =   195
      Left            =   480
      Picture         =   "frmCattleAM.frx":0286
      Top             =   660
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgRed 
      Height          =   195
      Left            =   180
      Picture         =   "frmCattleAM.frx":050C
      Top             =   660
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "frmCattleAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmCattleAM.frm
'' Description: Form for holding the App-Mail object for the Cattle stuff
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/07/2014   DAJ         Created
'' 04/08/2014   DAJ         Copied the Grid Scroll fix into DLL from NavSuite project to fix error
'' 05/22/2014   DAJ         Renamed g.Turnkey to g.Cattle
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    gdCattle.AutoSendInterval = 1000
    gdCattle.Active = True
    
    g.Styler.StyleForm Me
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAM.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up after ourseleves
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    gdCattle.Unload
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAM.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdCattle_MessageReceived
'' Description: Handle an incoming Cattle message
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdCattle_MessageReceived(msg As gdOCX.gdAppMailMsg)
On Error GoTo ErrSection:

    If Not g.Cattle Is Nothing Then
        g.Cattle.TurnkeyMessageReceived msg.MsgType, msg.Message
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAM.gdCattle_MessageReceived"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrGridScrollPressed_Timer
'' Description: Handle the grid scroll pressed timer going off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrGridScrollPressed_Timer()
On Error Resume Next

    g.AppBridge.TimerStart "frmCattleAM.tmrGridScrollPressed"
    If Not MouseIsPressed Then
        tmrGridScrollPressed.Enabled = False
        GridScrollCheck Nothing, 0, 0, 0, 0, False
    End If
    g.AppBridge.TimerEnd "frmCattleAM.tmrGridScrollPressed", tmrGridScrollPressed.Interval

End Sub

