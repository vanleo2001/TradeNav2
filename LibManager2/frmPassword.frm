VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2550
      TabIndex        =   4
      Top             =   1320
      Width           =   1260
   End
   Begin VB.CommandButton Corner 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4305
      TabIndex        =   3
      Top             =   1485
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   915
      TabIndex        =   2
      Top             =   1320
      Width           =   1260
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   293
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter the password:"
      Height          =   435
      Left            =   300
      TabIndex        =   0
      Top             =   180
      Width           =   4095
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmPassword.frm
'' Description: Asks the user for a Password
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Private Type mPrivate
    strPassword As String               ' Password the user typed in
    bOK As Boolean                      ' Did the user hit OK or Cancel?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user hits Cancel, unload the form
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
    RaiseError "frmPassword.cmdCancel.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form loads, initialize the necessary variables
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
       
    Screen.MousePointer = vbDefault
    
    'Center the form
    CenterTheForm Me
    Me.Icon = Picture16("kSelect")
        
    txtPassword.Text = GetIniFileProperty("LastPasswordUsed", 0, "Misc", g.strIniFile)
    m.strPassword = txtPassword.Text
    
    txtPassword.SelLength = Len(txtPassword.Text)
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPassword.Form.Load", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user hits OK, return the password
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOk_Click()
On Error GoTo ErrSection:

    If Len(txtPassword.Text) < 5 Then
        InfBox "Password must be 5 or more characters", "i", , "Error"
        Exit Sub
    End If
    
    m.strPassword = txtPassword.Text
    m.bOK = True
    
ErrExit:
    Me.Hide
    Exit Sub

ErrSection:
    RaiseError "frmPassword.cmdOK.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hits the X in the corner, it means they are
''              cancelling out of this form, so set the cancel property to true
'' Inputs:      Whether or not to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
    
    If UnloadMode = vbFormControlMenu Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPassword.Form.QueryUnload", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Show the form and return the password if the user hit OK
'' Inputs:      Optional Item to show in the Label
'' Returns:     Password typed in or empty string if cancelled
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Optional ByVal strItem As String = "") As String
On Error GoTo ErrSection:

    If strItem <> "" Then
        Label1.Caption = "Please enter the password for:" & vbCrLf & strItem
    End If
    
    ShowForm Me, True
    
    ShowMe = ""
    If m.bOK Then
        ShowMe = m.strPassword
    End If
    
    Unload Me
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPassword.ShowMe", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Function
