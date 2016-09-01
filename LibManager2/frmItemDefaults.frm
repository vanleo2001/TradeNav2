VERSION 5.00
Begin VB.Form frmItemDefaults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Default Item Permissions"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2243
      TabIndex        =   6
      Top             =   2460
      Width           =   1935
   End
   Begin VB.OptionButton optItemPermission 
      Caption         =   "&FULL permissions (can edit and view item)"
      Height          =   360
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   5715
   End
   Begin VB.OptionButton optItemPermission 
      Caption         =   "&PARTIAL permission (can view but not edit without a password)"
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   5715
   End
   Begin VB.OptionButton optItemPermission 
      Caption         =   "&RESTRICTED permission (cannot edit or view without a password)"
      Height          =   360
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   900
      Width           =   5715
   End
   Begin VB.OptionButton optItemPermission 
      Caption         =   "&NO ACCESS permission (no access to item.  It does not show up in menus)"
      Height          =   360
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1260
      Width           =   5715
   End
   Begin VB.TextBox txtDefaultPassword 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   1860
      Width           =   2490
   End
   Begin VB.Label Label6 
      Caption         =   "&Default Password:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1890
      Width           =   1425
   End
End
Attribute VB_Name = "frmItemDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmItemDefaults.frm
'' Description: Allows the user to choose default permissions for items added
''              to a library
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' 06/04/02  D Jarmuth   Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: When the user hits OK, hide the form and let the ShowMe finish
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOk_Click()
On Error GoTo ErrSection:

    Me.Hide

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmItemDefaults.cmdOK.Click", eGDRaiseError_Show, g.strAppPath
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
    RaiseError "frmItemDefaults.Form.KeyDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, center it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = Picture16("kLibrary")
    CenterTheForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmItemDefaults.Form.Load", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Show the form with the defaults passed in, then return the
''              user's choices
'' Inputs:      Security Level, Password
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(SecurityLevel As Byte, strPassword As String)
On Error GoTo ErrSection:

    optItemPermission(SecurityLevel).Value = True
    txtDefaultPassword.Text = strPassword
    
    ShowForm Me, True
    
    If Not Verify Then ShowForm Me, True
    
    Select Case True
        Case optItemPermission(0).Value = True
            SecurityLevel = 0
        Case optItemPermission(1).Value = True
            SecurityLevel = 1
        Case optItemPermission(2).Value = True
            SecurityLevel = 2
        Case optItemPermission(3).Value = True
            SecurityLevel = 3
    End Select
    strPassword = txtDefaultPassword
    
ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    RaiseError "frmItemDefaults.ShowMe", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user unloads the form with the X in the corner, hide the
''              form and let ShowMe finish
'' Inputs:      Whether or not to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection

    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmItemDefaults.Form.QueryUnload", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Verify
'' Description: Make sure that the user enters in a password for a restricted
''              security level, and no password for a non-restricted
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Verify() As Boolean
On Error GoTo ErrSection:

    Verify = True
    
    If optItemPermission(0).Value = True And txtDefaultPassword <> "" Then
        InfBox "No Password is needed if no restrictions are set", "!", , "Error"
        txtDefaultPassword.Text = ""
    ElseIf optItemPermission(0).Value = False And txtDefaultPassword.Text = "" Then
        InfBox "Password is required for this Security Level", "!", , "Error"
        txtDefaultPassword.Text = ""
        MoveFocus txtDefaultPassword
        Verify = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmItemDefaults.Verify", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Function
