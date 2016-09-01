VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWelcome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Welcome to Genesis"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   435
      Left            =   1073
      TabIndex        =   1
      Top             =   4560
      Width           =   4635
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   3180
         TabIndex        =   4
         Top             =   0
         Width           =   1395
      End
      Begin VB.CommandButton cmdUser 
         Caption         =   "&Genesis User"
         Height          =   435
         Left            =   1590
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "&Register Now"
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1455
      End
   End
   Begin RichTextLib.RichTextBox rtbMessage 
      Height          =   4275
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7541
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmWelcome.frx":0000
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmWelcome.frm
'' Description: Shows the welcome message to the user with the options to
''              register or to enter in their user information
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' 04/29/02  D Jarmuth   Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Enum eWelcome
    eWelcome_Register
    eWelcome_GenesisUser
    eWelcome_Cancel
End Enum

Private Type mPrivate
    Return As eWelcome
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:   cmdCancel_Click
'' Desription: If the user clicks on the Cancel button, set the return code and
''             let the ShowMe take over
'' Inputs:     None
'' Returns:    None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.Return = eWelcome_Cancel
    Me.Hide

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmWelcome.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:   cmdRegister_Click
'' Desription: If the user clicks on the Register button, set the return code and
''             let the ShowMe take over
'' Inputs:     None
'' Returns:    None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRegister_Click()
On Error GoTo ErrSection:

    m.Return = eWelcome_Register
    Me.Hide

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmWelcome.cmdRegister.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:   cmdUser_Click
'' Desription: If the user clicks on the User Info button, set the return code and
''             let the ShowMe take over
'' Inputs:     None
'' Returns:    None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdUser_Click()
On Error GoTo ErrSection:

    m.Return = eWelcome_GenesisUser
    Me.Hide

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmWelcome.cmdUser.Click", eGDRaiseError_Show
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
    RaiseError "frmWelcome.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:   Form_Load
'' Desription: When the form is loaded, center it and load the Welcome.RTF file
'' Inputs:     None
'' Returns:    None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strText As String
    Dim strFile As String
    
    If Not FileExist(App.Path & "\Provided\Install.CFG") Then
        If gbShowRegistration Then
            strFile = AddSlash(App.Path) & "Info\Welcome.RTF"
        Else
            strFile = AddSlash(App.Path) & "Info\Welcome2.RTF"
            cmdRegister.Visible = False
            cmdUser.Caption = "&OK"
            fraButtons.Left = fraButtons.Left - (cmdUser.Left / 2)
        End If
    Else
        strFile = AddSlash(App.Path) & "Info\Welcome.RTF"
    End If
    
    Me.Icon = Picture16(ToolbarIcon("ID_News"))
    CenterTheForm Me
    
    If Not FileExist(strFile) Then
        m.Return = eWelcome_Cancel
        Me.Hide
        Exit Sub
    End If
    
    strText = FileToString(strFile)
    rtbMessage.TextRTF = strText

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmWelcome.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:   ShowMe
'' Desription: Show the form and return the users response
'' Inputs:     None
'' Returns:    Cancel, Register, or User Info
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe() As eWelcome
On Error GoTo ErrSection:

    ShowForm Me, True
    
    ShowMe = m.Return
    Unload Me

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmWelcome.ShowMe", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:   Form_QueryUnload
'' Desription: If the user closes the form with the X, return a Cancel code and
''             let the ShowMe take over
'' Inputs:     Whether or not to Cancel the Unload, Mode of the Unload
'' Returns:    None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        m.Return = eWelcome_Cancel
        Me.Hide
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmWelcome.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub
