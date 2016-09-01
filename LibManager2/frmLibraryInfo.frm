VERSION 5.00
Begin VB.Form frmLibraryInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   2145
      TabIndex        =   11
      Top             =   3960
      Width           =   2535
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   1320
         TabIndex        =   13
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame fraPermissions 
      Caption         =   "Library Permissions"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   6495
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   3720
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optRestricted 
         Caption         =   "&Password Required to View or Edit Library:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton optFull 
         Caption         =   "&No Password Required to View or Edit Library"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "General Information"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txtAuthor 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtDescription 
         Height          =   885
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblAuthor 
         Caption         =   "&Author:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblDescription 
         Caption         =   "&Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblName 
         Caption         =   "&Library Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmLibraryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLibraryInfo.frm
'' Description: Allow the user to enter in some necessary library information
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click OK or Cancel?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Name, Author, Description, Permission, Password
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(strName As String, strAuthor As String, strDesc As String, bFull As Boolean, strPassword As String) As Boolean
On Error GoTo ErrSection:

    txtName.Text = strName
    If Len(strAuthor) > 0 Then
        txtAuthor.Text = strAuthor
    Else
        txtAuthor.Text = GetIniFileProperty("Author", "", "Library", g.strIniFile)
    End If
    txtDescription.Text = strDesc
    optFull.Value = bFull
    optRestricted.Value = Not bFull
    txtPassword.Text = strPassword

    ShowForm Me, eForm_Modal, g.frmOwner
    
    If m.bOK Then
        strName = Trim(txtName.Text)
        strAuthor = Trim(txtAuthor.Text)
        strDesc = Trim(txtDescription.Text)
        bFull = optFull.Value
        If optRestricted.Value = True Then strPassword = Trim(txtPassword.Text)
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmLibraryInfo.ShowMe", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Close the form without retaining the information
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
    RaiseError "frmLibraryInfo.cmdCancel.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Verify the information entered, then close the form and retain
''              the information the user entered
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOk_Click()
On Error GoTo ErrSection:

    If Len(Trim(txtName.Text)) = 0 Then
        MoveFocus txtName
        InfBox "You must enter a name for the library", "!", , "Library Error"
        Exit Sub
    End If

    If Len(Trim(txtName.Text)) > 50 Then
        MoveFocus txtName
        InfBox "Library Name must be less than 50 characters in length", "!", , "Library Error"
        Exit Sub
    End If
    
    If Len(Trim(txtAuthor.Text)) = 0 Then
        MoveFocus txtAuthor
        InfBox "You must enter an author for the library", "!", , "Library Error"
        Exit Sub
    End If
    
    If Len(Trim(txtAuthor.Text)) > 50 Then
        MoveFocus txtAuthor
        InfBox "Author must be less than 50 characters in length", "!", , "Library Error"
        Exit Sub
    End If
    
    If Len(Trim(txtDescription.Text)) > 255 Then
        MoveFocus txtDescription
        InfBox "Description must be less than 255 characters in length", "!", , "Library Error"
        Exit Sub
    End If
    
    If optRestricted.Value = True And Len(Trim(txtPassword.Text)) = 0 Then
        MoveFocus txtPassword
        InfBox "You must enter a password for a restricted library", "!", , "Library Error"
        Exit Sub
    End If
    
    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryInfo.cmdOK.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

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

    Icon = Picture16("kLibrary")
    Caption = "Library Information"
    CenterTheForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryInfo.Form.Load", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: When the user clicks the X, allow ShowMe to unload the form
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.bOK = False
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryInfo.Form.QueryUnload", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and size controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    ' Make sure to center the buttons frame...
    With fraButtons
        .Move ScaleWidth / 2 - .Width / 2
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAuthor_GotFocus
'' Description: Upon receiving the focus, select all of the text in the box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAuthor_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtAuthor

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryInfo.txtAuthor.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDescription_GotFocus
'' Description: Upon receiving the focus, select all of the text in the box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDescription_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtDescription

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryInfo.txtDescription.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtName_GotFocus
'' Description: Upon receiving the focus, select all of the text in the box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtName_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryInfo.txtName.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_GotFocus
'' Description: Upon receiving the focus, select all of the text in the box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryInfo.txtPassword.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub
