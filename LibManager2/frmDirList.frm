VERSION 5.00
Begin VB.Form frmDirList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select a Location to Save your file"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSelectedLocation 
      Height          =   330
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3990
      Width           =   4875
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   5340
      TabIndex        =   3
      Top             =   975
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   450
      Left            =   5340
      TabIndex        =   2
      Top             =   390
      Width           =   1125
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   225
      TabIndex        =   1
      Top             =   345
      Width           =   4875
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   225
      TabIndex        =   0
      Top             =   1065
      Width           =   4875
   End
   Begin VB.Label Label1 
      Caption         =   "Selected File:"
      Height          =   195
      Index           =   2
      Left            =   255
      TabIndex        =   7
      Top             =   3765
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Folder:"
      Height          =   225
      Index           =   1
      Left            =   255
      TabIndex        =   5
      Top             =   825
      Width           =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "Drive:"
      Height          =   225
      Index           =   0
      Left            =   225
      TabIndex        =   4
      Top             =   120
      Width           =   1560
   End
End
Attribute VB_Name = "frmDirList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    bOK As Boolean
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form without saving
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
    RaiseError "frmDirList.cmdCancel.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Unload the form and return the selected path
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDirList.cmdOK.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Dir1_Change
'' Description: Change the text box as the user changes the tree
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Dir1_Change()
On Error GoTo ErrSection:

    txtSelectedLocation.Text = FileNameDisplay(Dir1.Path)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "DirPath.Dir1.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Dir1_Click
'' Description: Change the text box when the user clicks on a path
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Dir1_Click()
On Error GoTo ErrSection:

    txtSelectedLocation.Text = FileNameDisplay(Dir1.Path)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "DirPath.Dir1.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Drive1_Change
'' Description: Update the text box and the directory tree on a drive change
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Drive1_Change()
On Error GoTo ErrSection:

    Dir1.Path = Drive1
    Dir1.Refresh
    txtSelectedLocation.Text = FileNameDisplay(Dir1.Path)

ErrExit:
    Exit Sub

ErrSection:
    If Err.Number = 68 Then
        InfBox "Please insert a disk into drive " & UCase(Drive1), "!", , "Error"
    Else
        RaiseError "DirPath.Drive1.Change", eGDRaiseError_Show, g.strAppPath
    End If
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form and controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = Picture16("kBlank")
    CenterTheForm Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "DirPath.Form.Load", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      None
'' Returns:     Path selected
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe() As String
On Error GoTo ErrSection:
    
    Drive1.Drive = g.strAppPath
    Dir1.Path = g.strAppPath
    
    ShowForm Me, True
    If m.bOK Then ShowMe = FileNameDisplay(txtSelectedLocation.Text)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDirList.ShowMe", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Unload the form if the user hits the 'X'
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        m.bOK = False
        Me.Hide
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDirList.Form.QueryUnload", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub
