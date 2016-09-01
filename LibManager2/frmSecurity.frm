VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.12#0"; "gdocx.ocx"
Begin VB.Form frmSecurity 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Library Security"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkExpDate 
      Caption         =   "&Expiration Date:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2190
      Width           =   1455
   End
   Begin VB.CheckBox chkCustID 
      Caption         =   "Customer &ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1110
      Width           =   1215
   End
   Begin gdOCX.gdSelectDate gdExpiration 
      Height          =   315
      Left            =   1860
      TabIndex        =   5
      Top             =   2160
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
   End
   Begin VB.TextBox txtCustID 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   2475
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2115
      Left            =   4320
      TabIndex        =   8
      Top             =   300
      Width           =   1455
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   0
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   435
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Label lblExpDateDesc 
      Caption         =   "If you would like the Library to Expire on a certain date, enter that date below"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3795
   End
   Begin VB.Label lblCustIDDesc 
      Caption         =   $"frmSecurity.frx":0000
      Height          =   675
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   3795
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmSecurity.frm
'' Description: Allow the user to enter in some security for the library
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCustID_Click
'' Description: Enable/Disable the controls based on what the user did here
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCustID_Click()
On Error GoTo ErrSection:

    EnableControls
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "Security.chkCustID.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkExpDate_Click
'' Description: Enable/Disable the controls based on what the user did here
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkExpDate_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "Security.chkExpDate.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form without saving changes
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
    RaiseError "Security.cmdCancel.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Unload the form and save changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOk_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "Security.cmdOK.Click", eGDRaiseError_Show, g.strAppPath
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

    Caption = "Library Security"
    
    txtCustID.Text = ""
    gdExpiration.Value = Date
    chkCustID = vbUnchecked
    chkExpDate = vbUnchecked

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "Security.Form.Load", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and Show the form
'' Inputs:      Customer ID and Expiration Date
'' Returns:     True if OK, False if Cancel
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(strCustID As String, dExpDate As Double) As Boolean
On Error GoTo ErrSection:

    If strCustID <> "" Then
        txtCustID.Text = strCustID
        chkCustID = vbChecked
    End If
    If dExpDate > 0 Then
        gdExpiration.Value = dExpDate
        chkExpDate = vbChecked
    End If
    EnableControls

    ShowForm Me, True
    
    strCustID = ""
    dExpDate = 0
    If m.bOK Then
        If chkCustID = vbChecked Then strCustID = Trim(txtCustID.Text)
        If chkExpDate = vbChecked Then dExpDate = gdExpiration.Value
    End If

ErrExit:
    ShowMe = m.bOK
    Exit Function
    
ErrSection:
    RaiseError "Security.ShowMe", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hits the 'X', unload the form without saving
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "Security.Form.QueryUnload", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Enable txtCustID, (chkCustID = vbChecked)
    Enable gdExpiration, (chkExpDate = vbChecked)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "Security.EnableControls", eGDRaiseError_Raise, g.strAppPath

End Sub
