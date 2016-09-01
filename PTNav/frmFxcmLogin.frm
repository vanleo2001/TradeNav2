VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFxcmLogin 
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtfDisclaimer 
      Height          =   4275
      Left            =   4020
      TabIndex        =   15
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7541
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmFxcmLogin.frx":0000
   End
   Begin VB.Frame fraServerInfo 
      Caption         =   "Server Information"
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   3735
      Begin VB.TextBox txtURL 
         Height          =   255
         Left            =   660
         TabIndex        =   14
         Top             =   660
         Width           =   2955
      End
      Begin VB.ComboBox cboConnection 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblURL 
         Caption         =   "&URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   660
         Width           =   435
      End
      Begin VB.Label lblConnection 
         Caption         =   "C&onnection:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.ComboBox cboAccount 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1380
      Width           =   1395
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton cmdNewAccount 
      Caption         =   "Add Logi&n"
      Height          =   315
      Left            =   2700
      TabIndex        =   2
      Top             =   1380
      Width           =   1155
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   3735
      Begin VB.CommandButton cmdLogin 
         Caption         =   "&Login"
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   7
         Top             =   420
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   1080
         TabIndex        =   8
         Top             =   420
         Width           =   975
      End
      Begin VB.CheckBox chkShowIP 
         Caption         =   "&Show Server Information"
         Height          =   435
         Left            =   2220
         TabIndex        =   9
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lblAgree 
         Caption         =   "Choosing to login states that you agree to the terms and conditions on the right"
         Height          =   435
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   3675
      End
   End
   Begin VB.Image imgFxcm 
      Height          =   1125
      Left            =   502
      Picture         =   "frmFxcmLogin.frx":008B
      Top             =   120
      Width           =   2970
   End
   Begin VB.Label lblAccount 
      Caption         =   "&User Name:"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label lblPassword 
      Caption         =   "&Password:"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   1860
      Width           =   915
   End
End
Attribute VB_Name = "frmFxcmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmFxcmLogin.frm
'' Description: Allow the user to choose their login information
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO 80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strIniFile As String                ' INI file
    strRealUrl As String                ' URL for the real server
    strDemoUrl As String                ' URL for the demo server
    
    bOK As Boolean                      ' Did the user press OK or Cancel?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Account, Show IP?, Are we switching?
'' Returns:     True if login, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Optional ByVal strAccount As String = "", Optional ByVal bShowIP As Boolean = False, Optional ByVal bSwitching As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strLastAccount As String        ' Last account logged into
    Dim strLogins As String             ' Logins from the INI file
    Dim astrLogins As New cGdArray      ' Array of logins from the INI file
    Dim lIndex As Long                  ' Index into a for loop

    m.bOK = False
    m.strIniFile = AddSlash(App.Path) & "Fxcm.INI"
    m.strDemoUrl = GetIniFileProperty("Demo", "", "IP", AddSlash(App.Path) & "Provided\FxcmIps.INI")
    m.strRealUrl = GetIniFileProperty("Real", "", "IP", AddSlash(App.Path) & "Provided\FxcmIps.INI")
    
    strLastAccount = UCase(GetIniFileProperty("LastAccount", "", "User", m.strIniFile))
    LoadCombo
    If cboAccount.ListCount > 0 Then
        If SetCombo(strAccount) = False Then
            If SetCombo(strLastAccount) = False Then
                strLastAccount = GetIniFileProperty("UserID", "", "User", m.strIniFile)
                If SetCombo(strLastAccount) = False Then
                    cboAccount.ListIndex = 0
                End If
            End If
        End If
    End If

    If (cboAccount.ListCount = 0) Or ((cboAccount.ListCount = 1) And (bSwitching = True)) Then
        NewAccount
    End If
    
    If (cboAccount.ListCount > 1) Or ((cboAccount.ListCount = 1) And (bSwitching = False)) Then
        If bShowIP = True Then chkShowIP.Value = vbChecked Else chkShowIP.Value = vbUnchecked
        
        MoveFocus txtPassword

        ShowForm Me, eForm_Modal, frmMain
        
        If m.bOK = True Then
            If Not g.FXCM Is Nothing Then
                g.FXCM.UserName = cboAccount.Text
                g.FXCM.Password = Trim(txtPassword.Text)
                g.FXCM.URL = Trim(txtURL.Text)
                g.FXCM.Connection = Trim(cboConnection.Text)
                
                SetIniFileProperty "LastAccount", cboAccount.Text, "User", m.strIniFile
                strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
                astrLogins.SplitFields strLogins, ","
                
                Select Case UCase(cboConnection.Text)
                    Case "REAL"
                        If Trim(txtURL.Text) <> m.strRealUrl Then
                            SetIniFileProperty "Real", Trim(txtURL.Text), "IPS", m.strIniFile
                        Else
                            SetIniFileProperty "Real", "", "IPS", m.strIniFile
                        End If
                        
                        For lIndex = 0 To astrLogins.Size - 1
                            If Parse(astrLogins(lIndex), ";", 1) = cboAccount.Text Then
                                If Parse(astrLogins(lIndex), ";", 2) <> "0" Then
                                    astrLogins(lIndex) = cboAccount.Text & ";0"
                                End If
                                
                                Exit For
                            End If
                        Next lIndex
                        
                    Case "DEMO"
                        If Trim(txtURL.Text) <> m.strDemoUrl Then
                            SetIniFileProperty "Demo", Trim(txtURL.Text), "IPS", m.strIniFile
                        Else
                            SetIniFileProperty "Demo", "", "IPS", m.strIniFile
                        End If
                        
                        For lIndex = 0 To astrLogins.Size - 1
                            If Parse(astrLogins(lIndex), ";", 1) = cboAccount.Text Then
                                If Parse(astrLogins(lIndex), ";", 2) <> "1" Then
                                    astrLogins(lIndex) = cboAccount.Text & ";1"
                                End If
                                
                                Exit For
                            End If
                        Next lIndex
                        
                End Select
                
                SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile
            End If
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmFxcmLogin.ShowMe"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccount_Click
'' Description: When the user changes the account give the access key the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccount_Click()
On Error GoTo ErrSection:

    Select Case cboAccount.ItemData(cboAccount.ListIndex)
        Case 0:
            cboConnection.Text = "Real"
        Case 1:
            cboConnection.Text = "Demo"
    End Select
    MoveFocus txtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.cboAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboConnection_Click
'' Description: If the user changes the connection box, change the URL
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboConnection_Click()
On Error GoTo ErrSection:

    Select Case UCase(cboConnection.Text)
        Case "REAL"
            txtURL.Text = m.strRealUrl
            
        Case "DEMO"
            txtURL.Text = m.strDemoUrl
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.cboConnection_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkShowIP_Click
'' Description: Show/Hide the Server IP information as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkShowIP_Click()
On Error GoTo ErrSection:

    Form_Resize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.chkShowIP_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the user to cancel out of the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLogin_Click
'' Description: Allow the user to login to the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLogin_Click()
On Error GoTo ErrSection:

    If Len(Trim(txtURL.Text)) = 0 Then
        chkShowIP.Value = vbChecked
        MoveFocus txtURL
        InfBox "Please enter in the URL for the FXCM server", "!", , "FXCM Login Error"
        GoTo ErrExit
    End If

    If Len(Trim(txtPassword.Text)) = 0 Then
        MoveFocus txtPassword
        InfBox "Please enter in an password to login to the FXCM servers", "!", , "FXCM Login Error"
        GoTo ErrExit
    End If

    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.cmdLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNewAccount_Click
'' Description: Allow the user to enter in a new FXCM login account number
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNewAccount_Click()
On Error GoTo ErrSection:

    NewAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.cmdNewAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Make sure when the form is activated that access key gets focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    MoveFocus txtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = Picture16("kBlank")
    Caption = "FXCM Login Information"
    
    rtfDisclaimer.Locked = True
    rtfDisclaimer.Font.Name = "Arial"
    rtfDisclaimer.BackColor = vbButtonFace
    rtfDisclaimer.TextRTF = FileToString(AddSlash(App.Path) & "Info\BrokerDisclaimer.rtf")
    
    CenterTheForm Me
    
    cboConnection.AddItem "Real"
    cboConnection.AddItem "Demo"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Cancel out of the form if the user clicks on the X
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    If chkShowIP.Value = vbChecked Then
        fraServerInfo.Visible = True
    Else
        fraServerInfo.Visible = False
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_GotFocus
'' Description: When the control gets the focus, select all of the text
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
    RaiseError "frmFxcmLogin.txtPassword_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtURL_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtURL_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtURL

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.txtURL_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    If Len(cboAccount.Text) > 0 Then
        chkShowIP.Enabled = True
    Else
        chkShowIP.Enabled = False
        chkShowIP.Value = vbUnchecked
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load the accounts combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo()
On Error GoTo ErrSection:

    Dim strLogins As String             ' Login information string from INI file
    Dim astrLogins As New cGdArray      ' Array of login information
    Dim lIndex As Long                  ' Index into a for loop
    Dim strAccount As String            ' Account already in the INI file
    Dim strIP As String                 ' IP address from the INI file
    
    strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
    If strLogins <> UCase(strLogins) Then
        strLogins = UCase(strLogins)
        SetIniFileProperty "Logins", strLogins, "User", m.strIniFile
    End If
    
    If Len(strLogins) > 0 Then
        astrLogins.SplitFields strLogins, ","
        For lIndex = 0 To astrLogins.Size - 1
            strAccount = Parse(astrLogins(lIndex), ";", 1)
            If Len(strAccount) > 0 Then
                cboAccount.AddItem strAccount
                cboAccount.ItemData(cboAccount.NewIndex) = CLng(Val(Parse(astrLogins(lIndex), ";", 2)))
            End If
        Next lIndex
    End If
    
    If cboAccount.ListCount = 0 Then
        strAccount = UCase(GetIniFileProperty("UserID", "", "User", m.strIniFile))
        If Len(strAccount) > 0 Then
            cboAccount.AddItem strAccount
            cboAccount.ItemData(cboAccount.NewIndex) = 0
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.LoadCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCombo
'' Description: Set the accounts combo box to the given account if possible
'' Inputs:      Account
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetCombo(ByVal strAccount As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the item found?
    
    bFound = False
    If (cboAccount.ListCount > 0) And (Len(strAccount) > 0) Then
        For lIndex = 0 To cboAccount.ListCount - 1
            If UCase(cboAccount.List(lIndex)) = UCase(strAccount) Then
                bFound = True
                cboAccount.ListIndex = lIndex
                Select Case cboAccount.ItemData(lIndex)
                    Case 0:
                        cboConnection.Text = "Real"
                    Case 1:
                        cboConnection.Text = "Demo"
                End Select
            End If
        Next lIndex
    End If
    
    SetCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmFxcmLogin.SetCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewAccount
'' Description: Allow the user to give us a new account number
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewAccount()
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account number from the user
    Dim strAcctType As String           ' Account type from the user
    Dim strNewLogin As String           ' New login to save to INI file
    Dim strLogins As String             ' Login string from the INI file
    
    strAccount = UCase(InfBox("What is your FXCM user name?", "?", , "FXCM User Name", , , , , , "string"))
    If Len(strAccount) > 0 Then
        If SetCombo(strAccount) = False Then
            strAcctType = InfBox("Is this a Real or a Demo user?", "?", "+-Real|Demo", "FXCM Account Type")
            If UCase(strAcctType) = "R" Then
                cboAccount.AddItem strAccount
                cboAccount.ItemData(cboAccount.NewIndex) = 0
                strNewLogin = strAccount & ";0"
            Else
                cboAccount.AddItem strAccount
                cboAccount.ItemData(cboAccount.NewIndex) = 1
                strNewLogin = strAccount & ";1"
            End If
            
            SetCombo strAccount
            MoveFocus txtPassword
            
            strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
            If Len(strLogins) = 0 Then
                strLogins = strNewLogin
            Else
                strLogins = strLogins & "," & strNewLogin
            End If
            SetIniFileProperty "Logins", strLogins, "User", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFxcmLogin.NewAccount"
    
End Sub
