VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOecLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   915
      Left            =   120
      TabIndex        =   6
      Top             =   900
      Width           =   3495
      Begin VB.CommandButton cmdLogin 
         Caption         =   "&Login"
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   1080
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkShowIP 
         Caption         =   "&Show Server Information"
         Height          =   435
         Left            =   2220
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblAgree 
         Caption         =   "Choosing to login states that you agree to the terms and conditions on the right"
         Height          =   435
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.Frame fraServerInfo 
      Caption         =   "Server Information"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   3495
      Begin VB.CommandButton cmdRestore 
         Caption         =   "&Restore Default Server Information"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtServerIP 
         Height          =   285
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   2700
         TabIndex        =   15
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblServerIP 
         Caption         =   "&IP:"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   285
         Width           =   315
      End
      Begin VB.Label lblPort 
         Caption         =   "P&ort:"
         Height          =   195
         Left            =   2280
         TabIndex        =   14
         Top             =   285
         Width           =   315
      End
   End
   Begin VB.Frame fraLoginInfo 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmdNewLogin 
         Caption         =   "Add Logi&n"
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         Top             =   0
         Width           =   1155
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   360
         Width           =   2475
      End
      Begin VB.ComboBox cboUserName 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblPassword 
         Caption         =   "&Password:"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   420
         Width           =   975
      End
      Begin VB.Label lblUserName 
         Caption         =   "&User Name:"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   60
         Width           =   975
      End
   End
   Begin RichTextLib.RichTextBox rtfDisclaimer 
      Height          =   2775
      Left            =   3780
      TabIndex        =   17
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4895
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmOecLogin.frx":0000
   End
End
Attribute VB_Name = "frmOecLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmOecLogin.frm
'' Description: Dialog to allow user to enter login information for Open E-Cry
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user press OK or Cancel?
    
    strIniFile As String                ' INI file
    strBrokerName As String             ' Broker name
    
    strDefaultIP As String              ' Default IP address
    strDefaultPort As String            ' Default port
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Broker, User Name, Show IP?, Are we switching?
'' Returns:     True if login, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal nBroker As eTT_AccountType, Optional ByVal strUserName As String = "", Optional ByVal bShowIP As Boolean = False, Optional ByVal bSwitching As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strLastUser As String           ' Last user logged into
    Dim strConnectIni As String         ' INI file for connection information
    Dim Broker As cBroker               ' Broker object
    Dim strIP As String                 ' IP Address for the server
    Dim strPort As String               ' Port for the server

    m.bOK = False
    Set Broker = g.Broker.Broker(nBroker)
    If Not Broker Is Nothing Then
        strConnectIni = Broker.ConnectIni
        m.strIniFile = Broker.IniFile
        m.strBrokerName = Broker.BrokerName
        m.strDefaultIP = GetIniFileProperty("IP", "", "Server", strConnectIni)
        m.strDefaultPort = GetIniFileProperty("Port", "", "Server", strConnectIni)
        strIP = GetIniFileProperty("IP", "", "Overrides", m.strIniFile)
        If Len(strIP) = 0 Then
            strIP = m.strDefaultIP
        End If
        strPort = GetIniFileProperty("Port", "", "Overrides", m.strIniFile)
        If Len(strPort) = 0 Then
            strPort = m.strDefaultPort
        End If
        
        Caption = m.strBrokerName & " Login Information"
        txtServerIP.Text = strIP
        txtPort.Text = strPort
        
        strLastUser = UCase(GetIniFileProperty("LastUser", "", "User", m.strIniFile))
        LoadCombo
        If cboUserName.ListCount > 0 Then
            If SetCombo(strUserName) = False Then
                If SetCombo(strLastUser) = False Then
                    strLastUser = GetIniFileProperty("UserName", "", "User", m.strIniFile)
                    If SetCombo(strLastUser) = False Then
                        cboUserName.ListIndex = 0
                    End If
                End If
            End If
        End If
        
        If (cboUserName.ListCount = 0) Or ((cboUserName.ListCount = 1) And (bSwitching = True)) Then
            NewLogin
        End If
        
        If (cboUserName.ListCount > 1) Or ((cboUserName.ListCount = 1) And (bSwitching = False)) Then
            CheckBoxValue(chkShowIP) = bShowIP
            
            MoveFocus txtPassword
    
            ShowForm Me, eForm_Modal, frmMain
            
            If m.bOK = True Then
                Select Case nBroker
                    Case eTT_AccountType_Oec
                        If Not g.Oec Is Nothing Then
                            g.Oec.UserName = cboUserName.Text
                            g.Oec.Password = Trim(txtPassword.Text)
                            g.Oec.IP = Trim(txtServerIP.Text)
                            g.Oec.Port = Trim(txtPort.Text)
                        End If
                        
                    Case eTT_AccountType_OptionsXpress
                        If Not g.OptXpress Is Nothing Then
                            g.OptXpress.UserName = cboUserName.Text
                            g.OptXpress.Password = Trim(txtPassword.Text)
                            g.OptXpress.IP = Trim(txtServerIP.Text)
                            g.OptXpress.Port = Trim(txtPort.Text)
                        End If
                
                End Select
            
                SetIniFileProperty "LastUser", cboUserName.Text, "User", m.strIniFile
            End If
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmOecLogin.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboUserName_Click
'' Description: When the user changes the user name give the password the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboUserName_Click()
On Error GoTo ErrSection:

    MoveFocus txtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.cboUserName_Click"
    
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

    fraServerInfo.Visible = CheckBoxValue(chkShowIP)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.chkShowIP_Click"
    
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
    RaiseError "frmOecLogin.cmdCancel_Click"
    
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

    If Len(Trim(txtServerIP.Text)) = 0 Then
        chkShowIP.Value = vbChecked
        MoveFocus txtServerIP
        InfBox "Please enter in an IP address for the " & m.strBrokerName & " server", "!", , m.strBrokerName & " Login Error"
        GoTo ErrExit
    End If

    If Len(Trim(txtPort.Text)) = 0 Then
        chkShowIP.Value = vbChecked
        MoveFocus txtPort
        InfBox "Please enter in the port for the " & m.strBrokerName & " server", "!", , m.strBrokerName & " Login Error"
        GoTo ErrExit
    End If

    If Len(Trim(txtPassword.Text)) = 0 Then
        MoveFocus txtPassword
        InfBox "Please enter in a password", "!", , m.strBrokerName & " Login Error"
        GoTo ErrExit
    End If
    
    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.cmdLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNewLogin_Click
'' Description: Allow the user to enter in a new login user name
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNewLogin_Click()
On Error GoTo ErrSection:

    NewLogin

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.cmdNewLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRestore_Click
'' Description: Allow the user to restore default server information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRestore_Click()
On Error GoTo ErrSection:

    txtServerIP.Text = m.strDefaultIP
    txtPort.Text = m.strDefaultPort
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.cmdRestore_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Make sure when the form is activated that the password gets focus
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
    RaiseError "frmOecLogin.Form_Activate"
    
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

    rtfDisclaimer.Locked = True
    rtfDisclaimer.Font.Name = "Arial"
    rtfDisclaimer.BackColor = vbButtonFace
    rtfDisclaimer.TextRTF = FileToString(AddSlash(App.Path) & "Info\BrokerDisclaimer.rtf")
    
    CenterTheForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.Form_Load"
    
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
    RaiseError "frmOecLogin.Form_QueryUnload"
    
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
    RaiseError "frmOecLogin.txtPassword_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPort_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPort_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.txtPort_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPort_LostFocus
'' Description: If the user is overriding the port, save the override
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPort_LostFocus()
On Error GoTo ErrSection:

    Dim strPort As String               ' Port that is in the text box
    
    strPort = Trim(txtPort.Text)
    If Len(strPort) = 0 Then
        txtPort.Text = m.strDefaultPort
    ElseIf strPort = m.strDefaultPort Then
        SetIniFileProperty "Port", "", "Overrides", m.strIniFile
    Else
        SetIniFileProperty "Port", strPort, "Overrides", m.strIniFile
    End If
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.txtPort_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtServerIP_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtServerIP_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtServerIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.txtServerIP_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtServerIP_LostFocus
'' Description: Check to see if the server needs to change based on the IP
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtServerIP_LostFocus()
On Error GoTo ErrSection:

    Dim strIP As String                 ' Server IP that is in the text box
    
    strIP = Trim(txtServerIP.Text)
    If Len(strIP) = 0 Then
        txtServerIP.Text = m.strDefaultIP
    ElseIf strIP = m.strDefaultIP Then
        SetIniFileProperty "IP", "", "Overrides", m.strIniFile
    Else
        SetIniFileProperty "IP", strIP, "Overrides", m.strIniFile
    End If
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.txtServerIP_LostFocus"
    
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

    cmdRestore.Enabled = ((Trim(txtServerIP.Text) <> m.strDefaultIP) Or (Trim(txtPort.Text) <> m.strDefaultPort))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.EnableControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load the user name combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo()
On Error GoTo ErrSection:

    Dim strLogins As String             ' Login information string from INI file
    Dim astrLogins As New cGdArray      ' Array of login information
    Dim lIndex As Long                  ' Index into a for loop
    
    strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
    If Len(strLogins) > 0 Then
        astrLogins.SplitFields strLogins, ","
        For lIndex = 0 To astrLogins.Size - 1
            cboUserName.AddItem astrLogins(lIndex)
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.LoadCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCombo
'' Description: Set the user name combo box to the given user name if possible
'' Inputs:      User Name
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetCombo(ByVal strUserName As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the item found?
    
    bFound = False
    If (cboUserName.ListCount > 0) And (Len(strUserName) > 0) Then
        For lIndex = 0 To cboUserName.ListCount - 1
            If UCase(cboUserName.List(lIndex)) = UCase(strUserName) Then
                bFound = True
                cboUserName.ListIndex = lIndex
            End If
        Next lIndex
    End If
    
    SetCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOecLogin.SetCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewLogin
'' Description: Allow the user to give us a new user name
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewLogin()
On Error GoTo ErrSection:

    Dim strUserName As String           ' User Name from the user
    Dim strLogins As String             ' Login string from the INI file
    
    strUserName = InfBox("What is your " & m.strBrokerName & " user name?", "?", , m.strBrokerName & " User Name", , , , , , "string")
    If Len(strUserName) > 0 Then
        If SetCombo(strUserName) = False Then
            cboUserName.AddItem strUserName
            
            SetCombo strUserName
            MoveFocus txtPassword
            
            strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
            If Len(strLogins) = 0 Then
                strLogins = strUserName
            Else
                strLogins = strLogins & "," & strUserName
            End If
            SetIniFileProperty "Logins", strLogins, "User", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOecLogin.NewLogin"
    
End Sub
