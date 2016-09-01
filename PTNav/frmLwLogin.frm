VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLwLogin 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMfGlobal 
      Height          =   855
      Left            =   120
      Picture         =   "frmLwLogin.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.PictureBox picLw 
      Height          =   735
      Left            =   120
      Picture         =   "frmLwLogin.frx":0BC1
      ScaleHeight     =   675
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Frame fraLogin 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
      Begin VB.CommandButton cmdNewAccount 
         Caption         =   "Add Logi&n"
         Height          =   315
         Left            =   2580
         TabIndex        =   7
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1020
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cboUserIds 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   420
         Width           =   1395
      End
      Begin VB.OptionButton optUserID 
         Caption         =   "&User ID:"
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   1035
      End
      Begin VB.OptionButton optAccount 
         Caption         =   "&Account:"
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   60
         Width           =   1035
      End
      Begin VB.ComboBox cboAccounts 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label lblPassword 
         Caption         =   "&Password:"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   900
         Width           =   915
      End
   End
   Begin RichTextLib.RichTextBox rtfDisclaimer 
      Height          =   4635
      Left            =   4260
      TabIndex        =   29
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8176
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmLwLogin.frx":1EC0
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   3975
      Begin VB.CommandButton cmdLogin 
         Caption         =   "&Login"
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   12
         Top             =   420
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   1080
         TabIndex        =   13
         Top             =   420
         Width           =   975
      End
      Begin VB.CheckBox chkShowIP 
         Caption         =   "&Show Server Information"
         Height          =   435
         Left            =   2220
         TabIndex        =   14
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lblAgree 
         Caption         =   "Choosing to login states that you agree to the terms and conditions on the right"
         Height          =   435
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   3915
      End
   End
   Begin VB.Frame fraServerInfo 
      Caption         =   "Server Information"
      Height          =   1395
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   4035
      Begin VB.OptionButton optLive 
         Caption         =   "Li&ve"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   675
      End
      Begin VB.OptionButton optDemo 
         Caption         =   "&Demo"
         Height          =   195
         Left            =   1260
         TabIndex        =   27
         Top             =   240
         Width           =   795
      End
      Begin VB.OptionButton optFill 
         Caption         =   "&Fill"
         Height          =   195
         Left            =   2400
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame fraHost 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   540
         Width           =   3495
         Begin VB.TextBox txtServerIP 
            Height          =   285
            Left            =   660
            TabIndex        =   18
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label lblServerIP 
            Caption         =   "Host &IP:"
            Height          =   195
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Frame fraLindWaldock 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   900
         Width           =   3495
         Begin VB.TextBox txtPort 
            Height          =   285
            Left            =   360
            TabIndex        =   21
            Top             =   15
            Width           =   735
         End
         Begin VB.TextBox txtFirm 
            Height          =   315
            Left            =   1620
            TabIndex        =   23
            Top             =   0
            Width           =   495
         End
         Begin VB.TextBox txtSubsystem 
            Height          =   315
            Left            =   3120
            TabIndex        =   25
            Top             =   0
            Width           =   315
         End
         Begin VB.Label lblPort 
            Caption         =   "Po&rt:"
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   30
            Width           =   315
         End
         Begin VB.Label lblFirm 
            Caption         =   "Fir&m:"
            Height          =   255
            Left            =   1200
            TabIndex        =   22
            Top             =   30
            Width           =   435
         End
         Begin VB.Label lblSubsystem 
            Caption         =   "Su&bsystem:"
            Height          =   255
            Left            =   2220
            TabIndex        =   24
            Top             =   30
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmLwLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLwLogin.frm
'' Description: Allow the user to choose their login information for Lind
''              Express brokers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO 80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    nBroker As eTT_AccountType          ' Broker for which this form is getting called
    strBroker As String                 ' Name of the broker for display purposes
    
    strIniFile As String                ' Ini file for the broker
    strLiveIP As String                 ' Default Live Host IP
    strLivePort As String               ' Default Live Host Port
    strLiveFirm As String               ' Default Live Firm
    strLiveSubsystem As String          ' Default Live Subsystem
    strDemoIP As String                 ' Default Demo Host IP
    strDemoPort As String               ' Default Demo Host Port
    strDemoFirm As String               ' Default Demo Firm
    strDemoSubsystem As String          ' Default Demo Subsystem
    strFillIP As String                 ' Default Fill Host IP
    strFillPort As String               ' Default Fill Host Port
    strFillFirm As String               ' Default Fill Firm
    strFillSubsystem As String          ' Default Fill Subsystem
    
    bShowDemo As Boolean                ' Show the demo option button?
    bShowFill As Boolean                ' Show the fill option button?
    
    bOK As Boolean                      ' Did the user click on OK?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Broker, Login, Show IP?, Are we switching?
'' Returns:     True if login, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal nBroker As eTT_AccountType, Optional ByVal strLoginName As String = "", Optional ByVal bShowIP As Boolean = False, Optional ByVal bSwitching As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strLastAccount As String        ' Last account logged into
    Dim strLogins As String             ' Logins from the INI file
    Dim astrLogins As New cGdArray      ' Array of logins from the INI file
    Dim lIndex As Long                  ' Index into a for loop
    Dim strLogin As String              ' Login selected

    m.bOK = False
    
    m.nBroker = nBroker
    m.strBroker = g.Broker.BrokerName(m.nBroker)
    Caption = m.strBroker & " Login Information"
    
    Select Case nBroker
        Case eTT_AccountType_LindWaldock
            m.strIniFile = AddSlash(App.Path) & "LindWaldock.INI"
            m.strLiveIP = GetIniFileProperty("Live", "", "IP", AddSlash(App.Path) & "Provided\LwIps.INI")
            m.strLivePort = GetIniFileProperty("Live", "", "Port", AddSlash(App.Path) & "Provided\LwIps.INI")
            m.strLiveFirm = GetIniFileProperty("Live", "", "Firm", AddSlash(App.Path) & "Provided\LwIps.INI")
            m.strLiveSubsystem = GetIniFileProperty("Live", "", "Subsystem", AddSlash(App.Path) & "Provided\LwIps.INI")
            m.strDemoIP = GetIniFileProperty("Demo", "", "IP", AddSlash(App.Path) & "Provided\LwIps.INI")
            m.strDemoPort = GetIniFileProperty("Demo", "", "Port", AddSlash(App.Path) & "Provided\LwIps.INI")
            m.strDemoFirm = GetIniFileProperty("Demo", "", "Firm", AddSlash(App.Path) & "Provided\LwIps.INI")
            m.strDemoSubsystem = GetIniFileProperty("Demo", "", "Subsystem", AddSlash(App.Path) & "Provided\LwIps.INI")
            m.strFillIP = GetIniFileProperty("Fill", "", "IP", AddSlash(App.Path) & "Provided\LwIps.INI")
            m.strFillPort = GetIniFileProperty("Fill", "", "Port", AddSlash(App.Path) & "Provided\LwIps.INI")
            m.strFillFirm = GetIniFileProperty("Fill", "", "Firm", AddSlash(App.Path) & "Provided\LwIps.INI")
            m.strFillSubsystem = GetIniFileProperty("Fill", "", "Subsystem", AddSlash(App.Path) & "Provided\LwIps.INI")
            
            picLw.Visible = True
            picMfGlobal.Visible = False
                        
        Case eTT_AccountType_ManExpress
            m.strIniFile = AddSlash(App.Path) & "ManExpress.INI"
            m.strLiveIP = GetIniFileProperty("Live", "", "IP", AddSlash(App.Path) & "Provided\MxIps.INI")
            m.strLivePort = GetIniFileProperty("Live", "", "Port", AddSlash(App.Path) & "Provided\MxIps.INI")
            m.strLiveFirm = GetIniFileProperty("Live", "", "Firm", AddSlash(App.Path) & "Provided\MxIps.INI")
            m.strLiveSubsystem = GetIniFileProperty("Live", "", "Subsystem", AddSlash(App.Path) & "Provided\MxIps.INI")
            m.strDemoIP = GetIniFileProperty("Demo", "", "IP", AddSlash(App.Path) & "Provided\MxIps.INI")
            m.strDemoPort = GetIniFileProperty("Demo", "", "Port", AddSlash(App.Path) & "Provided\MxIps.INI")
            m.strDemoFirm = GetIniFileProperty("Demo", "", "Firm", AddSlash(App.Path) & "Provided\MxIps.INI")
            m.strDemoSubsystem = GetIniFileProperty("Demo", "", "Subsystem", AddSlash(App.Path) & "Provided\MxIps.INI")
            m.strFillIP = GetIniFileProperty("Fill", "", "IP", AddSlash(App.Path) & "Provided\MxIps.INI")
            m.strFillPort = GetIniFileProperty("Fill", "", "Port", AddSlash(App.Path) & "Provided\MxIps.INI")
            m.strFillFirm = GetIniFileProperty("Fill", "", "Firm", AddSlash(App.Path) & "Provided\MxIps.INI")
            m.strFillSubsystem = GetIniFileProperty("Fill", "", "Subsystem", AddSlash(App.Path) & "Provided\MxIps.INI")
            
            picLw.Visible = False
            picMfGlobal.Visible = True
            
    End Select
    
    strLastAccount = UCase(GetIniFileProperty("LastAccount", "", "User", m.strIniFile))
    
    LoadCombos
    If (cboAccounts.ListCount > 0) Or (cboUserIds.ListCount > 0) Then
        If SetCombo(strLoginName) = False Then
            If SetCombo(strLastAccount) = False Then
                If cboAccounts.ListCount > 0 Then
                    If optAccount.Value = False Then
                        optAccount.Value = True
                    End If
                    cboAccounts.ListIndex = 0
                ElseIf cboUserIds.ListCount > 0 Then
                    If optUserID.Value = False Then
                        optUserID.Value = True
                    End If
                    cboUserIds.ListIndex = 0
                End If
            End If
        End If
    End If
    
    If ((cboAccounts.ListCount = 0) And (cboUserIds.ListCount = 0)) Or ((cboAccounts.ListCount + cboUserIds.ListCount = 1) And (bSwitching = True)) Then
        NewAccount
    End If
    
    If (cboAccounts.ListCount + cboUserIds.ListCount > 1) Or ((cboAccounts.ListCount + cboUserIds.ListCount = 1) And (bSwitching = False)) Then
        If bShowIP = True Then chkShowIP.Value = vbChecked Else chkShowIP.Value = vbUnchecked
        
        EnableControls
        MoveFocus txtPassword
        
        Form_Resize
        ShowForm Me, eForm_Modal, frmMain
        
        If m.bOK = True Then
            Select Case nBroker
                Case eTT_AccountType_LindWaldock
                    If Not g.LindWaldock Is Nothing Then
                        If optAccount.Value = True Then
                            g.LindWaldock.UserName = ""
                            g.LindWaldock.LoginAccount = cboAccounts.Text
                        Else
                            g.LindWaldock.UserName = cboUserIds.Text
                            g.LindWaldock.LoginAccount = ""
                        End If
                        g.LindWaldock.Password = Trim(txtPassword.Text)
                        g.LindWaldock.HostIP = Trim(txtServerIP.Text)
                        g.LindWaldock.HostPort = Trim(txtPort.Text)
                        g.LindWaldock.Firm = Trim(txtFirm.Text)
                        g.LindWaldock.Subsystem = Trim(txtSubsystem.Text)
                    End If
                                
                Case eTT_AccountType_ManExpress
                    If Not g.ManExpress Is Nothing Then
                        If optAccount.Value = True Then
                            g.ManExpress.UserName = ""
                            g.ManExpress.LoginAccount = cboAccounts.Text
                        Else
                            g.ManExpress.UserName = cboUserIds.Text
                            g.ManExpress.LoginAccount = ""
                        End If
                        g.ManExpress.Password = Trim(txtPassword.Text)
                        g.ManExpress.HostIP = Trim(txtServerIP.Text)
                        g.ManExpress.HostPort = Trim(txtPort.Text)
                        g.ManExpress.Firm = Trim(txtFirm.Text)
                        g.ManExpress.Subsystem = Trim(txtSubsystem.Text)
                    End If
                                
            End Select
        
            If optAccount.Value = True Then
                strLogin = "A:" & cboAccounts.Text
            Else
                strLogin = "U:" & cboUserIds.Text
            End If
            SetIniFileProperty "LastAccount", strLogin, "User", m.strIniFile
            strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
            astrLogins.SplitFields strLogins, ","
        
            Select Case True
                Case optLive
                    If (Trim(txtServerIP.Text) <> m.strLiveIP) Or (Trim(txtPort.Text) <> m.strLivePort) Then
                        SetIniFileProperty "LiveIP", Trim(txtServerIP.Text) & ":" & Trim(txtPort.Text), "IPS", m.strIniFile
                    Else
                        SetIniFileProperty "LiveIP", "", "IPS", m.strIniFile
                    End If
                    
                    If (Trim(txtFirm.Text) <> m.strLiveFirm) Or (Trim(txtSubsystem.Text) <> m.strLiveSubsystem) Then
                        SetIniFileProperty "LiveFirm", Trim(txtFirm.Text) & ":" & Trim(txtSubsystem.Text), "IPS", m.strIniFile
                    Else
                        SetIniFileProperty "LiveFirm", "", "IPS", m.strIniFile
                    End If
                    
                    For lIndex = 0 To astrLogins.Size - 1
                        If Parse(astrLogins(lIndex), ";", 1) = strLogin Then
                            If Parse(astrLogins(lIndex), ";", 2) <> "0" Then
                                astrLogins(lIndex) = strLogin & ";0"
                            End If
                            Exit For
                        End If
                    Next lIndex
                    
                Case optDemo
                    If (Trim(txtServerIP.Text) <> m.strDemoIP) Or (Trim(txtPort.Text) <> m.strDemoPort) Then
                        SetIniFileProperty "DemoIP", Trim(txtServerIP.Text) & ":" & Trim(txtPort.Text), "IPS", m.strIniFile
                    Else
                        SetIniFileProperty "DemoIP", "", "IPS", m.strIniFile
                    End If
                    
                    If (Trim(txtFirm.Text) <> m.strDemoFirm) Or (Trim(txtSubsystem.Text) <> m.strDemoSubsystem) Then
                        SetIniFileProperty "DemoFirm", Trim(txtFirm.Text) & ":" & Trim(txtSubsystem.Text), "IPS", m.strIniFile
                    Else
                        SetIniFileProperty "DemoFirm", "", "IPS", m.strIniFile
                    End If
                    
                    For lIndex = 0 To astrLogins.Size - 1
                        If Parse(astrLogins(lIndex), ";", 1) = strLogin Then
                            If Parse(astrLogins(lIndex), ";", 2) <> "1" Then
                                astrLogins(lIndex) = strLogin & ";1"
                            End If
                            Exit For
                        End If
                    Next lIndex
                    
                Case optFill
                    If (Trim(txtServerIP.Text) <> m.strFillIP) Or (Trim(txtPort.Text) <> m.strFillPort) Then
                        SetIniFileProperty "FillIP", Trim(txtServerIP.Text) & ":" & Trim(txtPort.Text), "IPS", m.strIniFile
                    Else
                        SetIniFileProperty "FillIP", "", "IPS", m.strIniFile
                    End If
                    
                    If (Trim(txtFirm.Text) <> m.strFillFirm) Or (Trim(txtSubsystem.Text) <> m.strFillSubsystem) Then
                        SetIniFileProperty "FillFirm", Trim(txtFirm.Text) & ":" & Trim(txtSubsystem.Text), "IPS", m.strIniFile
                    Else
                        SetIniFileProperty "FillFirm", "", "IPS", m.strIniFile
                    End If
                    
                    For lIndex = 0 To astrLogins.Size - 1
                        If Parse(astrLogins(lIndex), ";", 1) = strLogin Then
                            If Parse(astrLogins(lIndex), ";", 2) <> "2" Then
                                astrLogins(lIndex) = strLogin & ";2"
                            End If
                            Exit For
                        End If
                    Next lIndex
                    
            End Select
            
            SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmLwLogin.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccounts_Click
'' Description: When the user changes the account give the password the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccounts_Click()
On Error GoTo ErrSection:

    Select Case cboAccounts.ItemData(cboAccounts.ListIndex)
        Case 0:
            If optLive.Value = False Then
                optLive.Value = True
            End If
        Case 1:
            If optDemo.Value = False Then
                optDemo.Value = True
            End If
        Case 2:
            If optFill.Value = False Then
                optFill.Value = True
            End If
    End Select
    
    If optAccount.Value = False Then
        optAccount.Value = True
    End If
    MoveFocus txtPassword
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.cboAccounts_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboUserIds_Click
'' Description: When the user changes the user ID give the password the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboUserIds_Click()
On Error GoTo ErrSection:

    Select Case cboUserIds.ItemData(cboUserIds.ListIndex)
        Case 0:
            If optLive.Value = False Then
                optLive.Value = True
            End If
        Case 1:
            If optDemo.Value = False Then
                optDemo.Value = True
            End If
        Case 2:
            If optFill.Value = False Then
                optFill.Value = True
            End If
    End Select
    
    If optUserID.Value = False Then
        optUserID.Value = True
    End If
    MoveFocus txtPassword

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.cboUserIds_Click"
    
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
    RaiseError "frmLwLogin.chkShowIP_Click"
    
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
    RaiseError "frmLwLogin.cmdCancel_Click"
    
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
        InfBox "Please enter in an IP address for the " & m.strBroker & " server", "!", , m.strBroker & " Login Error"
        GoTo ErrExit
    End If

    If Len(Trim(txtPort.Text)) = 0 Then
        chkShowIP.Value = vbChecked
        MoveFocus txtPort
        InfBox "Please enter in a server port for the " & m.strBroker & " server", "!", , m.strBroker & " Login Error"
        GoTo ErrExit
    End If

    If Len(Trim(txtFirm.Text)) = 0 Then
        chkShowIP.Value = vbChecked
        MoveFocus txtFirm
        InfBox "Please enter in a firm for the " & m.strBroker & " server", "!", , m.strBroker & " Login Error"
        GoTo ErrExit
    End If

    If Len(Trim(txtSubsystem.Text)) = 0 Then
        chkShowIP.Value = vbChecked
        MoveFocus txtSubsystem
        InfBox "Please enter in a subsystem for the " & m.strBroker & " server", "!", , m.strBroker & " Login Error"
        GoTo ErrExit
    End If

    If Len(Trim(txtPassword.Text)) = 0 Then
        MoveFocus txtPassword
        InfBox "Please enter in a password to login to the " & m.strBroker & " servers", "!", , m.strBroker & " Login Error"
        GoTo ErrExit
    End If

    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.cmdLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNewAccount_Click
'' Description: Allow the user to enter in a new Lind Express login
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
    RaiseError "frmLwLogin.cmdNewAccount_Click"
    
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
    RaiseError "frmLwLogin.Form_Activate"
    
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
    
    m.bShowDemo = True ' FileExist(AddSlash(App.Path) & "AllowLwDemo.FLG")
    m.bShowFill = FileExist(AddSlash(App.Path) & "AllowLwFill.FLG")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.Form_Load"
    
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
    RaiseError "frmLwLogin.Form_QueryUnload"
    
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
'' Function:    optAccount_Click
'' Description: Change to the selected login in the accounts combo
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAccount_Click()
On Error GoTo ErrSection:

    If cboAccounts.ListCount > 0 Then
        Select Case cboAccounts.ItemData(cboAccounts.ListIndex)
            Case 0:
                If optLive.Value = False Then
                    optLive.Value = True
                End If
            Case 1:
                If optDemo.Value = False Then
                    optDemo.Value = True
                End If
            Case 2:
                If optFill.Value = False Then
                    optFill.Value = True
                End If
        End Select
        MoveFocus txtPassword
        
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.optAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optDemo_Click
'' Description: Change the IP and Port numbers to the demo server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optDemo_Click()
On Error GoTo ErrSection:

    Dim strIP As String                 ' IP address from the INI file
    Dim strFirm As String               ' Firm from the INI file
    
    If Len(m.strIniFile) > 0 Then
        strIP = GetIniFileProperty("DemoIP", "", "IPS", m.strIniFile)
        strFirm = GetIniFileProperty("DemoFirm", "", "IPS", m.strIniFile)
        
        If Len(strIP) = 0 Then
            txtServerIP.Text = m.strDemoIP
            txtPort.Text = m.strDemoPort
        Else
            txtServerIP.Text = Parse(strIP, ":", 1)
            txtPort.Text = Parse(strIP, ":", 2)
        End If
        
        If Len(strFirm) = 0 Then
            txtFirm.Text = m.strDemoFirm
            txtSubsystem.Text = m.strDemoSubsystem
        Else
            txtFirm.Text = Parse(strFirm, ":", 1)
            txtSubsystem.Text = Parse(strFirm, ":", 2)
        End If
        
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.optDemo_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optFill_Click
'' Description: Change the IP and Port numbers to the fill management server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optFill_Click()
On Error GoTo ErrSection:

    Dim strIP As String                 ' IP address from the INI file
    Dim strFirm As String               ' Firm from the INI file
    
    strIP = GetIniFileProperty("FillIP", "", "IPS", m.strIniFile)
    strFirm = GetIniFileProperty("FillFirm", "", "IPS", m.strIniFile)
    
    If Len(m.strIniFile) > 0 Then
        If Len(strIP) = 0 Then
            txtServerIP.Text = m.strFillIP
            txtPort.Text = m.strFillPort
        Else
            txtServerIP.Text = Parse(strIP, ":", 1)
            txtPort.Text = Parse(strIP, ":", 2)
        End If
        
        If Len(strFirm) = 0 Then
            txtFirm.Text = m.strFillFirm
            txtSubsystem.Text = m.strFillSubsystem
        Else
            txtFirm.Text = Parse(strFirm, ":", 1)
            txtSubsystem.Text = Parse(strFirm, ":", 2)
        End If
        
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.optFill_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optLive_Click
'' Description: Change the IP and Port numbers to the live server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optLive_Click()
On Error GoTo ErrSection:

    Dim strIP As String                 ' IP address from the INI file
    Dim strFirm As String               ' Firm from the INI file
    
    If Len(m.strIniFile) > 0 Then
        strIP = GetIniFileProperty("LiveIP", "", "IPS", m.strIniFile)
        strFirm = GetIniFileProperty("LiveFirm", "", "IPS", m.strIniFile)
        
        If Len(strIP) = 0 Then
            txtServerIP.Text = m.strLiveIP
            txtPort.Text = m.strLivePort
        Else
            txtServerIP.Text = Parse(strIP, ":", 1)
            txtPort.Text = Parse(strIP, ":", 2)
        End If
        
        If Len(strFirm) = 0 Then
            txtFirm.Text = m.strLiveFirm
            txtSubsystem.Text = m.strLiveSubsystem
        Else
            txtFirm.Text = Parse(strFirm, ":", 1)
            txtSubsystem.Text = Parse(strFirm, ":", 2)
        End If
        
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.optLive_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optUserID_Click
'' Description: Change to the selected login in the user ID combo
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optUserID_Click()
On Error GoTo ErrSection:

    If cboUserIds.ListCount > 0 Then
        Select Case cboUserIds.ItemData(cboUserIds.ListIndex)
            Case 0:
                If optLive.Value = False Then
                    optLive.Value = True
                End If
            Case 1:
                If optDemo.Value = False Then
                    optDemo.Value = True
                End If
            Case 2:
                If optFill.Value = False Then
                    optFill.Value = True
                End If
        End Select
        MoveFocus txtPassword
        
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.optUserID_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFirm_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFirm_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtFirm

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.txtFirm_GotFocus"
    
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
    RaiseError "frmLwLogin.txtPassword_GotFocus"
    
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
    RaiseError "frmLwLogin.txtPort_GotFocus"
    
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
    RaiseError "frmLwLogin.txtServerIP_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSubsystem_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSubsystem_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtSubsystem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.txtSubsystem_GotFocus"
    
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

    If (Len(cboAccounts.Text) > 0) Or (Len(cboUserIds.Text) > 0) Then
        chkShowIP.Enabled = True
    Else
        chkShowIP.Enabled = False
        chkShowIP.Value = vbUnchecked
    End If
    
    optAccount.Enabled = (cboAccounts.ListCount > 0)
    cboAccounts.Enabled = (cboAccounts.ListCount > 0)
    optUserID.Enabled = (cboUserIds.ListCount > 0)
    cboUserIds.Enabled = (cboUserIds.ListCount > 0)
    
    optFill.Visible = m.bShowFill
    optDemo.Visible = (m.bShowDemo Or m.bShowFill)
    optLive.Visible = (m.bShowDemo Or m.bShowFill)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombos
'' Description: Load the account and user ID combo boxes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombos()
On Error GoTo ErrSection:

    Dim strLogins As String             ' Login information string from INI file
    Dim astrLogins As New cGdArray      ' Array of login information
    Dim lIndex As Long                  ' Index into a for loop
    Dim strAccount As String            ' Account already in the INI file
    Dim strIP As String                 ' IP address from the INI file
    
    ' Load up the accounts combo box...
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
                If Parse(strAccount, ":", 1) = "A" Then
                    cboAccounts.AddItem Parse(strAccount, ":", 2)
                    cboAccounts.ItemData(cboAccounts.NewIndex) = CLng(Val(Parse(astrLogins(lIndex), ";", 2)))
                    If cboAccounts.ItemData(cboAccounts.NewIndex) = 1 Then
                        m.bShowDemo = True
                    ElseIf cboAccounts.ItemData(cboAccounts.NewIndex) = 2 Then
                        m.bShowFill = True
                    End If
                ElseIf Parse(strAccount, ":", 1) = "U" Then
                    cboUserIds.AddItem Parse(strAccount, ":", 2)
                    cboUserIds.ItemData(cboUserIds.NewIndex) = CLng(Val(Parse(astrLogins(lIndex), ";", 2)))
                    If cboUserIds.ItemData(cboUserIds.NewIndex) = 1 Then
                        m.bShowDemo = True
                    ElseIf cboUserIds.ItemData(cboUserIds.NewIndex) = 2 Then
                        m.bShowFill = True
                    End If
                End If
            End If
        Next lIndex
    End If
    
    If cboAccounts.ListCount > 0 Then cboAccounts.ListIndex = 0
    If cboUserIds.ListCount > 0 Then cboUserIds.ListIndex = 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLwLogin.LoadCombos"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCombo
'' Description: Set the accounts combo box or the user ID combo box to the
''              given account if possible
'' Inputs:      Account
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetCombo(ByVal strAccount As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the item found?
    
    bFound = False
    If Len(strAccount) > 0 Then
        If Parse(strAccount, ":", 1) = "U" Then
            If cboUserIds.ListCount > 0 Then
                For lIndex = 0 To cboUserIds.ListCount - 1
                    If UCase(cboUserIds.List(lIndex)) = UCase(Parse(strAccount, ":", 2)) Then
                        bFound = True
                        cboUserIds.ListIndex = lIndex
                        Select Case cboUserIds.ItemData(lIndex)
                            Case 0:
                                If optLive.Value = False Then
                                    optLive.Value = True
                                End If
                            Case 1:
                                If optDemo.Value = False Then
                                    optDemo.Value = True
                                End If
                            Case 2:
                                If optFill.Value = False Then
                                    optFill.Value = True
                                End If
                        End Select
                        
                        Exit For
                    End If
                Next lIndex
                
                If bFound And optUserID.Value = False Then
                    optUserID.Value = True
                End If
            End If
        ElseIf Parse(strAccount, ":", 1) = "A" Then
            
            If cboAccounts.ListCount > 0 Then
                For lIndex = 0 To cboAccounts.ListCount - 1
                    If UCase(cboAccounts.List(lIndex)) = UCase(Parse(strAccount, ":", 2)) Then
                        bFound = True
                        cboAccounts.ListIndex = lIndex
                        Select Case cboAccounts.ItemData(lIndex)
                            Case 0:
                                If optLive.Value = False Then
                                    optLive.Value = True
                                End If
                            Case 1:
                                If optDemo.Value = False Then
                                    optDemo.Value = True
                                End If
                            Case 2:
                                If optFill.Value = False Then
                                    optFill.Value = True
                                End If
                        End Select
                        
                        Exit For
                    End If
                Next lIndex
            
                If bFound And optAccount.Value = False Then
                    optAccount.Value = True
                End If
            End If
        End If
    End If
    
    SetCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLwLogin.SetCombo"
    
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

    Dim strType As String               ' Type of login information
    Dim strAccount As String            ' Account number from the user
    Dim strAcctType As String           ' Account type from the user
    Dim strNewLogin As String           ' New login to save to INI file
    Dim strLogins As String             ' Login string from the INI file
    
    strType = InfBox("Are you entering in a " & m.strBroker & " account number or a " & m.strBroker & " user ID?", "?", "+-Account|User ID", m.strBroker & " Login")
    If strType = "A" Then
        strAccount = "A:" & UCase(InfBox("What is your " & m.strBroker & " account number?", "?", , m.strBroker & " Account Number", , , , , , "string"))
    Else
        strAccount = "U:" & UCase(InfBox("What is your " & m.strBroker & " user ID?", "?", , m.strBroker & " User ID", , , , , , "string"))
    End If
    
    If Len(strAccount) > 0 Then
        If SetCombo(strAccount) = False Then
            If Parse(strAccount, ":", 1) = "A" Then
                cboAccounts.AddItem Parse(strAccount, ":", 2)
                cboAccounts.ItemData(cboAccounts.NewIndex) = 0
                strNewLogin = strAccount & ";0"
            ElseIf Parse(strAccount, ":", 1) = "U" Then
                cboUserIds.AddItem Parse(strAccount, ":", 2)
                cboUserIds.ItemData(cboUserIds.NewIndex) = 0
                strNewLogin = strAccount & ";0"
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
    RaiseError "frmLwLogin.NewAccount"
    
End Sub

