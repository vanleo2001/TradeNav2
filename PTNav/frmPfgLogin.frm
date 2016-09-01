VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPfgLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFintec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   570
      Picture         =   "frmPfgLogin.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   2595
      TabIndex        =   22
      Top             =   60
      Width           =   2595
   End
   Begin VB.Frame fraLoginInfo 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   675
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   3495
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   0
         Width           =   1275
      End
      Begin VB.TextBox txtAccessKey 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   360
         Width           =   2475
      End
      Begin VB.CommandButton cmdNewAccount 
         Caption         =   "Add Logi&n"
         Height          =   315
         Left            =   2280
         TabIndex        =   17
         Top             =   0
         Width           =   1155
      End
      Begin VB.Label lblAccount 
         Caption         =   "&Account:"
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   60
         Width           =   975
      End
      Begin VB.Label lblAccessKey 
         Caption         =   "Access &Key:"
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   420
         Width           =   975
      End
   End
   Begin VB.PictureBox picCTG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   780
      Picture         =   "frmPfgLogin.frx":0E2A
      ScaleHeight     =   855
      ScaleWidth      =   2955
      TabIndex        =   15
      Top             =   240
      Width           =   2955
   End
   Begin VB.CheckBox chkBrokerOCO 
      Caption         =   "Submit &Order-Cancel Order links to PFG"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2100
      Width           =   3495
   End
   Begin RichTextLib.RichTextBox rtfDisclaimer 
      Height          =   4335
      Left            =   3780
      TabIndex        =   14
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7646
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmPfgLogin.frx":19D4
   End
   Begin VB.PictureBox picPFG 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   367
      Picture         =   "frmPfgLogin.frx":1A5F
      ScaleHeight     =   750
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   300
      Width           =   3000
   End
   Begin VB.Frame fraServerInfo 
      Caption         =   "Server Information"
      Height          =   1035
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   3495
      Begin VB.ComboBox cboServers 
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   225
         Width           =   2595
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   2700
         TabIndex        =   13
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox txtServerIP 
         Height          =   285
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label lblServer 
         Caption         =   "&Server:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   285
         Width           =   555
      End
      Begin VB.Label lblPort 
         Caption         =   "Po&rt:"
         Height          =   195
         Left            =   2280
         TabIndex        =   12
         Top             =   645
         Width           =   315
      End
      Begin VB.Label lblServerIP 
         Caption         =   "I&P:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   645
         Width           =   315
      End
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   915
      Left            =   120
      TabIndex        =   2
      Top             =   2460
      Width           =   3495
      Begin VB.CheckBox chkShowIP 
         Caption         =   "Show Server &Information"
         Height          =   435
         Left            =   2220
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   1080
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "&Login"
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblAgree 
         Caption         =   "Choosing to login states that you agree to the terms and conditions on the right"
         Height          =   435
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmPfgLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmPfgLogin.frm
'' Description: Allow the user to choose their PFG login information
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO 80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 09/16/2009   DAJ         Allow user to use Genesis password if access key stored
'' 10/07/2009   DAJ         Added support for Linked Orders at broker
'' 04/21/2010   DAJ         Took out flag file check for allowing Broker OCO's
'' 03/29/2011   DAJ         Only allow login to "B" account if authorized
'' 12/14/2011   DAJ         Added Capital Trading Group and Fintec for PFG
'' 01/03/2012   DAJ         Added the Fintec logo
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strIniFile As String                ' INI file
    strBrokerName As String             ' Broker name
    strConnectIni As String             ' INI file for default connection information
    
    astrIps As cGdArray                 ' Array of IP address information
    astrPortOverrides As cGdArray       ' Array of Port overrides
    bOK As Boolean                      ' Did the user press OK or Cancel?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Broker, Account, Show IP?, Are we switching?
'' Returns:     True if login, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal nBroker As eTT_AccountType, Optional ByVal strAccount As String = "", Optional ByVal bShowIP As Boolean = False, Optional ByVal bSwitching As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strLastAccount As String        ' Last account logged into
    Dim strLogins As String             ' Logins from the INI file
    Dim astrLogins As New cGdArray      ' Array of logins from the INI file
    Dim lIndex As Long                  ' Index into a for loop

    m.bOK = False
    
    Select Case nBroker
        Case eTT_AccountType_CtgPfg
            m.strIniFile = g.CtgPfg.IniFile
            m.strBrokerName = g.CtgPfg.BrokerName
            m.strConnectIni = g.CtgPfg.ConnectIni
            
            CheckBoxValue(chkBrokerOCO) = False
            chkBrokerOCO.Visible = False
            
            picPFG.Visible = False
            picCTG.Visible = True
            picFintec.Visible = False
            
        Case eTT_AccountType_FintecPfg
            m.strIniFile = g.FintecPfg.IniFile
            m.strBrokerName = g.FintecPfg.BrokerName
            m.strConnectIni = g.FintecPfg.ConnectIni
            
            CheckBoxValue(chkBrokerOCO) = False
            chkBrokerOCO.Visible = False
            
            picPFG.Visible = False
            picCTG.Visible = False
            picFintec.Visible = True
            
        Case eTT_AccountType_PFG
            m.strIniFile = g.PFG.IniFile
            m.strBrokerName = g.PFG.BrokerName
            m.strConnectIni = g.PFG.ConnectIni
            
            chkBrokerOCO.Visible = True
            
            picPFG.Visible = True
            picCTG.Visible = False
            picFintec.Visible = False
            
    End Select
    
    Caption = m.strBrokerName & " Login Information"
    
    LoadServersCombo
    FixLogins
    
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
        
        MoveFocus txtAccessKey

        ShowForm Me, eForm_Modal, frmMain
        
        If m.bOK = True Then
            Select Case nBroker
                Case eTT_AccountType_CtgPfg
                    If Not g.CtgPfg Is Nothing Then
                        g.CtgPfg.UserName = cboAccount.Text
                        g.CtgPfg.Password = Trim(txtAccessKey.Text)
                        g.CtgPfg.HostIP = Trim(txtServerIP.Text)
                        g.CtgPfg.HostPort = Trim(txtPort.Text)
                        g.CtgPfg.Server = Parse(m.astrIps(cboServers.ItemData(cboServers.ListIndex)), ":", 4)
                        
                        SetIniFileProperty "LastAccount", cboAccount.Text, "User", m.strIniFile
                        strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
                        astrLogins.SplitFields strLogins, ","
                        
                        If cboServers.ItemData(cboServers.ListIndex) = "0" Then
                            SetIniFileProperty "IP0", Trim(txtServerIP.Text), "Overrides", m.strIniFile
                        End If
                        
                        For lIndex = 0 To astrLogins.Size - 1
                            If Parse(astrLogins(lIndex), "|", 1) = cboAccount.Text Then
                                If Parse(astrLogins(lIndex), "|", 2) <> Str(cboServers.ItemData(cboServers.ListIndex)) Then
                                    astrLogins(lIndex) = cboAccount.Text & "|" & Str(cboServers.ItemData(cboServers.ListIndex))
                                End If
                                Exit For
                            End If
                        Next lIndex
                                            
                        SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile
                    End If
                    
                Case eTT_AccountType_FintecPfg
                    If Not g.FintecPfg Is Nothing Then
                        g.FintecPfg.UserName = cboAccount.Text
                        g.FintecPfg.Password = Trim(txtAccessKey.Text)
                        g.FintecPfg.HostIP = Trim(txtServerIP.Text)
                        g.FintecPfg.HostPort = Trim(txtPort.Text)
                        g.FintecPfg.Server = Parse(m.astrIps(cboServers.ItemData(cboServers.ListIndex)), ":", 4)
                        
                        SetIniFileProperty "LastAccount", cboAccount.Text, "User", m.strIniFile
                        strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
                        astrLogins.SplitFields strLogins, ","
                        
                        If cboServers.ItemData(cboServers.ListIndex) = "0" Then
                            SetIniFileProperty "IP0", Trim(txtServerIP.Text), "Overrides", m.strIniFile
                        End If
                        
                        For lIndex = 0 To astrLogins.Size - 1
                            If Parse(astrLogins(lIndex), "|", 1) = cboAccount.Text Then
                                If Parse(astrLogins(lIndex), "|", 2) <> Str(cboServers.ItemData(cboServers.ListIndex)) Then
                                    astrLogins(lIndex) = cboAccount.Text & "|" & Str(cboServers.ItemData(cboServers.ListIndex))
                                End If
                                Exit For
                            End If
                        Next lIndex
                                            
                        SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile
                    End If
                    
                Case eTT_AccountType_PFG
                    If Not g.PFG Is Nothing Then
                        g.PFG.UserName = cboAccount.Text
                        g.PFG.Password = Trim(txtAccessKey.Text)
                        g.PFG.HostIP = Trim(txtServerIP.Text)
                        g.PFG.HostPort = Trim(txtPort.Text)
                        g.PFG.Server = Parse(m.astrIps(cboServers.ItemData(cboServers.ListIndex)), ":", 4)
                        
                        SetIniFileProperty "LastAccount", cboAccount.Text, "User", m.strIniFile
                        strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
                        astrLogins.SplitFields strLogins, ","
                        
                        If cboServers.ItemData(cboServers.ListIndex) = "0" Then
                            SetIniFileProperty "IP0", Trim(txtServerIP.Text), "Overrides", m.strIniFile
                        End If
                        
                        For lIndex = 0 To astrLogins.Size - 1
                            If Parse(astrLogins(lIndex), "|", 1) = cboAccount.Text Then
                                If Parse(astrLogins(lIndex), "|", 2) <> Str(cboServers.ItemData(cboServers.ListIndex)) Then
                                    astrLogins(lIndex) = cboAccount.Text & "|" & Str(cboServers.ItemData(cboServers.ListIndex))
                                End If
                                Exit For
                            End If
                        Next lIndex
                                            
                        SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile
                    End If
                    
            End Select
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmPfgLogin.ShowMe"
    
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

    CheckBoxValue(chkBrokerOCO) = g.Broker.HoldOcoAtBroker(cboAccount.Text)
    SetServersCombo cboAccount.ItemData(cboAccount.ListIndex)
    MoveFocus txtAccessKey

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin.cboAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboServers_Click
'' Description: Change the IP and Port based on the server selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboServers_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into the array
    
    lIndex = cboServers.ItemData(cboServers.ListIndex)
    txtServerIP.Text = Parse(m.astrIps(lIndex), ":", 2)
    If Len(m.astrPortOverrides(lIndex)) > 0 Then
        txtPort.Text = m.astrPortOverrides(lIndex)
    Else
        txtPort.Text = Parse(m.astrIps(lIndex), ":", 3)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin.cboServers_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkBrokerOCO_Click
'' Description: Save the value if the user changes it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkBrokerOCO_Click()
On Error GoTo ErrSection:

    If CheckBoxValue(chkBrokerOCO) <> g.Broker.HoldOcoAtBroker(cboAccount.Text) Then
        g.Broker.HoldOcoAtBroker(cboAccount.Text) = CheckBoxValue(chkBrokerOCO)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin.chkBrokerOCO_Click"
    
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
    RaiseError "frmPfgLogin.chkShowIP_Click"
    
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
    RaiseError "frmPfgLogin.cmdCancel_Click"
    
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

    Dim strAccessKey As String          ' Access Key user typed in
    Dim strRegAccessKey As String       ' Access Key out of the registry
    Dim strErrorCaption As String       ' Caption for an error dialog
    
    strErrorCaption = m.strBrokerName & " Login Error"

    If Len(Trim(txtServerIP.Text)) = 0 Then
        chkShowIP.Value = vbChecked
        MoveFocus txtServerIP
        InfBox "Please enter in an IP address for the server", "!", , strErrorCaption
        GoTo ErrExit
    End If

    If Len(Trim(txtPort.Text)) = 0 Then
        chkShowIP.Value = vbChecked
        MoveFocus txtPort
        InfBox "Please enter in a server port for the server", "!", , strErrorCaption
        GoTo ErrExit
    End If
    
    strAccessKey = Trim(txtAccessKey.Text)
    If Len(strAccessKey) = 0 Then
        MoveFocus txtAccessKey
        InfBox "Please enter in an Access Key to login to the servers", "!", , strErrorCaption
        GoTo ErrExit
        
    ' DAJ 09/16/2009: If the user types in their Genesis password for the access key,
    ' make an attempt to load the PFG Access Key from the registry.  If it is stored there,
    ' use it.  Otherwise, if the Genesis password doesn't look like an access key
    ' (seven digit numeric), then make the user put in a valid access key...
    ElseIf UCase(strAccessKey) = UCase(RI_GetUserPassword) Then
        strRegAccessKey = RI_GetPfgAccessKey(cboAccount.Text)
        If Len(strRegAccessKey) = 0 Then
            If (Len(strAccessKey) <> 7) Or (IsNumeric(strAccessKey) = False) Then
                txtAccessKey.Text = ""
                MoveFocus txtAccessKey
                InfBox "Please enter in a valid Access Key to login to the servers", "!", , strErrorCaption
                GoTo ErrExit
            End If
        Else
            txtAccessKey.Text = strRegAccessKey
        End If
    End If

    If UCase(Left(cboAccount.Text, 1)) = "B" Then
        If HasModule("B_PFGB") = False Then
            MoveFocus cboAccount
            InfBox "You are not authorized to login with a 'B' account", "!", , strErrorCaption
            GoTo ErrExit
        End If
    ElseIf Left(cboAccount.Text, 1) <> "D" Then
        If Not FileExist(AddSlash(App.Path) & "PfgLive.FLG") Then
            MoveFocus cboAccount
            InfBox "You are not currently authorized to trade a live " & m.strBrokerName & " account with Trade Navigator", "!", , strErrorCaption
            GoTo ErrExit
        End If
    End If
    
    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin.cmdLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNewAccount_Click
'' Description: Allow the user to enter in a new PFG login account number
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
    RaiseError "frmPfgLogin.cmdNewAccount_Click"
    
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

    MoveFocus txtAccessKey

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin.Form_Activate"
    
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
    
    Set m.astrIps = New cGdArray
    m.astrIps.Create eGDARRAY_Strings
    Set m.astrPortOverrides = New cGdArray
    m.astrPortOverrides.Create eGDARRAY_Strings
    
    picCTG.Move picPFG.Left, picPFG.Top
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin.Form_Load"
    
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
    RaiseError "frmPfgLogin.Form_QueryUnload"
    
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

    fraServerInfo.Visible = CheckBoxValue(chkShowIP)

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is getting unloaded
'' Inputs:      Cancel Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.astrIps = Nothing
    Set m.astrPortOverrides = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAccessKey_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAccessKey_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtAccessKey

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin.txtAccessKey_GotFocus"
    
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
    RaiseError "frmPfgLogin.txtPort_GotFocus"
    
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

    Dim lIndex As Long                  ' Index into the array
    
    lIndex = cboServers.ItemData(cboServers.ListIndex)
    If (lIndex = 0) Or (Parse(m.astrIps(lIndex), ":", 3) <> Trim(txtPort.Text)) Then
        m.astrPortOverrides(lIndex) = Trim(txtPort.Text)
        SetIniFileProperty "Port" & Str(lIndex), Trim(txtPort.Text), "Overrides", m.strIniFile
    Else
        m.astrPortOverrides(lIndex) = ""
        SetIniFileProperty "Port" & Str(lIndex), "", "Overrides", m.strIniFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin.txtPort_LostFocus"
    
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
    RaiseError "frmPfgLogin.txtServerIP_GotFocus"
    
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

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Did we find the IP?
    
    bFound = False
    If Len(Trim(txtServerIP.Text)) > 0 Then
        For lIndex = 1 To m.astrIps.Size - 1
            If Parse(m.astrIps(lIndex), ":", 2) = Trim(txtServerIP.Text) Then
                If cboServers.ItemData(cboServers.ListIndex) <> lIndex Then
                    SetServersCombo lIndex
                    bFound = True
                    Exit For
                End If
            End If
        Next lIndex
    End If
    
    If bFound = False Then
        m.astrIps(0) = "(User Defined):" & Trim(txtServerIP.Text) & ":" & Trim(txtPort.Text) & ":"
        SetServersCombo 0&
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin_LostFocus"
    
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
    RaiseError "frmPfgLogin.EnableControls"
    
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
            strAccount = Parse(astrLogins(lIndex), "|", 1)
            If Len(strAccount) > 0 Then
                cboAccount.AddItem strAccount
                cboAccount.ItemData(cboAccount.NewIndex) = CLng(Val(Parse(astrLogins(lIndex), "|", 2)))
            End If
        Next lIndex
    End If
    
    If cboAccount.ListCount = 0 Then
        strAccount = UCase(GetIniFileProperty("UserID", "", "User", m.strIniFile))
        If Len(strAccount) > 0 Then
            cboAccount.AddItem strAccount
            If Left(strAccount, 1) <> "D" Then
                cboAccount.ItemData(cboAccount.NewIndex) = 3
                SetIniFileProperty "Logins", strAccount & "|3", "User", m.strIniFile
            Else
                strIP = GetIniFileProperty("HostIP", "", "User", m.strIniFile)
                If strIP = "12.36.73.70" Then
                    cboAccount.ItemData(cboAccount.NewIndex) = 5
                    SetIniFileProperty "Logins", strAccount & "|5", "User", m.strIniFile
                Else
                    cboAccount.ItemData(cboAccount.NewIndex) = 4
                    SetIniFileProperty "Logins", strAccount & "|4", "User", m.strIniFile
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin.LoadCombo"
    
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
                SetServersCombo cboAccount.ItemData(lIndex)
            End If
        Next lIndex
    End If
    
    SetCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPfgLogin.SetCombo"
    
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
    
    strAccount = UCase(InfBox("What is your " & m.strBrokerName & " account number?", "?", , m.strBrokerName & " Account Number", , , , , , "string"))
    If Len(strAccount) > 0 Then
        If SetCombo(strAccount) = False Then
            If Left(UCase(strAccount), 1) = "D" Then
                strAcctType = InfBox("Is this a Demo or a Contest account?", "?", "+-Demo|Contest", m.strBrokerName & " Account Type")
                If UCase(strAcctType) = "D" Then
                    cboAccount.AddItem strAccount
                    cboAccount.ItemData(cboAccount.NewIndex) = 4
                    strNewLogin = strAccount & "|4"
                Else
                    cboAccount.AddItem strAccount
                    cboAccount.ItemData(cboAccount.NewIndex) = 5
                    strNewLogin = strAccount & "|5"
                End If
            Else
                cboAccount.AddItem strAccount
                cboAccount.ItemData(cboAccount.NewIndex) = 3
                strNewLogin = strAccount & "|3"
            End If
            
            SetCombo strAccount
            MoveFocus txtAccessKey
            
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
    RaiseError "frmPfgLogin.NewAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadServersCombo
'' Description: Load up the servers combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadServersCombo()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lNumIps As Long                 ' Number of IP addresses
    Dim strIP As String                 ' IP information
    
    cboServers.Clear
    cboServers.AddItem "(User Defined)"
    cboServers.ItemData(cboServers.NewIndex) = 0
    
    m.astrIps.Clear
    
    lNumIps = GetIniFileProperty("NumIps", 0&, "IPS", m.strConnectIni)
    For lIndex = 1 To lNumIps
        strIP = GetIniFileProperty("IP" & Str(lIndex), "", "IPS", m.strConnectIni)
        If Len(strIP) > 0 Then
            cboServers.AddItem Parse(strIP, ":", 1)
            cboServers.ItemData(cboServers.NewIndex) = lIndex
            
            m.astrIps.Add strIP, lIndex
        End If
        
        m.astrPortOverrides(lIndex) = GetIniFileProperty("Port" & Str(lIndex), "", "Overrides", m.strIniFile)
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin.LoadServersCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetServersCombo
'' Description: Set the servers combo to the given item
'' Inputs:      Item Data
'' Returns:     Item Found?
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetServersCombo(ByVal lItemData As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    For lIndex = 0 To cboServers.ListCount - 1
        If cboServers.ItemData(lIndex) = lItemData Then
            cboServers.ListIndex = lIndex
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    SetServersCombo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPfgLogin.SetServersCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixLogins
'' Description: Fix logins to the new version of server information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixLogins()
On Error GoTo ErrSection:

    Dim strLogins As String             ' Login string from the INI file
    Dim astrLogins As cGdArray          ' Array of logins split out
    Dim lIndex As Long                  ' Index into a for loop
    Dim strLiveIP As String             ' Live IP address override
    Dim strLiveSystem As String         ' Live system to use
    Dim bFound As Boolean               ' Was the IP address found?
    
    Set astrLogins = New cGdArray
    astrLogins.Create eGDARRAY_Strings
    
    strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
    If Len(strLogins) > 0 Then
        If InStr(strLogins, ";") <> 0 Then
            astrLogins.SplitFields strLogins, ","
            
            strLiveIP = GetIniFileProperty("Live", "", "IPS", m.strIniFile)
            If Len(strLiveIP) > 0 Then
                bFound = False
                For lIndex = 1 To m.astrIps.Size - 1
                    If Parse(strLiveIP, ":", 1) = Parse(m.astrIps(lIndex), ":", 2) Then
                        If Parse(strLiveIP, ":", 2) = Parse(m.astrIps(lIndex), ":", 3) Then
                            strLiveSystem = Str(lIndex)
                            bFound = True
                            Exit For
                        End If
                    End If
                Next lIndex
                
                If bFound = False Then
                    strLiveSystem = "0"
                    SetIniFileProperty "IP0", Parse(strLiveIP, ":", 1), "Overrides", m.strIniFile
                    SetIniFileProperty "Port0", Parse(strLiveIP, ":", 2), "Overrides", m.strIniFile
                    m.astrPortOverrides(0) = Parse(strLiveIP, ":", 2)
                    m.astrIps(0) = "(User Defined):" & strLiveIP & ":"
                Else
                    SetIniFileProperty "IP0", "", "Overrides", m.strIniFile
                    SetIniFileProperty "Port0", "", "Overrides", m.strIniFile
                    m.astrPortOverrides(0) = ""
                    m.astrIps(0) = "(User Defined):::"
                End If
            Else
                strLiveSystem = "3"
            End If
                        
            For lIndex = 0 To astrLogins.Size - 1
                Select Case Parse(astrLogins(lIndex), ";", 2)
                    Case "0"
                        astrLogins(lIndex) = Parse(astrLogins(lIndex), ";", 1) & "|" & strLiveSystem
                    Case "1"
                        astrLogins(lIndex) = Parse(astrLogins(lIndex), ";", 1) & "|4"
                    Case "2"
                        astrLogins(lIndex) = Parse(astrLogins(lIndex), ";", 1) & "|5"
                End Select
            Next lIndex
            
            strLogins = astrLogins.JoinFields(",")
            SetIniFileProperty "Logins", strLogins, "User", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPfgLogin.FixLogins"
    
End Sub
