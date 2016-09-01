VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmBrokerLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraLogin 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3735
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   0
         Width           =   1395
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1020
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   420
         Width           =   2655
      End
      Begin VB.CommandButton cmdNewAccount 
         Caption         =   "Add Logi&n"
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   0
         Width           =   1155
      End
      Begin VB.Label lblAccount 
         Caption         =   "&User Name:"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   60
         Width           =   915
      End
      Begin VB.Label lblPassword 
         Caption         =   "&Password:"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.PictureBox picBroker 
      Height          =   1155
      Left            =   120
      Picture         =   "frmBrokerLogin.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin RichTextLib.RichTextBox rtfDisclaimer 
      Height          =   4695
      Left            =   4020
      TabIndex        =   27
      Top             =   120
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   8281
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmBrokerLogin.frx":0CD7
   End
   Begin VB.Frame fraServerInfo 
      Caption         =   "Server Information"
      Height          =   1395
      Left            =   120
      TabIndex        =   12
      Top             =   3420
      Width           =   3735
      Begin VB.Frame fraPats 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   3495
         Begin VB.CheckBox chkSuperTAS 
            Caption         =   "Super &TAS"
            Height          =   195
            Left            =   2340
            TabIndex        =   26
            Top             =   60
            Width           =   1095
         End
         Begin VB.ComboBox cboEnvironment 
            Height          =   315
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblEnvironment 
            Caption         =   "En&vironment:"
            Height          =   225
            Left            =   0
            TabIndex        =   24
            Top             =   45
            Width           =   1035
         End
      End
      Begin VB.Frame fraPrice 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   3495
         Begin VB.TextBox txtPriceIP 
            Height          =   285
            Left            =   660
            TabIndex        =   20
            Top             =   0
            Width           =   1575
         End
         Begin VB.TextBox txtPricePort 
            Height          =   285
            Left            =   2760
            TabIndex        =   22
            Top             =   0
            Width           =   675
         End
         Begin VB.Label lblPriceIP 
            Caption         =   "Pric&e IP:"
            Height          =   195
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblPricePort 
            Caption         =   "P&ort:"
            Height          =   195
            Left            =   2340
            TabIndex        =   21
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.Frame fraHost 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3495
         Begin VB.TextBox txtPort 
            Height          =   285
            Left            =   2760
            TabIndex        =   17
            Top             =   0
            Width           =   675
         End
         Begin VB.TextBox txtServerIP 
            Height          =   285
            Left            =   660
            TabIndex        =   15
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label lblPort 
            Caption         =   "Po&rt:"
            Height          =   195
            Left            =   2340
            TabIndex        =   16
            Top             =   0
            Width           =   315
         End
         Begin VB.Label lblServerIP 
            Caption         =   "Host &IP:"
            Height          =   195
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   615
         End
      End
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   915
      Left            =   120
      TabIndex        =   7
      Top             =   2340
      Width           =   3735
      Begin VB.CheckBox chkShowIP 
         Caption         =   "&Show Server Information"
         Height          =   435
         Left            =   2220
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   1080
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "&Login"
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblAgree 
         Caption         =   "Choosing to login states that you agree to the terms and conditions on the right"
         Height          =   435
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   3675
      End
   End
End
Attribute VB_Name = "frmBrokerLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmBrokerLogin.frm
'' Description: Allow the user to choose their login information
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO 80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/03/2012   DAJ         Remove Alaron, Cadent, Lotus, LindWaldock, ManExpress, Robbins
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    nBroker As eTT_AccountType
    strBroker As String
    
    strIniFile As String
    strDefaultIP As String
    strDefaultPort As String
    strDefaultPriceIP As String
    strDefaultPricePort As String
    strDefaultFirm As String
    strDefaultSubsystem As String
    strDefaultEnvironment As String
    strDefaultSuperTAS As String
    
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
Public Function ShowMe(ByVal nBroker As eTT_AccountType, Optional ByVal strAccount As String = "", Optional ByVal bShowIP As Boolean = False, Optional ByVal bSwitching As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strLastAccount As String        ' Last account logged into
    Dim strLogins As String             ' Logins from the INI file
    Dim astrLogins As New cGdArray      ' Array of logins from the INI file
    Dim lIndex As Long                  ' Index into a for loop
    Dim strIP As String                 ' IP override
    Dim strPriceIP As String            ' Price IP override
    Dim strFirm As String               ' Firm override
    Dim strPats As String               ' Extra Pats connection information

    m.bOK = False
    
    m.nBroker = nBroker
    m.strBroker = g.Broker.BrokerName(m.nBroker)
    Caption = m.strBroker & " Login Information"
    
    Select Case nBroker
        Case eTT_AccountType_PATS
            m.strIniFile = AddSlash(App.Path) & "Pats3.INI"
            m.strDefaultIP = GetIniFileProperty("Live", "", "IP", AddSlash(App.Path) & "Provided\PatsIps.INI")
            m.strDefaultPort = GetIniFileProperty("Live", "", "Port", AddSlash(App.Path) & "Provided\PatsIps.INI")
            m.strDefaultPriceIP = GetIniFileProperty("Price", "", "IP", AddSlash(App.Path) & "Provided\PatsIps.INI")
            m.strDefaultPricePort = GetIniFileProperty("Price", "", "Port", AddSlash(App.Path) & "Provided\PatsIps.INI")
            m.strDefaultEnvironment = GetIniFileProperty("Environment", "", "Misc", AddSlash(App.Path) & "Provided\PatsIps.INI")
            m.strDefaultSuperTAS = GetIniFileProperty("SuperTAS", "", "Misc", AddSlash(App.Path) & "Provided\PatsIps.INI")
            strIP = GetIniFileProperty("Live", "", "IPS", m.strIniFile)
            strPriceIP = GetIniFileProperty("Price", "", "IPS", m.strIniFile)
            strPats = GetIniFileProperty("Misc", "", "IPS", m.strIniFile)
            
            fraPrice.Visible = True
            fraPats.Top = 960
            fraPats.Visible = True
            fraServerInfo.Height = 1395
            
            picBroker.Visible = False
            fraLogin.Top = 120
            fraButtons.Top = 1020
            fraServerInfo.Top = 2100
            Height = 4530
            
        Case eTT_AccountType_Rosenthal
            m.strIniFile = AddSlash(App.Path) & "Rosenthal.INI"
            m.strDefaultIP = GetIniFileProperty("Live", "", "IP", AddSlash(App.Path) & "Provided\RoseIps.INI")
            m.strDefaultPort = GetIniFileProperty("Live", "", "Port", AddSlash(App.Path) & "Provided\RoseIps.INI")
            m.strDefaultPriceIP = GetIniFileProperty("Price", "", "IP", AddSlash(App.Path) & "Provided\RoseIps.INI")
            m.strDefaultPricePort = GetIniFileProperty("Price", "", "Port", AddSlash(App.Path) & "Provided\RoseIps.INI")
            m.strDefaultEnvironment = GetIniFileProperty("Environment", "", "Misc", AddSlash(App.Path) & "Provided\RoseIps.INI")
            m.strDefaultSuperTAS = GetIniFileProperty("SuperTAS", "", "Misc", AddSlash(App.Path) & "Provided\RoseIps.INI")
            strIP = GetIniFileProperty("Live", "", "IPS", m.strIniFile)
            strPriceIP = GetIniFileProperty("Price", "", "IPS", m.strIniFile)
            strPats = GetIniFileProperty("Misc", "", "IPS", m.strIniFile)
            
            fraPrice.Visible = True
            fraPats.Top = 960
            fraPats.Visible = True
            fraServerInfo.Height = 1395
            
            picBroker.Visible = False
            fraLogin.Top = 120
            fraButtons.Top = 1020
            fraServerInfo.Top = 2100
            Height = 4530
            
    End Select
    
    If Len(strIP) > 0 Then
        txtServerIP.Text = Parse(strIP, ":", 1)
        txtPort.Text = Parse(strIP, ":", 2)
    Else
        txtServerIP.Text = m.strDefaultIP
        txtPort.Text = m.strDefaultPort
    End If
    
    If Len(strPriceIP) > 0 Then
        txtPriceIP.Text = Parse(strPriceIP, ":", 1)
        txtPricePort.Text = Parse(strPriceIP, ":", 2)
    Else
        txtPriceIP.Text = m.strDefaultPriceIP
        txtPricePort.Text = m.strDefaultPricePort
    End If
    
    If Len(strPats) > 0 Then
        SetEnvironmentCombo Parse(strPats, ":", 1)
        If Parse(strPats, ":", 2) = "Y" Then chkSuperTAS.Value = vbChecked Else chkSuperTAS.Value = vbUnchecked
    Else
        SetEnvironmentCombo m.strDefaultEnvironment
        If m.strDefaultSuperTAS = "Y" Then chkSuperTAS.Value = vbChecked Else chkSuperTAS.Value = vbUnchecked
    End If
    
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
        
        Form_Resize
        ShowForm Me, eForm_Modal, frmMain
        
        If m.bOK = True Then
            Select Case nBroker
                Case eTT_AccountType_PATS
                    If Not g.Pats Is Nothing Then
                        g.Pats.UserName = cboAccount.Text
                        g.Pats.Password = Trim(txtPassword.Text)
                        g.Pats.HostIP = Trim(txtServerIP.Text)
                        g.Pats.HostPort = Trim(txtPort.Text)
                        g.Pats.PriceIP = Trim(txtPriceIP.Text)
                        g.Pats.PricePort = Trim(txtPricePort.Text)
                        g.Pats.Environment = cboEnvironment.Text
                        If chkSuperTAS.Value = vbChecked Then g.Pats.SuperTAS = "Y" Else g.Pats.SuperTAS = "N"
                    
                        SetIniFileProperty "LastAccount", cboAccount.Text, "User", m.strIniFile
                        strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
                        astrLogins.SplitFields strLogins, ","
                    
                        If (Trim(txtServerIP.Text) <> m.strDefaultIP) Or (Trim(txtPort.Text) <> m.strDefaultPort) Then
                            SetIniFileProperty "Live", Trim(txtServerIP.Text) & ":" & Trim(txtPort.Text), "IPS", m.strIniFile
                        Else
                            SetIniFileProperty "Live", "", "IPS", m.strIniFile
                        End If
                        
                        If (Trim(txtPriceIP.Text) <> m.strDefaultPriceIP) Or (Trim(txtPricePort.Text) <> m.strDefaultPricePort) Then
                            SetIniFileProperty "Price", Trim(txtPriceIP.Text) & ":" & Trim(txtPricePort.Text), "IPS", m.strIniFile
                        Else
                            SetIniFileProperty "Price", "", "IPS", m.strIniFile
                        End If
                        
                        If (cboEnvironment.Text <> m.strDefaultEnvironment) Or (g.Pats.SuperTAS <> m.strDefaultSuperTAS) Then
                            SetIniFileProperty "Misc", cboEnvironment.Text & ":" & g.Pats.SuperTAS, "IPS", m.strIniFile
                        Else
                            SetIniFileProperty "Misc", "", "IPS", m.strIniFile
                        End If
                        
                        SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile
                    End If
                    
                Case eTT_AccountType_Rosenthal
                    If Not g.Rosenthal Is Nothing Then
                        g.Rosenthal.UserName = cboAccount.Text
                        g.Rosenthal.Password = Trim(txtPassword.Text)
                        g.Rosenthal.HostIP = Trim(txtServerIP.Text)
                        g.Rosenthal.HostPort = Trim(txtPort.Text)
                        g.Rosenthal.PriceIP = Trim(txtPriceIP.Text)
                        g.Rosenthal.PricePort = Trim(txtPricePort.Text)
                        g.Rosenthal.PricePort = Trim(txtPricePort.Text)
                        g.Rosenthal.Environment = cboEnvironment.Text
                        If chkSuperTAS.Value = vbChecked Then g.Rosenthal.SuperTAS = "Y" Else g.Rosenthal.SuperTAS = "N"
                    
                        SetIniFileProperty "LastAccount", cboAccount.Text, "User", m.strIniFile
                        strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
                        astrLogins.SplitFields strLogins, ","
                    
                        If (Trim(txtServerIP.Text) <> m.strDefaultIP) Or (Trim(txtPort.Text) <> m.strDefaultPort) Then
                            SetIniFileProperty "Live", Trim(txtServerIP.Text) & ":" & Trim(txtPort.Text), "IPS", m.strIniFile
                        Else
                            SetIniFileProperty "Live", "", "IPS", m.strIniFile
                        End If
                        
                        If (Trim(txtPriceIP.Text) <> m.strDefaultPriceIP) Or (Trim(txtPricePort.Text) <> m.strDefaultPricePort) Then
                            SetIniFileProperty "Price", Trim(txtPriceIP.Text) & ":" & Trim(txtPricePort.Text), "IPS", m.strIniFile
                        Else
                            SetIniFileProperty "Price", "", "IPS", m.strIniFile
                        End If
                        
                        If (cboEnvironment.Text <> m.strDefaultEnvironment) Or (g.Rosenthal.SuperTAS <> m.strDefaultSuperTAS) Then
                            SetIniFileProperty "Misc", cboEnvironment.Text & ":" & g.Rosenthal.SuperTAS, "IPS", m.strIniFile
                        Else
                            SetIniFileProperty "Misc", "", "IPS", m.strIniFile
                        End If
                        
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
    RaiseError "frmBrokerLogin.ShowMe"
    
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

    MoveFocus txtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerLogin.cboAccount_Click"
    
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
    RaiseError "frmBrokerLogin.chkShowIP_Click"
    
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
    RaiseError "frmBrokerLogin.cmdCancel_Click"
    
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
        InfBox "Please enter in an IP address for " & m.strBroker & " server", "!", , m.strBroker & " Login Error"
        GoTo ErrExit
    End If

    If Len(Trim(txtPort.Text)) = 0 Then
        chkShowIP.Value = vbChecked
        MoveFocus txtPort
        InfBox "Please enter in a port for " & m.strBroker & " server", "!", , m.strBroker & " Login Error"
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
    RaiseError "frmBrokerLogin.cmdLogin_Click"
    
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
    RaiseError "frmBrokerLogin.cmdNewAccount_Click"
    
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
    RaiseError "frmBrokerLogin.Form_Activate"
    
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

    With cboEnvironment
        .AddItem "Gateway"
        .AddItem "Client"
        .AddItem "Test Client"
        .AddItem "Test Gateway"
        .AddItem "Demo Client"
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerLogin.Form_Load"
    
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
    RaiseError "frmBrokerLogin.Form_QueryUnload"
    
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

    Dim dDiff As Double                 ' Difference between the height and the scale height
    
    dDiff = Height - ScaleHeight
    Height = fraServerInfo.Top + fraServerInfo.Height + 120 + dDiff
    
    With rtfDisclaimer
        .Move .Left, .Top, .Width, ScaleHeight - (.Top * 2)
    End With
    
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
    RaiseError "frmBrokerLogin.txtPassword_GotFocus"
    
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
    RaiseError "frmBrokerLogin.txtPort_GotFocus"
    
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
    RaiseError "frmBrokerLogin.txtServerIP_GotFocus"
    
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
    RaiseError "frmBrokerLogin.EnableControls"
    
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
    
    If Len(strLogins) > 0 Then
        astrLogins.SplitFields strLogins, ","
        For lIndex = 0 To astrLogins.Size - 1
            cboAccount.AddItem astrLogins(lIndex)
        Next lIndex
    End If
    
    If cboAccount.ListCount = 0 Then
        strAccount = UCase(GetIniFileProperty("UserID", "", "User", m.strIniFile))
        If Len(strAccount) > 0 Then
            cboAccount.AddItem strAccount
            SetIniFileProperty "Logins", strAccount, "User", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerLogin.LoadCombo"
    
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
            End If
        Next lIndex
    End If
    
    SetCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerLogin.SetCombo"
    
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
    
    strAccount = UCase(InfBox("What is your " & m.strBroker & " user name?", "?", , m.strBroker & " Login", , , , , , "string"))
    If Len(strAccount) > 0 Then
        If SetCombo(strAccount) = False Then
            cboAccount.AddItem strAccount
            strNewLogin = strAccount
            
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
    RaiseError "frmBrokerLogin.NewAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetEnvironmentCombo
'' Description: Set the environment combo box to the given environment if possible
'' Inputs:      Environment
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetEnvironmentCombo(ByVal strEnvironment As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the item found?
    
    bFound = False
    If (cboEnvironment.ListCount > 0) And (Len(strEnvironment) > 0) Then
        For lIndex = 0 To cboEnvironment.ListCount - 1
            If UCase(cboEnvironment.List(lIndex)) = UCase(strEnvironment) Then
                bFound = True
                cboEnvironment.ListIndex = lIndex
            End If
        Next lIndex
    End If
    
    If bFound = False Then cboEnvironment.ListIndex = 0
    SetEnvironmentCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerLogin.SetEnvironmentCombo"
    
End Function
