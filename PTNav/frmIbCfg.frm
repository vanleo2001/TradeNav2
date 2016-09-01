VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmIbCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniTextBoxXP txtEnabledSymbols 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2940
      Width           =   7515
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483630
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmIbCfg.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   2
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmIbCfg.frx":0020
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbCfg.frx":0040
   End
   Begin HexUniControls.ctlUniRichTextBoxXP rtfDisclaimer 
      Height          =   2415
      Left            =   4020
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4260
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmIbCfg.frx":005C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   -1
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   2
      PasswordChar    =   ""
      TrapTab         =   0   'False
      RaiseChangeEvent=   -1  'True
      RaiseUpdateEvent=   0   'False
      RaiseSelChangeEvent=   -1  'True
      Tip             =   "frmIbCfg.frx":007C
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbCfg.frx":009C
      ViewMode        =   0
      TextModeText    =   2
      TextModeUndoLevel=   8
      TextModeCodePage=   32
      AutoURLDetect   =   0   'False
      FileName        =   ""
      VerticalLayout  =   0   'False
      OnlyNumbers     =   0   'False
      NoIME           =   0   'False
      SelfIME         =   0   'False
      LanguageOptions =   150
      RaiseRequestResizeEvent=   0   'False
      RaiseMsgFilterEvent=   0   'False
      SubClassPaintMessage=   0   'False
      TabSize         =   4
      TypographyOptions=   0
      BlockAutoCopy   =   0   'False
      BlockAutoCut    =   0   'False
      BlockAutoPaste  =   0   'False
      BlockAutoUndo   =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   750
      TabIndex        =   3
      Top             =   2040
      Width           =   2535
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmIbCfg.frx":00B8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmIbCfg.frx":00E4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbCfg.frx":0104
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Top             =   0
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmIbCfg.frx":0120
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmIbCfg.frx":014E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmIbCfg.frx":016E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmIbCfg.frx":018A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmIbCfg.frx":01B0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmIbCfg.frx":01D0
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txtPort 
      Height          =   315
      Left            =   3180
      TabIndex        =   6
      Top             =   1020
      Width           =   675
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmIbCfg.frx":01EC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   0   'False
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmIbCfg.frx":020C
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbCfg.frx":022C
   End
   Begin HexUniControls.ctlUniTextBoxXP txtIP 
      Height          =   315
      Left            =   1140
      TabIndex        =   4
      Top             =   1020
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmIbCfg.frx":0248
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   0   'False
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmIbCfg.frx":0268
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbCfg.frx":0288
   End
   Begin HexUniControls.ctlUniTextBoxXP txtClientID 
      Height          =   255
      Left            =   1740
      TabIndex        =   1
      Top             =   180
      Width           =   2115
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmIbCfg.frx":02A4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   0   'False
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmIbCfg.frx":02C4
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbCfg.frx":02E4
   End
   Begin HexUniControls.ctlUniLabelXP lblEnabledSymbols 
      Height          =   195
      Left            =   120
      Top             =   2700
      Width           =   2475
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmIbCfg.frx":0300
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmIbCfg.frx":0340
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbCfg.frx":0360
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAgree 
      Height          =   435
      Left            =   180
      Top             =   1500
      Width           =   3675
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmIbCfg.frx":037C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmIbCfg.frx":043C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbCfg.frx":045C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblPort 
      Height          =   255
      Left            =   2700
      Top             =   1050
      Width           =   435
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmIbCfg.frx":0478
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmIbCfg.frx":04A4
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbCfg.frx":04C4
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblInfo 
      Height          =   435
      Left            =   180
      Top             =   600
      Width           =   3615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmIbCfg.frx":04E0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmIbCfg.frx":059C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbCfg.frx":05BC
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblHostIP 
      Height          =   255
      Left            =   180
      Top             =   1050
      Width           =   915
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmIbCfg.frx":05D8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmIbCfg.frx":0610
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbCfg.frx":0630
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblClientID 
      Height          =   255
      Left            =   180
      Top             =   180
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmIbCfg.frx":064C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmIbCfg.frx":069A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbCfg.frx":06BA
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmIbCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmIbCfg.frm
'' Description: Allows the user to enter Interactive Brokers configuration info
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 02/06/2012   DAJ         Change name for 'I-Deal' to 'TWS Pro'
'' 08/13/2014   DAJ         Added enabled symbols
'' 12/22/2014   DAJ         Remove Forex symbols out of a copy of enabled symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
    
    nBroker As eTT_AccountType          ' Broker
    strBrokerName As String             ' Broker name
    strIniFile As String                ' INI file
    
    strDefaultClient As String          ' Default Client ID
    strDefaultIP As String              ' Default IP address
    strDefaultPort As String            ' Default Port
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Broker
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal Broker As cBroker) As Boolean
On Error GoTo ErrSection:

    Dim strClient As String             ' Client override from INI file
    Dim strIP As String                 ' IP override from INI file
    Dim strPort As String               ' Port override from INI file
    Dim astrEnabledSymbols As cGdArray  ' Enabled symbols array
    Dim lIndex As Long                  ' Index into a for loop
    
    m.nBroker = Broker.Broker
    m.strBrokerName = Broker.BrokerName
    m.strIniFile = Broker.IniFile
    Caption = m.strBrokerName & " Configuration"

    m.strDefaultClient = GetIniFileProperty("ID", "0", "Client", AddSlash(App.Path) & "Provided\IbIps.INI")
    m.strDefaultIP = GetIniFileProperty("Live", "", "IP", AddSlash(App.Path) & "Provided\IbIps.INI")
    m.strDefaultPort = GetIniFileProperty("Live", "", "Port", AddSlash(App.Path) & "Provided\IbIps.INI")
    
    strClient = GetIniFileProperty("Client", "", "LastLogin", m.strIniFile)
    strIP = GetIniFileProperty("IP", "", "LastLogin", m.strIniFile)
    strPort = GetIniFileProperty("Port", "", "LastLogin", m.strIniFile)
    
    If Len(strClient) = 0 Then
        txtClientID.Text = m.strDefaultClient
    Else
        txtClientID.Text = strClient
    End If
    
    If Len(strIP) = 0 Then
        txtIP.Text = m.strDefaultIP
    Else
        txtIP.Text = strIP
    End If
    
    If Len(strPort) = 0 Then
        txtPort.Text = m.strDefaultPort
    Else
        txtPort.Text = strPort
    End If
    
    If Broker.EnabledSymbols Is Nothing Then
        txtEnabledSymbols.Text = ""
    ElseIf Broker.EnabledSymbols.Size = 0 Then
        txtEnabledSymbols.Text = "None"
    Else
        Set astrEnabledSymbols = Broker.EnabledSymbols.MakeCopy
        
        For lIndex = astrEnabledSymbols.Size - 1 To 0 Step -1
            If (astrEnabledSymbols(lIndex) = "!") Or (InStr(astrEnabledSymbols(lIndex), "@") > 0) Or (InStr(astrEnabledSymbols(lIndex), "O:") > 0) Then
                astrEnabledSymbols.Remove lIndex
            End If
        Next lIndex
        
        txtEnabledSymbols.Text = astrEnabledSymbols.JoinFields(", ")
    End If

    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK Then
        If Trim(txtClientID.Text) = m.strDefaultClient Then
            SetIniFileProperty "Client", "", "LastLogin", m.strIniFile
        Else
            SetIniFileProperty "Client", Trim(txtClientID.Text), "LastLogin", m.strIniFile
        End If
        
        If Trim(txtIP.Text) = m.strDefaultIP Then
            SetIniFileProperty "IP", "", "LastLogin", m.strIniFile
        Else
            SetIniFileProperty "IP", Trim(txtIP.Text), "LastLogin", m.strIniFile
        End If
        
        If Trim(txtPort.Text) = m.strDefaultPort Then
            SetIniFileProperty "Port", "", "LastLogin", m.strIniFile
        Else
            SetIniFileProperty "Port", Trim(txtPort.Text), "LastLogin", m.strIniFile
        End If
        SetIniFileProperty "Asked", 1&, "User", m.strIniFile
        
        Select Case m.nBroker
            Case eTT_AccountType_Ideal
                g.Ideal.ClientID = Trim(txtClientID.Text)
                g.Ideal.HostIP = Trim(txtIP.Text)
                g.Ideal.HostPort = Trim(txtPort.Text)
            
            Case eTT_AccountType_IntBrokers
                g.IntBroker.ClientID = Trim(txtClientID.Text)
                g.IntBroker.HostIP = Trim(txtIP.Text)
                g.IntBroker.HostPort = Trim(txtPort.Text)
        End Select
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmIbCfg.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the ShowMe to unload the form but do not save changes
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
    RaiseError "frmIbCfg.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Allow the ShowMe to unload the form but save changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If Len(Trim(txtClientID.Text)) = 0 Then
        MoveFocus txtClientID
        InfBox "You must enter in a Client ID", "!", , m.strBrokerName & " Configuration Error"
        Exit Sub
    End If

    If IsAlpha(Trim(txtClientID.Text)) = True Then
        MoveFocus txtClientID
        InfBox "Client ID must be a number", "!", , m.strBrokerName & " Configuration Error"
        Exit Sub
    End If

    If Len(Trim(txtPort.Text)) = 0 Then
        MoveFocus txtPort
        InfBox "You must enter in a connection port", "!", , m.strBrokerName & " Configuration Error"
        Exit Sub
    End If

    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIbCfg.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form upon loading
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    g.Styler.StyleForm Me
    
    Icon = Picture16("kBlank")
    
    rtfDisclaimer.Locked = True
    rtfDisclaimer.Font.Name = "Arial"
    rtfDisclaimer.BackColor = vbButtonFace
    rtfDisclaimer.TextRTF = FileToString(AddSlash(App.Path) & "Info\BrokerDisclaimer.rtf")
    
    txtEnabledSymbols.Enabled = True
    txtEnabledSymbols.BackColor = cmdCancel.BackColor
    txtEnabledSymbols.Locked = True

    CenterTheForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIbCfg.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Allow the ShowMe routine to unload the form
'' Inputs:      Whether to Cancel Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.bOK = False
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIbCfg.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtClientID_GotFocus
'' Description: When the control gets the focus, highlight all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtClientID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtClientID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIbCfg.txtClientID_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtIP_GotFocus
'' Description: When the control gets the focus, highlight all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtIP_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIbCfg.txtIP_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPort_GotFocus
'' Description: When the control gets the focus, highlight all of the text
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
    RaiseError "frmIbCfg.txtPort_GotFocus"
    
End Sub

