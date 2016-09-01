VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmLoginBrokerUrl 
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraLoginInfo 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
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
      Caption         =   "frmLoginBrokerUrl.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginBrokerUrl.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginBrokerUrl.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboUserName 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   0
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   -2147483630
         ButtonStyle     =   -1
         SelectorStyle   =   -1
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
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
         Tip             =   "frmLoginBrokerUrl.frx":0068
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmLoginBrokerUrl.frx":0088
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPassword 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmLoginBrokerUrl.frx":00A4
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
         PasswordChar    =   "*"
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmLoginBrokerUrl.frx":00C4
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginBrokerUrl.frx":00E4
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAddLogin 
         Height          =   315
         Left            =   2940
         TabIndex        =   3
         Top             =   0
         Width           =   375
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
         Caption         =   "frmLoginBrokerUrl.frx":0100
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginBrokerUrl.frx":0136
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginBrokerUrl.frx":0156
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemoveLogin 
         Height          =   315
         Left            =   3360
         TabIndex        =   4
         Top             =   0
         Width           =   375
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
         Caption         =   "frmLoginBrokerUrl.frx":0172
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginBrokerUrl.frx":01AE
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginBrokerUrl.frx":01CE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblUserName 
         Height          =   255
         Left            =   0
         Top             =   60
         Width           =   975
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
         Caption         =   "frmLoginBrokerUrl.frx":01EA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginBrokerUrl.frx":0220
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginBrokerUrl.frx":0240
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPassword 
         Height          =   255
         Left            =   0
         Top             =   420
         Width           =   975
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
         Caption         =   "frmLoginBrokerUrl.frx":025C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginBrokerUrl.frx":0290
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginBrokerUrl.frx":02B0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraServerInfo 
      Height          =   1395
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   3735
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
      Caption         =   "frmLoginBrokerUrl.frx":02CC
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginBrokerUrl.frx":0310
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginBrokerUrl.frx":0330
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraHost 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   3495
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
         Caption         =   "frmLoginBrokerUrl.frx":034C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmLoginBrokerUrl.frx":0378
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginBrokerUrl.frx":0398
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtUrl 
            Height          =   285
            Left            =   480
            TabIndex        =   5
            Top             =   0
            Width           =   2955
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLoginBrokerUrl.frx":03B4
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
            Tip             =   "frmLoginBrokerUrl.frx":03D4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginBrokerUrl.frx":03F4
         End
         Begin HexUniControls.ctlUniLabelXP lblUrl 
            Height          =   195
            Left            =   0
            Top             =   45
            Width           =   615
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
            Caption         =   "frmLoginBrokerUrl.frx":0410
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmLoginBrokerUrl.frx":043A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginBrokerUrl.frx":045A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   915
      Left            =   120
      TabIndex        =   7
      Top             =   900
      Width           =   3735
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
      Caption         =   "frmLoginBrokerUrl.frx":0476
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginBrokerUrl.frx":04A2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginBrokerUrl.frx":04C2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkShowIP 
         Height          =   435
         Left            =   2460
         TabIndex        =   11
         Top             =   480
         Width           =   1215
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
         Caption         =   "frmLoginBrokerUrl.frx":04DE
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmLoginBrokerUrl.frx":052E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmLoginBrokerUrl.frx":054E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   1080
         TabIndex        =   10
         Top             =   480
         Width           =   975
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
         Caption         =   "frmLoginBrokerUrl.frx":056A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginBrokerUrl.frx":0598
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginBrokerUrl.frx":05B8
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdLogin 
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   9
         Top             =   480
         Width           =   975
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
         Caption         =   "frmLoginBrokerUrl.frx":05D4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginBrokerUrl.frx":0600
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginBrokerUrl.frx":0620
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAgree 
         Height          =   435
         Left            =   0
         Top             =   0
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
         Caption         =   "frmLoginBrokerUrl.frx":063C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginBrokerUrl.frx":06FC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginBrokerUrl.frx":071C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniRichTextBoxXP rtfDisclaimer 
      Height          =   3195
      Left            =   4020
      TabIndex        =   8
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5636
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmLoginBrokerUrl.frx":0738
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
      Tip             =   "frmLoginBrokerUrl.frx":0758
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginBrokerUrl.frx":0778
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
End
Attribute VB_Name = "frmLoginBrokerUrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLoginBrokerUrl.frm
'' Description: Allow the user to enter login information for a broker that
''              connects to a URL
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user press OK or Cancel?
    strIniFile As String                ' INI file for the broker
    strConnectIni As String             ' INI file for connection information
    strBrokerName As String             ' Display name for the broker
    
    strUserName As String               ' User name that the user chose
    strPassword As String               ' Password from the user
    strUrl As String                    ' URL to connect to
End Type
Private m As mPrivate

Public Property Get UserName() As String
    UserName = m.strUserName
End Property

Public Property Get Password() As String
    Password = m.strPassword
End Property

Public Property Get Url() As String
    Url = m.strUrl
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Broker, UserID, Are we switching?, Show IP?
'' Returns:     True if login, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal Broker As cBroker, Optional ByVal strUserName As String = "", Optional ByVal bSwitching As Boolean = False, Optional ByVal bShowIP As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strLastUserName As String       ' Last user name logged into

    m.bOK = False
    m.strIniFile = Broker.IniFile
    m.strConnectIni = Broker.ConnectIni
    m.strBrokerName = Broker.BrokerName
    Caption = m.strBrokerName & " Login Information"
    
    If Len(m.strIniFile) > 0 Then
        strLastUserName = GetIniFileProperty("LastUserName", "", "User", m.strIniFile)
        LoadCombo
        If cboUserName.ListCount > 0 Then
            If SetCombo(strUserName) = False Then
                If SetCombo(strLastUserName) = False Then
                    strLastUserName = GetIniFileProperty("UserName", "", "User", m.strIniFile)
                    If SetCombo(strLastUserName) = False Then
                        cboUserName.ListIndex = 0
                    End If
                End If
            End If
        End If
        
        If (cboUserName.ListCount = 0) Or ((cboUserName.ListCount = 1) And (bSwitching = True)) Then
            AddLogin
        End If
        
        SetServerControls
        
        If (cboUserName.ListCount > 1) Or ((cboUserName.ListCount = 1) And (bSwitching = False)) Then
            CheckBoxValue(chkShowIP) = bShowIP
            fraServerInfo.Visible = bShowIP
            
            With rtfDisclaimer
                .Move .Left, .Top, .Width, ScaleHeight - (.Top * 2)
            End With
                        
            MoveFocus txtPassword
    
            ShowForm Me, eForm_Modal, frmMain
            
            If m.bOK = True Then
                SetServerOverrides
                
                m.strUserName = cboUserName.Text
                m.strPassword = Trim(txtPassword.Text)
                m.strUrl = Trim(txtUrl.Text)
                
                SetIniFileProperty "LastUserName", cboUserName.Text, "User", m.strIniFile
            End If
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmLoginBrokerUrl.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkShowIP_Click
'' Description: Show/Hide the server information as appropriate
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
    RaiseError "frmLoginBrokerUrl.chkShowIP_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddLogin_Click
'' Description: Allow the user to add a login
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddLogin_Click()
On Error GoTo ErrSection:

    AddLogin

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginBrokerUrl.cmdAddLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Close the dialog without logging in
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
    RaiseError "frmLoginBrokerUrl.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLogin_Click
'' Description: Verify the user information and pass back to Rithmic object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLogin_Click()
On Error GoTo ErrSection:

    If cboUserName.ListIndex < 0 Then
        MoveFocus cboUserName
        InfBox "Please enter in a User Name", "!", , "Login Error"
    ElseIf Len(Trim(txtPassword.Text)) = 0 Then
        MoveFocus txtPassword
        InfBox "Please enter in a Password", "!", , "Login Error"
    ElseIf Len(Trim(txtUrl.Text)) = 0 Then
        CheckBoxValue(chkShowIP) = True
        MoveFocus txtUrl
        InfBox "Please enter in an URL for the server", "!", , "Login Error"
    Else
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginBrokerUrl.cmdLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemoveLogin_Click
'' Description: Allow the user to remove one or more logins
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemoveLogin_Click()
On Error GoTo ErrSection:

    RemoveLogin

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginBrokerUrl.cmdRemoveLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Make sure when the form is activated that password gets focus
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
    RaiseError "frmLoginBrokerUrl.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Do some initialization when the form is loaded
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
    
    g.Styler.StyleForm Me
    
    cmdAddLogin.ToolTipText = "Add a user name"
    cmdRemoveLogin.ToolTipText = "Remove user name(s)"

    m.strUserName = ""
    m.strPassword = ""
    m.strUrl = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginBrokerUrl.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hit the X, let ShowMe unload the form
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
    RaiseError "frmLoginBrokerUrl.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtUrl_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtUrl_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtUrl

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginBrokerUrl.txtUrl.GotFocus"
    
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
    Dim strUserName As String           ' User Name already in the INI file
    Dim strIP As String                 ' IP address from the INI file
    
    cboUserName.Clear
    strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
    If Len(strLogins) > 0 Then
        astrLogins.SplitFields strLogins, ","
        For lIndex = 0 To astrLogins.Size - 1
            If Len(astrLogins(lIndex)) > 0 Then
                cboUserName.AddItem astrLogins(lIndex)
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginBrokerUrl.LoadCombo"
    
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
    RaiseError "frmLoginBrokerUrl.SetCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddLogin
'' Description: Allow the user to give us a new user name
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddLogin()
On Error GoTo ErrSection:

    Dim strUserName As String           ' User name from the user
    Dim strNewLogin As String           ' New login to save to INI file
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
    RaiseError "frmLoginBrokerUrl.AddLogin"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveLogin
'' Description: Allow the user to remove login information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveLogin()
On Error GoTo ErrSection:

    Dim strLogins As String             ' Login information string from INI file
    Dim astrLogins As cGdArray          ' Array of login information
    Dim astrList As cGdArray            ' List to send to the delete form
    Dim astrToDelete As cGdArray        ' List of logins to delete
    Dim strSelected As String           ' Currently selected login
    Dim lIndex As Long                  ' Index into a for loop

    strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
    If Len(strLogins) > 0 Then
        Set astrLogins = New cGdArray
        Set astrList = New cGdArray
        astrList.Create eGDARRAY_Strings
        
        astrLogins.SplitFields strLogins, ","
        For lIndex = 0 To astrLogins.Size - 1
            astrList.Add astrLogins(lIndex) & vbTab & Str(lIndex)
        Next lIndex
        
        strSelected = cboUserName.Text
        
        Set astrToDelete = frmDelete.ShowMe(astrList, strSelected)
        If Not astrToDelete Is Nothing Then
            For lIndex = astrToDelete.Size - 1 To 0 Step -1
                astrLogins.Remove CLng(Val(astrToDelete(lIndex)))
            Next lIndex
            
            SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile
            
            LoadCombo
            If SetCombo(strSelected) = False Then
                If cboUserName.ListCount > 0 Then
                    cboUserName.ListIndex = 0
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginBrokerUrl.RemoveLogin"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetServerControls
'' Description: Set the server controls from the INI files
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetServerControls()
On Error GoTo ErrSection:

    Dim strUrl As String                ' URL to override

    txtUrl.Text = GetIniFileProperty("URL", "", "Server", m.strConnectIni)
    
    strUrl = GetIniFileProperty("URL", "", "Override", m.strIniFile)
    If Len(strUrl) > 0 Then
        If strUrl = txtUrl.Text Then
            SetIniFileProperty "URL", "", "Override", m.strIniFile
        Else
            txtUrl.Text = strUrl
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginBrokerUrl.SetServerControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetServerOverrides
'' Description: Set the server overrides in the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetServerOverrides()
On Error GoTo ErrSection:

    Dim strUrl As String                ' URL to override

    strUrl = GetIniFileProperty("URL", "", "Server", m.strConnectIni)
    
    If Len(Trim(txtUrl.Text)) > 0 Then
        If Trim(txtUrl.Text) = strUrl Then
            SetIniFileProperty "URL", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "URL", Trim(txtUrl.Text), "Override", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginBrokerUrl.SetServerOverrides"
    
End Sub

