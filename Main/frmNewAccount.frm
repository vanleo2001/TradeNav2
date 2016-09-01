VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmNewAccount 
   Caption         =   "Account Configuration"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniFrameWL fraAccountInfo 
      Height          =   4755
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   6915
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
      Caption         =   "frmNewAccount.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmNewAccount.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmNewAccount.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optNew 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   3120
         Width           =   5055
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
         Caption         =   "frmNewAccount.frx":005C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   12582912
         Pressed         =   0   'False
         Tip             =   "frmNewAccount.frx":00B0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNewAccount.frx":00D0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExisting 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   900
         Width           =   5055
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
         Caption         =   "frmNewAccount.frx":00EC
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   12582912
         Pressed         =   -1  'True
         Tip             =   "frmNewAccount.frx":015A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNewAccount.frx":017A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraExisting 
         Height          =   1995
         Left            =   240
         TabIndex        =   3
         Top             =   900
         Width           =   6495
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
         Caption         =   "frmNewAccount.frx":0196
         Enabled         =   -1  'True
         ForeColor       =   16711680
         BackColor       =   -2147483633
         Tip             =   "frmNewAccount.frx":0208
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewAccount.frx":0228
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtPassword 
            Height          =   285
            Left            =   1620
            TabIndex        =   10
            Top             =   1140
            Width           =   1755
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmNewAccount.frx":0244
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
            Tip             =   "frmNewAccount.frx":0264
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewAccount.frx":0284
         End
         Begin HexUniControls.ctlUniTextBoxXP txtDataServID 
            Height          =   285
            Left            =   3840
            TabIndex        =   8
            Top             =   720
            Width           =   495
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmNewAccount.frx":02A0
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
            Tip             =   "frmNewAccount.frx":02C6
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewAccount.frx":02E6
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomerID 
            Height          =   285
            Left            =   1620
            TabIndex        =   6
            Top             =   720
            Width           =   855
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmNewAccount.frx":0302
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
            Tip             =   "frmNewAccount.frx":0322
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewAccount.frx":0342
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdContinue 
            Height          =   615
            Left            =   4680
            TabIndex        =   5
            Top             =   840
            Width           =   1275
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
            Caption         =   "frmNewAccount.frx":035E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmNewAccount.frx":0390
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmNewAccount.frx":03B0
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblGenesisUser 
            Height          =   255
            Left            =   300
            Top             =   1620
            Width           =   6075
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
            Caption         =   "frmNewAccount.frx":03CC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmNewAccount.frx":048C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewAccount.frx":04AC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblCustomerID 
            Height          =   255
            Left            =   480
            Top             =   780
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
            Caption         =   "frmNewAccount.frx":04C8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmNewAccount.frx":0502
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewAccount.frx":0522
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDataServiceID 
            Height          =   255
            Left            =   2400
            Top             =   780
            Width           =   1395
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
            Caption         =   "frmNewAccount.frx":053E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmNewAccount.frx":057A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewAccount.frx":059A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPassword 
            Height          =   255
            Left            =   480
            Top             =   1140
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
            Caption         =   "frmNewAccount.frx":05B6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmNewAccount.frx":05EA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewAccount.frx":060A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label1 
            Height          =   315
            Left            =   300
            Top             =   360
            Width           =   5895
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
            Caption         =   "frmNewAccount.frx":0626
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmNewAccount.frx":06EE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewAccount.frx":070E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraNew 
         Height          =   1395
         Left            =   240
         TabIndex        =   7
         Top             =   3180
         Width           =   6495
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
         Caption         =   "frmNewAccount.frx":072A
         Enabled         =   -1  'True
         ForeColor       =   16711680
         BackColor       =   -2147483633
         Tip             =   "frmNewAccount.frx":0786
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewAccount.frx":07A6
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdRegister 
            Height          =   495
            Left            =   3600
            TabIndex        =   9
            Top             =   780
            Visible         =   0   'False
            Width           =   2355
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
            Caption         =   "frmNewAccount.frx":07C2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmNewAccount.frx":080E
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmNewAccount.frx":082E
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblWarning 
            Height          =   435
            Left            =   480
            Top             =   840
            Visible         =   0   'False
            Width           =   2775
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
            Caption         =   "frmNewAccount.frx":084A
            BackColor       =   -2147483633
            ForeColor       =   255
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmNewAccount.frx":08EC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewAccount.frx":090C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblNewUser 
            Height          =   615
            Left            =   300
            Top             =   360
            Width           =   5835
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
            Caption         =   "frmNewAccount.frx":0928
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmNewAccount.frx":0A12
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewAccount.frx":0A32
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniLabelXP lblAccountInfo 
         Height          =   555
         Left            =   180
         Top             =   180
         Width           =   6495
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
         Caption         =   "frmNewAccount.frx":0A4E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmNewAccount.frx":0B94
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewAccount.frx":0BB4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin SHDocVwCtl.WebBrowser webBrowser 
      Height          =   3015
      Left            =   2160
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   3255
      ExtentX         =   5741
      ExtentY         =   5318
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmNewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmNewAccount.frm
'' Description: Allow a new user to fill in user information on the web
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/21/2015   DAJ         Send URL's through FixURL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strStartingAddress As String        ' Web address to start at
    strCustomerID As String             ' Customer ID
    strDataServiceID As String          ' Data Service ID
    strPassword As String               ' Password
    bRegistered As Boolean
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up and show the form
'' Inputs:      None
'' Returns:     True if user finished, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe() As Boolean
On Error GoTo ErrSection:

    Dim strPrevReg$, strMsg$, lCustID&
       
    ' Retrieve from ini file
    strPrevReg = GetIniFileProperty("TradeNavReg", "", "LOGIN", "navwin.ini")
    If ExtremeCharts >= 1 Then
        strPrevReg = "YES"
    End If
    
    'strMsg = GetProvidedProperty("CompanyName", , True) & " at " & GetProvidedProperty("SalesContact", , True)
    ' TLB 11/8/2011: skip company name since INI file now has longer phone # (both the 800 and 719 #)
    strMsg = GetProvidedProperty("SalesContact", , True)
    lblGenesisUser.Caption = Replace(lblGenesisUser.Caption, "GenesisFT at 800-808-3282", strMsg)
    
    ' hide Register frame if not allowed to create a new account
    If Len(strPrevReg) > 0 Then
        fraNew.Visible = False
        optNew.Visible = False
        optExisting.Visible = False
        fraExisting.Caption = " Account Information "
        'RH commented out fraExisting.BorderStyle = 1
        'If fraExisting.Top > fraNew.Top Then fraExisting.Top = fraNew.Top
        Me.Height = 3780
    End If
    
    ' see if had tried to put something in previously
    lCustID = RI_GetDataServiceID
    If lCustID >= 1000 Then
        m.strCustomerID = Int(lCustID / 1000)
        m.strDataServiceID = Right(Str(lCustID), 3)
        m.strPassword = Trim(RI_GetUserPassword)
    End If
    
    DebugLog "Before NewAccount shown:" & vbTab & m.strCustomerID & vbTab & m.strDataServiceID & vbTab & m.strPassword
    EnableButtons
    ShowForm Me, eForm_Modal, frmMain
    DebugLog "After NewAccount shown:" & vbTab & m.strCustomerID & vbTab & m.strDataServiceID & vbTab & m.strPassword
    
    If Len(m.strCustomerID) > 0 And Len(m.strDataServiceID) > 0 Then
        ' after auto-subscribing, we need to try to have them wait 30-60 seconds
        ' before downloading starter data set (so codes get propogated to the validators)
        If m.bRegistered Then
            SetIniFileProperty "TradeNavReg", m.strCustomerID & ":" & m.strDataServiceID, "LOGIN", "navwin.ini"
            strMsg = "Please write down your new account| information for future reference:||Customer ID = " _
                & m.strCustomerID & "|Data Service = " & m.strDataServiceID & "|Password = " & m.strPassword
            InfBox strMsg, "!", , "IMPORTANT"
        End If
        RI_SetDataServiceID CLng(Val(m.strCustomerID & Format(Val(m.strDataServiceID), "000")))
        RI_SetUserPassword m.strPassword
        Unload Me
        ShowMe = True
    Else
        If m.bRegistered Then
            ' if attempted to register but failed for some reason (problems sometimes with our servers?),
            ' then at least leave some kind of message for them (e.g. sometimes the email will still get sent)
            strMsg = "If a new account was created, you can check your email for your Customer ID and Password.|| Please call 800-808-3282 if any problems."
            InfBox strMsg, "i", , "New Account"
        End If
        Unload Me
        ShowMe = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmNewAccount.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdContinue_Click
'' Description: Allow the user to go to continue on once they entered info
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdContinue_Click()
On Error GoTo ErrSection:

    Dim strPassword As String           ' Password confirmation
    Dim d#

    If Trim(txtCustomerID.Text) = "" Then
        InfBox "You need to enter in a Customer ID|before you can continue using the software", "!", , "Error"
        Exit Sub
    End If
    ' TLB 11/9/2011: "1" is no longer a valid CustomerID (since some were inadvertantly putting their DataServiceID here)
    d = ValOfText(txtCustomerID.Text)
    If IsAlpha(Trim(txtCustomerID.Text)) Or d < 2 Or d > 999999 Then
        InfBox "This is not a valid Genesis Customer ID.|Please contact Genesis Sales at 800-808-3282.", "!", , "Invalid Customer ID"
        Exit Sub
    End If
    
    If Trim(txtDataServID.Text) = "" Then
        InfBox "You need to enter in a Data Service ID|before you can continue using the software", "!", , "Error"
        Exit Sub
    End If
    d = ValOfText(txtDataServID.Text)
    If IsAlpha(Trim(txtDataServID.Text)) Or d < 1 Or d > 999 Then
        InfBox "This is not a valid Data Service ID.|Please contact Genesis Sales at 800-808-3282.", "!", , "Invalid ID"
        Exit Sub
    End If
    
    If Len(Trim(txtPassword.Text)) = 0 Then
        InfBox "You need to enter in a Password|before you can continue using the software", "!", , "Error"
        Exit Sub
    End If

    strPassword = InfBox("Please enter password again to confirm ...", "?", , "Confirm Password", , , , , , "p")
    If strPassword <> txtPassword.Text Then
        InfBox "Second password does not match original", "!", , "Error"
        Exit Sub
    End If
    
    m.strCustomerID = Trim(txtCustomerID.Text)
    m.strDataServiceID = Trim(txtDataServID.Text)
    m.strPassword = Trim(txtPassword.Text)
    m.bRegistered = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewAccount.cmdContinue.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRegister_Click
'' Description: Allow the user to go to the Genesis website for registration
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRegister_Click()
On Error GoTo ErrSection:

    Dim aFile As New cGdArray

    fraAccountInfo.Visible = False
    WebBrowser.Visible = True
    WindowState = 2
    DoEvents

    aFile.FromFile App.Path & "\Provided\Install.CFG"
    m.strStartingAddress = Trim(aFile(1))
    If Len(m.strStartingAddress) = 0 Then
        m.strStartingAddress = "http://bryan/OrderWizard/TradeNavTrial/CustomerTrial.aspx"
    End If
    
    m.strStartingAddress = FixURL(m.strStartingAddress)
    
    WebBrowser.Navigate m.strStartingAddress
    m.bRegistered = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewAccount.cmdRegister.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form members
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    'Dim strText As String
    'Dim aStrings As New cGdArray

    ' form placement
    'Move Left, Top, 7935, 6000
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    Icon = Picture16("kBlank")

    ' get starting address
    ''aStrings.FromFile App.Path & "\Provided\???.adr"
    ''m.strStartingAddress = aStrings(0)
    'If Len(m.strStartingAddress) = 0 Then
    '    m.strStartingAddress = "http://bryan/OrderWizard/TradeNavTrial/CustomerTrial.aspx"
    'End If
    
    ' Set the encryption key
'    Dim mbKey As New cMemBuffer
'    If mbKey.Length = 0 Then
'        mbKey.PutByte 71
'        mbKey.PutByte 202
'        mbKey.PutByte 123
'        mbKey.PutByte 63
'        mbKey.PutByte 176
'        mbKey.PutByte 2
'        mbKey.PutByte 70
'        mbKey.PutByte 198
'        mbKey.PutByte 169
'        mbKey.PutByte 85
'        mbKey.PutByte 10
'    End If
    
    ' Get the data service ID, encrypt it, convert it to hex and add it to the address
'    Dim mbMemory As New cMemBuffer
'    mbMemory.PutStr RI_GetDataServiceID
'    Dim lRet As Long
'    lRet = gdEncrypt(True, mbMemory, mbKey)
'    Dim strParm As String
'    strParm = Space(mbMemory.Length * 2 + 1)
'    Dim nLen As Long
'    nLen = gdHexMemory(strParm, mbMemory.MemPtr, mbMemory.Length)
'    strParm = Left(strParm, nLen)
'    m.strStartingAddress = m.strStartingAddress & "?s=" & strParm
    
    ' Add the password
'    mbMemory.Clear
'    mbMemory.PutStr RI_GetUserPassword
'    lRet = gdEncrypt(True, mbMemory, mbKey)
'    strParm = Space(mbMemory.Length * 2 + 1)
'    nLen = gdHexMemory(strParm, mbMemory.MemPtr, mbMemory.Length)
'    strParm = Left(strParm, nLen)
'    m.strStartingAddress = m.strStartingAddress & "&p=" & strParm
    
    ' Add a date string
'    mbMemory.Clear
'    mbMemory.PutStr CStr(ConvertTimeZone(Now, "", "NY"))
'    lRet = gdEncrypt(True, mbMemory, mbKey)
'    strParm = Space(mbMemory.Length * 2 + 1)
'    nLen = gdHexMemory(strParm, mbMemory.MemPtr, mbMemory.Length)
'    strParm = Left(strParm, nLen)
'    m.strStartingAddress = m.strStartingAddress & "&d=" & strParm

    ' Load the new account page
    'WebBrowser.Navigate m.strStartingAddress

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewAccount.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Close the form ourselves if the user hits the 'X'
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewAccount.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move controls as the form gets resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next
    
    Dim nSpace&
    
    If LimitFormSize(Me, 2000, 2000) Then Exit Sub
    
    nSpace = 60
    With WebBrowser
        .Move nSpace, nSpace, Me.ScaleWidth - nSpace * 2, Me.ScaleHeight - nSpace * 2
    End With

End Sub

Private Sub optExisting_Click()
    EnableButtons
    MoveFocus txtCustomerID
End Sub

Private Sub optNew_Click()
    EnableButtons
    MoveFocus cmdRegister
End Sub

Private Sub txtCustomerID_Change()
    EnableButtons
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtCustomerID_GotFocus
'' Description: When the text box gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtCustomerID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtCustomerID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewAccount.txtCustomerID.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDataServID_Change()
    EnableButtons
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDataServID_GotFocus
'' Description: When the text box gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDataServID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtDataServID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewAccount.txtDataServID.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDataServID_LostFocus()
On Error GoTo ErrSection:

    Dim d#
    
    d = ValOfText(txtDataServID)
    If d < 1 Or d > 999 Then
        d = 1
    Else
        d = Int(d)
    End If
    txtDataServID = Format(d, "000")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewAccount.txtDataServID.LostFocus", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub txtPassword_Change()
    EnableButtons
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_GotFocus
'' Description: When the text box gets the focus, select all of the text
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
    RaiseError "frmNewAccount.txtPassword.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    WebBrowser_DownloadBegin
'' Description: Do some initialization as the download of the website begins
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub webBrowser_DownloadBegin()
On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    WebBrowser.Document.body.innerHtml = "<br><br><br><br><p align=center><font size=5 color='#009900'><bold>Loading...</bold></font></p>"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    WebBrowser_DownloadComplete
'' Description: The download of the website has completed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub webBrowser_DownloadComplete()
On Error Resume Next
    
    Screen.MousePointer = vbDefault

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    WebBrowser_DocumentComplete
'' Description: Once the user is finished, grab the necessary information
'' Inputs:      Display object, URL
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub webBrowser_DocumentComplete(ByVal pDisp As Object, Url As Variant)
On Error Resume Next
    
    If WebBrowser.Visible Then
        ' Get the customer ID, data service ID and password from the web page...
        If Not WebBrowser.Document.All("lblCustID") Is Nothing Then
            m.strCustomerID = WebBrowser.Document.All("lblCustID").innerText
            m.strDataServiceID = WebBrowser.Document.All("lblDSrvID").innerText
            m.strPassword = WebBrowser.Document.All("lblPassWord").innerText
            DebugLog "In NewAccount DocComplete:" & vbTab & m.strCustomerID & vbTab & m.strDataServiceID & vbTab & m.strPassword
        End If
        
        ' If the close label is filled in, then we need to close the form...
        If Not WebBrowser.Document.All("lblClose") Is Nothing Then
            If Len(WebBrowser.Document.All("lblClose").innerText) > 0 Then
                Me.Hide
            End If
        End If
    End If

End Sub

Public Sub EnableButtons()

    On Error Resume Next
    
    If optNew.Value <> 0 Then
        fraExisting.Enabled = False
        fraNew.Enabled = True
        lblWarning.Visible = True
        cmdRegister.Visible = True
        cmdContinue.Enabled = False
    Else
        fraExisting.Enabled = True
        fraNew.Enabled = False
        lblWarning.Visible = False
        cmdRegister.Visible = False
        cmdContinue.Enabled = True
    End If
    
  Exit Sub
    
    If Len(Trim(Me.txtCustomerID)) > 0 And Len(Trim(Me.txtPassword)) > 0 And Len(Trim(Me.txtDataServID)) > 0 Then
        If Not cmdContinue.Enabled Then
            cmdRegister.Enabled = False
            cmdContinue.Enabled = True
            cmdContinue.Default = True
        End If
    ElseIf Not cmdRegister.Enabled Then
        cmdRegister.Enabled = True
        cmdContinue.Enabled = False
    End If

End Sub

