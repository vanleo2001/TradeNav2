VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmImageServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Image Server"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniCheckXP chkActive 
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   2640
      Width           =   2235
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
      Caption         =   "frmImageServer.frx":0000
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmImageServer.frx":0048
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmImageServer.frx":0068
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   4380
      TabIndex        =   2
      Top             =   2580
      Width           =   1035
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
      Caption         =   "frmImageServer.frx":0084
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmImageServer.frx":00B2
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmImageServer.frx":00D2
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdSave 
      Default         =   -1  'True
      Height          =   435
      Left            =   3120
      TabIndex        =   1
      Top             =   2580
      Width           =   1035
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
      Caption         =   "frmImageServer.frx":00EE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmImageServer.frx":0118
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmImageServer.frx":0138
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5475
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
      Caption         =   "frmImageServer.frx":0154
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmImageServer.frx":0184
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmImageServer.frx":01A4
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtFontSize 
         Height          =   315
         Left            =   4740
         TabIndex        =   5
         Top             =   1770
         Width           =   555
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmImageServer.frx":01C0
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
         Tip             =   "frmImageServer.frx":01E6
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmImageServer.frx":0206
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPixels 
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   1770
         Width           =   615
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmImageServer.frx":0222
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
         Tip             =   "frmImageServer.frx":0248
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmImageServer.frx":0268
      End
      Begin HexUniControls.ctlUniTextBoxXP txtRexQ 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   1260
         Width           =   5115
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmImageServer.frx":0284
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
         Tip             =   "frmImageServer.frx":02DE
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmImageServer.frx":02FE
      End
      Begin HexUniControls.ctlUniTextBoxXP txtImgSrvQ 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   540
         Width           =   5115
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmImageServer.frx":031A
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
         Tip             =   "frmImageServer.frx":036A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmImageServer.frx":038A
      End
      Begin HexUniControls.ctlUniLabelXP Label5 
         Height          =   255
         Left            =   2880
         Top             =   1800
         Width           =   1815
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
         Caption         =   "frmImageServer.frx":03A6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmImageServer.frx":03FA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmImageServer.frx":041A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   180
         Top             =   1800
         Width           =   1755
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
         Caption         =   "frmImageServer.frx":0436
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmImageServer.frx":0488
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmImageServer.frx":04A8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   195
         Left            =   180
         Top             =   1020
         Width           =   3195
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
         Caption         =   "frmImageServer.frx":04C4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmImageServer.frx":0528
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmImageServer.frx":0548
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   195
         Left            =   180
         Top             =   300
         Width           =   3375
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
         Caption         =   "frmImageServer.frx":0564
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmImageServer.frx":05DE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmImageServer.frx":05FE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP Label4 
      Height          =   435
      Left            =   240
      Top             =   3120
      Width           =   5235
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
      Caption         =   "frmImageServer.frx":061A
      BackColor       =   -2147483633
      ForeColor       =   12582912
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmImageServer.frx":0754
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmImageServer.frx":0774
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmImageServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    strImgSrvQ As String
    strRexQ As String
    nPixels As Long
    nFontSize As Long
    bActive As Boolean
    bSaved As Boolean
End Type
Private m As mPrivate

Private Sub chkActive_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub

    ' turn active/inactive immediately
    If chkActive = 0 Then
        m.bActive = False
    Else
        m.bActive = True
    End If
    SetImgSrvSearcher
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmImageServer.chkActive.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmImageServer.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrSection:

    ' verify
    If chkActive Then
        If Not DirExist(Trim(txtRexQ)) Then
            InfBox "The specified RexQ does not exist!", "e", , "ERROR"
            MoveFocus txtRexQ
            Exit Sub
        End If
        If Len(Trim(txtImgSrvQ)) = 0 Then
            InfBox "No Image Server Queues specified!", "e", , "ERROR"
            MoveFocus txtImgSrvQ
            Exit Sub
        End If
    End If
    
    m.bSaved = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmImageServer.cmdSave.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmImageServer.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("ID_Tile"))
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmImageServer.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmImageServer.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Public Sub ShowMe(Optional ByVal bAutoStart As Boolean = False)
On Error GoTo ErrSection:

    ' load settings
    m.strImgSrvQ = GetIniFileProperty("ImgSrvQ", "", "Image Server", g.strIniFile)
    m.strRexQ = GetIniFileProperty("RexQ", "", "Image Server", g.strIniFile)
    m.nPixels = GetIniFileProperty("NumPixels", 480, "Image Server", g.strIniFile)
    m.nFontSize = GetIniFileProperty("FontSize", 150, "Image Server", g.strIniFile)
    m.bActive = GetIniFileProperty("Active", False, "Image Server", g.strIniFile)
    
    ' display settings
    txtImgSrvQ = m.strImgSrvQ
    txtRexQ = m.strRexQ
    txtPixels = CStr(m.nPixels)
    txtFontSize = CStr(m.nFontSize)
    If m.bActive Then
        chkActive = 1
    Else
        chkActive = 0
    End If
    
    If bAutoStart Then
        If InfBox("Auto-start the image server now?", "?", "+Auto-Start|-No", "Image Server", , 10) <> "A" Then
            bAutoStart = False
        End If
    End If
    
    If bAutoStart Then
        ' start now without showing form
        chkActive = 1
        m.bSaved = True
    Else
        ' show config form
        m.bSaved = False
        ShowForm Me, True
    End If
    
    If m.bSaved Then
        ' save settings
        m.strImgSrvQ = Trim(txtImgSrvQ)
        m.strRexQ = Trim(txtRexQ)
        m.nPixels = ValOfText(txtPixels)
        If m.nPixels <= 0 Then m.nPixels = 480
        m.nFontSize = ValOfText(txtFontSize)
        If m.nFontSize <= 0 Then m.nFontSize = 150
        If chkActive = 0 Then
            m.bActive = False
        Else
            m.bActive = True
        End If
        
        ' store settings
        Call SetIniFileProperty("ImgSrvQ", m.strImgSrvQ, "Image Server", g.strIniFile)
        Call SetIniFileProperty("RexQ", m.strRexQ, "Image Server", g.strIniFile)
        Call SetIniFileProperty("NumPixels", m.nPixels, "Image Server", g.strIniFile)
        Call SetIniFileProperty("FontSize", m.nFontSize, "Image Server", g.strIniFile)
        Call SetIniFileProperty("Active", m.bActive, "Image Server", g.strIniFile)
    End If
    
    SetImgSrvSearcher
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmImageServer.ShowMe", eGDRaiseError_Raise
    
End Sub

Public Property Get ImgSrvQ() As String
    ImgSrvQ = m.strImgSrvQ
End Property

Public Property Get RexQ() As String
    RexQ = AddSlash(m.strRexQ)
End Property

Public Property Get PixelWidth() As Long
    PixelWidth = m.nPixels
End Property

Public Property Get DefaultFontSize() As Long
    DefaultFontSize = m.nFontSize
End Property

Public Property Get Active() As Boolean
    Active = m.bActive
End Property

