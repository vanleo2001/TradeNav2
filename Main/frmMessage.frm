VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmMessage 
   Caption         =   "Message"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   Icon            =   "frmMessage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2700
      Visible         =   0   'False
      Width           =   3075
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
      Caption         =   "frmMessage.frx":0442
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMessage.frx":0476
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMessage.frx":0496
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdPrint 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   0
         Width           =   795
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
         Caption         =   "frmMessage.frx":04B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMessage.frx":04DE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMessage.frx":04FE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkTopmost 
         Height          =   220
         Left            =   1860
         TabIndex        =   3
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmMessage.frx":051A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMessage.frx":0552
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMessage.frx":0588
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClose 
         Cancel          =   -1  'True
         Default         =   -1  'True
         Height          =   300
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   795
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
         Caption         =   "frmMessage.frx":05A4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMessage.frx":05D0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMessage.frx":05F0
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniRichTextBoxXP rtbMessage 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   4683
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmMessage.frx":060C
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
      Tip             =   "frmMessage.frx":062C
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMessage.frx":064C
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
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum eMessageDisplayMode
    eNormalMessage = 0
    eStayOnTopMessage = 1
    eModalMessage = 2
End Enum

Private Type mPrivate
    bTopMost As Boolean
    bTextFile As Boolean
End Type
Private m As mPrivate

Private Sub chkTopmost_Click()
On Error GoTo ErrSection:

    TopMost = -chkTopmost.Value
    MoveFocus rtbMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMessage.chkTopmost.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    If Me Is frmMessage Then
        DockState(Me) = eHidden
    Else
        Unload Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMessage.cmdClose.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMessage.cmdPrint.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub Form_Activate()
On Error GoTo ErrSection:

    MoveFocus rtbMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMessage.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    SetPrevActiveForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMessage.Form.Deactivate", eGDRaiseError_Show
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
    RaiseError "frmMessage.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("ID_News"))
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMessage.Form.Load", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    If WindowStateX(Me) <> wsNormal Then TopMost = False
    
    If fraButtons.Enabled Then
        fraButtons.Top = Me.ScaleHeight - fraButtons.Height
        With rtbMessage
            .Move .Left, .Top, Me.ScaleWidth - .Left * 2, fraButtons.Top - .Top
        End With
    Else
        With rtbMessage
            .Move .Left, .Top, Me.ScaleWidth - .Left * 2, Me.ScaleHeight - .Top - .Left
        End With
    End If

End Sub

Public Sub ShowMe(ByVal strCaption$, ByVal strMsg$, _
        Optional ByVal eMessageMode As eMessageDisplayMode = eStayOnTopMessage, _
        Optional ByVal bCenteredIfTxtFile As Boolean = False)
On Error GoTo ErrSection:

    Dim i&
    Dim frmNew As frmMessage
    
    ' if a modal form is already up, then this must be modal as well
    If eMessageMode <> -99 And Not frmMain.Enabled Then
        eMessageMode = eModalMessage
    End If
    
    ' modal messages cannot use the docked form, so need
    ' to construct a new instance and use that one instead
    If eMessageMode = eModalMessage Then
        Set frmNew = New frmMessage
        frmNew.ShowMe strCaption, strMsg, -99, bCenteredIfTxtFile
        Set frmNew = Nothing
        Exit Sub
    ElseIf eMessageMode = -99 Then
        eMessageMode = eModalMessage
    End If

    chkTopmost.Visible = False
    fraButtons.Visible = False
    fraButtons.Enabled = False
    TopMost = False

    If Left(strMsg, 1) = "@" Then
        ' message is in this file
        strMsg = Mid(strMsg, 2)
        If InStr(Right(strMsg, 4), ".") = 0 Then
            ' look for RTF file first, then TXT file
            If FileExist(strMsg & ".rtf") Then
                strMsg = strMsg & ".rtf"
            ElseIf FileExist(strMsg & ".txt") Then
                strMsg = strMsg & ".txt"
            End If
        End If
        strMsg = FileToString(strMsg)
    End If
    
    If Len(strMsg) > 0 Then
        Me.Caption = strCaption
        With rtbMessage
            If Left(strMsg, 2) = "{\" Then
                .TextRTF = strMsg
                m.bTextFile = False
            Else
                m.bTextFile = True
                .Text = strMsg & vbCrLf
                If bCenteredIfTxtFile Then
                    .SelStart = 0
                    .SelLength = Len(.Text) + 10
                    .SelAlignment = rtfCenter
                End If
            End If
            
            .SelStart = 0 ' Len(.Text) + 10
            ' to get cursor out of view: set cursor to end,
            ' then use "SendMessage" to scroll to top
            'i = SendMessage(.hWnd, EM_GETLINECOUNT, 0, ByVal 0&)
            'i = -i
            'i = SendMessage(.hWnd, EM_LINESCROLL, 0, ByVal i)
        End With
        
        ' try with requested mode, if error (like trying to
        ' show non-modally over a modal form), then show modal
        If eMessageMode = eModalMessage Then GoTo DoModalMessage
        On Error GoTo DoModalMessage
        ''chkTopmost.Visible = True
        
        ' show based on if dockable form or if other instance
        If Me Is frmMessage Then
            DockState(Me) = eShowAsPrevious
        Else
            ShowForm Me, False, frmMain
        End If
        
        #If 0 Then
            ' if non-modal, set TopMost property
            If eMessageMode = eStayOnTopMessage Then
                TopMost = True
            Else
                TopMost = False
            End If
        #End If
    Else
        Beep
    End If
    Exit Sub
    
DoModalMessage:
    On Error Resume Next
    chkTopmost.Visible = False
    ' show buttons since being shown modally
    fraButtons.Visible = True
    fraButtons.Enabled = True
    ShowForm Me, True
    ''DockState(Me) = eShowAsPrevious
    Exit Sub

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMessage.ShowMe", eGDRaiseError_Raise
        
End Sub

Public Property Get TopMost() As Boolean
On Error GoTo ErrSection:

    TopMost = m.bTopMost

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmMessage.TopMost.Get", eGDRaiseError_Raise
        
End Property

Public Property Let TopMost(ByVal bTopMost As Boolean)
On Error GoTo ErrSection:

    If bTopMost Then
        SetFormTopmost Me, True
        chkTopmost = 1
    ElseIf m.bTopMost Then
        SetFormTopmost Me, False
        chkTopmost = 0
    End If
    m.bTopMost = bTopMost

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmMessage.TopMost.Let", eGDRaiseError_Raise
        
End Property

Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    With frmPrintPreview.vp
        .StartDoc
        .TextRTF = rtbMessage.TextRTF
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMessage.GenerateReport", eGDRaiseError_Raise
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    ' need this check so the "new modal instance"
    ' won't unload the docked instance
    If Me Is frmMessage Then
        frmMain.DockPro.RemoveForm Me.Name
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMessage.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Public Sub PrintMe()
On Error GoTo ErrSection:

    If m.bTopMost Then SetFormTopmost Me, False
    
    'rtbMessage.SelPrint Printer.hDC
    frmPrintPreview.ShowMe "CNV Message", Me

    MoveFocus rtbMessage
    If m.bTopMost Then SetFormTopmost Me, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMessage.PrintMe", eGDRaiseError_Raise
        
End Sub

