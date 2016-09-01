VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5115
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Default         =   -1  'True
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   3915
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
      Caption         =   "Splash.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "Splash.frx":0070
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "Splash.frx":0090
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   435
      Left            =   4200
      TabIndex        =   2
      Top             =   1440
      Width           =   735
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
      Caption         =   "Splash.frx":00AC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "Splash.frx":00D8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "Splash.frx":00F8
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniRichTextBoxXP rtbInfo 
      Height          =   3675
      Left            =   240
      TabIndex        =   1
      Top             =   2100
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6482
      BackColor       =   15332339
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "Splash.frx":0114
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Tip             =   "Splash.frx":0134
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Splash.frx":0154
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
   Begin vsOcx6LibCtl.vsElastic vseMessage 
      Height          =   315
      Left            =   60
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1020
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   3
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   7368816
      ForeColor       =   54000
      FloodColor      =   5177367
      ForeColorDisabled=   -2147483631
      Caption         =   "Message ..."
      Align           =   0
      Appearance      =   3
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   1
      FloodPercent    =   50
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
   Begin HexUniControls.ctlUniLabelXP lblMessage 
      Height          =   240
      Left            =   300
      Top             =   60
      Visible         =   0   'False
      Width           =   4875
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
      Caption         =   "Splash.frx":0170
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "Splash.frx":01E8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Splash.frx":0208
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lbl1 
      Height          =   675
      Left            =   360
      Top             =   240
      Width           =   4335
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
      Caption         =   "Splash.frx":0224
      BackColor       =   -2147483633
      ForeColor       =   12582912
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "Splash.frx":0262
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Splash.frx":0282
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Image imgSplash 
      BorderStyle     =   1  'Fixed Single
      Height          =   420
      Left            =   60
      Top             =   120
      Visible         =   0   'False
      Width           =   1080
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Me.Tag = "NO"
    SetIniFileProperty "Disclaimer", 0, "General", g.strIniFile
    HideButtons

End Sub

Private Sub cmdOK_Click()

    Me.Tag = ""
    HideButtons

End Sub

Private Sub Form_Activate()

    Static bAlreadyDone As Boolean
    
    If Not bAlreadyDone Then
        bAlreadyDone = True
        'SetFormTopmost Me, True
    End If

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
    RaiseError "frmSplash.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    'If KeyAscii = 27 Then KeyAscii = Asc("Q")
    'Me.Tag = Chr(KeyAscii)

    If KeyAscii = 13 Then
        Me.Tag = ""
        HideButtons
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSplash.Form.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim fName$, nSpace&, dtDisclaimer#
    
    g.Styler.StyleForm Me
    
    ' copy from old location if not in new location yet
    If Not FileExist(App.Path & "\Info\splash*.jpg") Then
        fName = "\Splash.jpg"
        If FileExist(App.Path & fName) Then
            FileCopy App.Path & fName, App.Path & "\Info" & fName
        End If
        fName = "\SplashBT.jpg"
        If FileExist(App.Path & fName) Then
            FileCopy App.Path & fName, App.Path & "\Info" & fName
        End If
    End If
    KillFile App.Path & "\SplashBT.jpg"
    
    g.strTitle = "Trade Navigator"
    fName = GetIniFileProperty("Splash", "", "General", g.strIniFile)
    If Len(fName) = 0 Then
        If ExtremeCharts >= 1 Then
            g.strTitle = "Extreme Charts"
            If HasModule("BTXA") Then
                fName = App.Path & "\Info\SplashBTXA.jpg"
                vseMessage.FloodColor = &H80&
            ElseIf IsRule1U Then
                fName = App.Path & "\Info\Rule1U.jpg"
            Else
                fName = App.Path & "\Info\SplashBT.jpg"
            End If
        ElseIf HasModule("STT", True) Then
            'g.strTitle = "STT Navigator"
            fName = App.Path & "\Info\STT.jpg"
        ElseIf HasModule("UNITEDFT", True) Then
            'g.strTitle = "STT Navigator"
            fName = App.Path & "\Info\UnitedFutures.jpg"
        ElseIf HasModule("TSU", True) Then
            g.strTitle = "TradeSmart Navigator"
            fName = App.Path & "\Info\TradeSmart.jpg"
        'ElseIf HasModule("LWIA", True) Then
            'g.strTitle = "LW Indicator Analyst"
            'fName = App.Path & "\Info\LWIA.jpg"
        ' TLB 11/29/2011: no more Rockwell Navigator per request by Markus
        'ElseIf HasModule("ROCKS", True) Or HasModule("ROCKF", True) Or HasModule("ROCKROOM", True) Then
        '    g.strTitle = "Rockwell Navigator"
        '    fName = App.Path & "\Info\Rockwell.jpg"
        ElseIf HasModule("WOODCCI", True) Then
            g.strTitle = "Woodies CCI Navigator"
            fName = App.Path & "\Info\WoodCCI.jpg"
        ElseIf GetSourceCode = "NATIONAL" Then
            g.strTitle = "Person's Navigator"
            fName = App.Path & "\Info\Persons2.jpg"
        ElseIf GetSourceCode = "FXWORLD" Then
            'g.strTitle = "FxProBE's Trade Navigator"
            fName = App.Path & "\Info\FxWorld.jpg"
        ElseIf UCase(Left(GetSourceCode, 5)) = "NISON" Then ' HasModule("NISON", True) Then
            g.strTitle = "Candlestick Navigator"
            fName = App.Path & "\Info\Nison.jpg"
        'ElseIf IsPfgVersion(False) Then 'TLB 10/14/2009: per Glen/Pete, now only check the source code
        ElseIf 0 Then ' TLB 5/24/2012: per Glen, no longer using "BEST Direct" title and splash
            g.strTitle = "BEST Direct Navigator"
            If HasModule("PFG-SECRETS", True) Then
                fName = App.Path & "\Info\BDSecrets.jpg"
            Else
                fName = App.Path & "\Info\BDNav.jpg"
            End If
        ElseIf GetSourceCode = "FXMTS" Or GetSourceCode = "FXPROBE" Then
            'g.strTitle = "FxProBE's Trade Navigator"
            fName = App.Path & "\Info\FxMTS.jpg"
        ElseIf IsLearnFxVersion Then
            g.strTitle = "Learn:Forex's Trade Navigator"
            fName = App.Path & "\Info\LearnFX.jpg"
        End If
    End If
If 0 And IsIDE Then
            g.strTitle = "Extreme Charts"
            fName = App.Path & "\Info\SplashBT.jpg"
            g.strTitle = "BEST Direct Navigator"
            fName = App.Path & "\Info\BDNav.jpg"
            g.strTitle = "Rockwell Navigator"
            fName = App.Path & "\Info\Rockwell.jpg"
            g.strTitle = "Woodies CCI Navigator"
            fName = App.Path & "\Info\WoodCCI.jpg"
End If
    If FileLength(fName) < 10 Then
        fName = App.Path & "\Info\Splash.jpg"
    ElseIf Not FileExist(fName) Then
        fName = App.Path & "\Info\Splash.jpg"
    End If
    
    lbl1.Visible = False
    If FileExist(fName) Then
        imgSplash.Picture = LoadPicture(fName)
        If UCase(FileBase(fName)) <> "SPLASH" Then
            ' if not main Genesis splash, then use black background for flood color
            vseMessage.FloodColor = 0
        End If
    End If
    vseMessage.FloodPercent = 0
    vseMessage.Caption = ""
    imgSplash.Visible = True
    imgSplash.Move 0, 0
    vseMessage.Move imgSplash.Left, imgSplash.Top + imgSplash.Height, imgSplash.Width
       
    ' Only need to show the disclaimer whenever it has changed (check for newer)
    dtDisclaimer = GetIniFileProperty("Disclaimer", 0#, "General", g.strIniFile)
If IsIDE Then
    'dtDisclaimer = 0
End If
    fName = App.Path & "\Info\Disclaimer.rtf"
    ' (require at least 2 hours newer just to bypass DST changes)
    If FileDate(fName) > dtDisclaimer + 121 / 1440# And FileLength(fName) > 50 Then
        ' WITH the disclaimer
        SetIniFileProperty "Disclaimer", FileDate(fName), "General", g.strIniFile
        rtbInfo.TextRTF = FileToString(fName)
        'rtbInfo.BackColor = Me.BackColor
        Me.Tag = "Wait"
        nSpace = Screen.TwipsPerPixelY * 4
        cmdOK.Move (vseMessage.Width - (cmdOK.Width + cmdCancel.Width)) / 3, _
                    vseMessage.Top + vseMessage.Height + nSpace
        cmdCancel.Move vseMessage.Width - cmdCancel.Width - cmdOK.Left, cmdOK.Top
        'rtbInfo.Move nSpace, cmdOK.Top + cmdOK.Height + nSpace, vseMessage.Width - nSpace * 2
        rtbInfo.Move 0, cmdOK.Top + cmdOK.Height + nSpace, vseMessage.Width
        Me.Move Me.Left, Me.Top, imgSplash.Width + imgSplash.Left * 2 + Me.Width - Me.ScaleWidth, _
                rtbInfo.Top + rtbInfo.Height + nSpace + Me.Height - Me.ScaleHeight
    Else
        ' WITHOUT the disclaimer
        Me.Tag = ""
        rtbInfo.Visible = False
        cmdOK.Visible = False
        cmdCancel.Visible = False
        Me.Move Me.Left, Me.Top, imgSplash.Width + imgSplash.Left * 2 + Me.Width - Me.ScaleWidth, _
                vseMessage.Top + vseMessage.Height + Me.Height - Me.ScaleHeight
    End If

    CenterTheForm Me
    
    SetIniFileProperty "Title", g.strTitle, "Main", App.Path & "\TNArchive.INI"
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSplash.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Public Sub Message(ByVal nPercent&, Optional ByVal strMessage$ = "")
On Error GoTo ErrSection:

    Static strPrevMessage$
    
    With frmSplash.vseMessage
        If nPercent >= 0 Then
            .FloodPercent = nPercent
        End If
        If Len(strMessage) = 0 Then
            strMessage = strPrevMessage
        Else
            strPrevMessage = strMessage
        End If
        'strMessage = CStr(nPercent) & "% " & strMessage
        .Caption = strMessage
        ''.Refresh
        DoEvents
    End With

    'Sleep 1
    'DoEvents

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSplash.Message", eGDRaiseError_Raise

End Sub

Private Sub HideButtons()
      
If 0 Then
    cmdCancel.Visible = False
    cmdOK.Visible = False
    vseMessage.Visible = True
    'rtbInfo.Height = cmdOK.Top + cmdOK.Height - rtbInfo.Top
    rtbInfo.Height = Me.ScaleHeight - rtbInfo.Top - rtbInfo.Left
    'Me.Height = vseMessage.Top + vseMessage.Height + Me.Height - Me.ScaleHeight
Else
    cmdCancel.Enabled = False
    cmdOK.Enabled = False
End If
    Me.Refresh

End Sub

