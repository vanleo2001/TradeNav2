VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAsk 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2820
   ClientLeft      =   2340
   ClientTop       =   2355
   ClientWidth     =   4470
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2820
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniCheckXP chkDontAsk 
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ask32.frx":0000
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "Ask32.frx":004E
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "Ask32.frx":006E
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtInput 
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "Ask32.frx":008A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
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
      Tip             =   "Ask32.frx":00B4
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      OLEDragMode     =   0
      MousePointer    =   0
      MouseIcon       =   "Ask32.frx":00D4
   End
   Begin MSComctlLib.ProgressBar barTimeout 
      Height          =   252
      Left            =   1080
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   3132
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdButtons 
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ask32.frx":00F0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "Ask32.frx":0124
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "Ask32.frx":0144
      RightToLeft     =   0   'False
   End
   Begin VB.Timer tmrAsk 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   2160
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdButton 
      Height          =   400
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ask32.frx":0160
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "Ask32.frx":0192
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "Ask32.frx":01B2
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdButton 
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ask32.frx":01CE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "Ask32.frx":0200
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "Ask32.frx":0220
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdButton 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ask32.frx":023C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "Ask32.frx":026E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "Ask32.frx":028E
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblDay 
      Height          =   252
      Left            =   3000
      Top             =   960
      Width           =   1212
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ask32.frx":02AA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "Ask32.frx":02DA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Ask32.frx":02FA
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
      UseMnemonic     =   -1  'True
   End
   Begin HexUniControls.ctlUniLabelXP lblStop 
      Height          =   285
      Left            =   6975
      Top             =   225
      Width           =   675
      _ExtentX        =   767
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ask32.frx":0316
      BackColor       =   -2147483643
      ForeColor       =   16777215
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   -1  'True
      Tip             =   "Ask32.frx":033E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Ask32.frx":035E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
      UseMnemonic     =   -1  'True
   End
   Begin HexUniControls.ctlUniLabelXP lblInput 
      Height          =   255
      Left            =   240
      Top             =   1245
      Width           =   975
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ask32.frx":037A
      BackColor       =   -2147483633
      ForeColor       =   8421504
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "Ask32.frx":03A6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Ask32.frx":03C6
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
      UseMnemonic     =   -1  'True
   End
   Begin HexUniControls.ctlUniLabelXP lblMessage_BU 
      Height          =   15
      Left            =   210
      Top             =   240
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   26
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ask32.frx":03E2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   -1  'True
      Tip             =   "Ask32.frx":0402
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Ask32.frx":0422
      RightToLeft     =   0   'False
      WordWrap        =   -1  'True
      UseMnemonic     =   -1  'True
   End
   Begin HexUniControls.ctlUniLabelXP lblTimeout 
      Height          =   255
      Left            =   240
      Top             =   2400
      Width           =   855
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ask32.frx":043E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "Ask32.frx":046E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Ask32.frx":048E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
      UseMnemonic     =   -1  'True
   End
   Begin HexUniControls.ctlImageBag icoBag 
      Left            =   960
      Top             =   480
      _ExtentX        =   794
      _ExtentY        =   794
      Pics            =   9
      Pic1_Image32    =   "Ask32.frx":04AA
      Pic2_Image32    =   "Ask32.frx":13DC
      Pic3_Image32    =   "Ask32.frx":22D9
      Pic4_Image32    =   "Ask32.frx":2E32
      Pic5_Image32    =   "Ask32.frx":3B23
      Pic6_Image32    =   "Ask32.frx":460E
      Pic7_Image32    =   "Ask32.frx":5284
      Pic8_Image32    =   "Ask32.frx":5F4D
      Pic9_Image32    =   "Ask32.frx":6DFC
   End
   Begin HexUniControls.ctlUniImageWL icoUni 
      Height          =   550
      Left            =   180
      Top             =   240
      Visible         =   0   'False
      Width           =   550
      _ExtentX        =   979
      _ExtentY        =   979
      BackStyle       =   0
      Tip             =   "Ask32.frx":7797
      Enabled         =   -1  'True
      Border          =   0   'False
      BackColor       =   -2147483633
      BorderColor     =   -1
      RoundedBorders  =   -1  'True
      Stretch         =   -1  'True
      QualityStretch  =   -1  'True
      TransparencyType=   -1
      TransparentColor=   -1
      XTransp         =   0
      YTransp         =   0
      OLEDropMode     =   0
      MousePointer    =   0
      MouseIcon       =   "Ask32.frx":77B7
      RightToLeft     =   0   'False
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   4095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strArgs As String    ' for passing to/from form
Public nDefaultFontSize As Long 'default font size for application
Public nDefaultBackClr As Long  'default back color for application

Dim ctlFocus As Control     ' control to receive focus
Dim iNumButtons As Long     ' # of buttons visible
Dim iTimeout As Long        ' timeout (# seconds), or 0
Dim strGetString As String  ' what type of input string to get

Public Property Let Progress(ByVal lValue As Long)
    barTimeout.Value = lValue
End Property
Public Property Get Progress() As Long
    Progress = barTimeout.Value
End Property

Private Sub Form_Click()

    If iNumButtons = 0 Then
        Finish cmdButton(0)
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Dim i As Integer, iPos As Long, s As String

    ' if not getting input (text box not visible)
    If Me.Visible And Not txtInput.Visible Then
        If iNumButtons > 0 Then
            ' then see if key hit matches a button hot-key
            For i = 0 To 2
                If cmdButton(i).Visible Then
                    s = cmdButton(i).Caption
                    iPos = InStr(s, "&")
                    If iPos > 0 Then
                        s = UCase(Mid(s, iPos + 1, 1))
                        If s = UCase(Chr(KeyAscii)) Then
                            ' if so, execute button
                            Finish cmdButton(i)
                            Exit For
                        End If
                    End If
                End If
            Next
        ElseIf KeyAscii = 13 Or KeyAscii = 27 Then
            ' if no buttons then need to allow user to clear the message
            ' just in case the form happened to come up modally
            ' (e.g. when a nowait state gets displayed over a modal form,
            ' it comes up in a "act modal" state)
            Finish cmdButton(0)
        End If
    End If

End Sub

Private Sub Form_Load()

    Dim iPos&, strArg$, lHeight&, lMaxHeight&, lHorizSpace&, nBackClr&, strTemp$
    Dim strHeader$, strButtons$, strIcon$, strMessage$
    Dim strDefault$, bNoWait As Boolean, strFontBold$, nFontSize&
    Dim lAlignment As eGDAlignment
    Dim bShowDontAsk As Boolean
    Dim lProgress As Long
    

    ' Initialize variables.
    Set ctlFocus = Nothing
    iNumButtons = 0
    iTimeout = 0
    strGetString = ""
    lHeight = 0
    lHorizSpace = 200
    nFontSize = nDefaultFontSize
    strButtons = ""
    If nDefaultBackClr > 0 Then
        nBackClr = nDefaultBackClr
    Else
        'nBackClr = Me.BackColor
        nBackClr = GetAppBackColor
    End If
    lAlignment = eGDAlign_Center
    bShowDontAsk = False

    ' NoWait?
    If InStr(UCase(strArgs), "=NOWAIT") Then
        bNoWait = True
    Else
        bNoWait = False
    End If

    ' Parse main parameter string.
    If Len(strArgs) = 0 Then
        'Unload Me
        Exit Sub
    Else
        strArg = Trim(strArgs)
        Do While strArg <> ""
            iPos = InStr(strArg, " ; ")
            If iPos > 0 Then
                strArgs = Trim(Mid(strArg, iPos + 2)) ' rest of args
                strArg = Trim(Left(strArg, iPos - 1)) ' current arg
            Else
                strArgs = ""
            End If

            ' Parse specific parameter.
            iPos = InStr(strArg, "=")
            If iPos > 0 Then
                Select Case UCase(Left(strArg, 1))
                    Case "A" ' Alignment
                        lAlignment = CLng(Mid(strArg, iPos + 1))
                    Case "B" 'Buttons
                        strButtons = Mid(strArg, iPos + 1)
                    Case "C" 'Color
                        strTemp = Mid(strArg, iPos + 1)
                        If IsDigit(strTemp, 1) Then
                            nBackClr = Val(strTemp)
                        Else
                            nBackClr = QbClr(strTemp)
                        End If
                    Case "D" 'Default input string
                        strDefault = Mid(strArg, iPos + 1)
                    Case "F" 'FontBold
                        strFontBold = Mid(strArg, iPos + 1)
                    Case "G" 'Get (=string,date,number)
                        strGetString = Mid(strArg, iPos + 1)
                    Case "H" 'Header
                        strHeader = Mid(strArg, iPos + 1)
                    Case "I" 'Icon
                        strIcon = Mid(strArg, iPos + 1)
                    Case "M" 'Message
                        strMessage = Mid(strArg, iPos + 1)
                    Case "P"
                        lProgress = Val(Mid(strArg, iPos + 1))
                    ''Case "S" 'Stop icon
                        ''strIcon = "[" + Mid(strArg, iPos + 1) + "]"
                    Case "S" 'Font Size
                        nFontSize = Val(Mid(strArg, iPos + 1))
                    Case "T" 'Timeout
                        iTimeout = Val(Mid(strArg, iPos + 1))
                    Case "Z" 'Show Don't Ask box
                        bShowDontAsk = True
                End Select
            Else
                ' If no "=" sign, make it the message.
                strMessage = Mid(strArg, iPos + 1)
            End If

            strArg = Trim(strArgs)
        Loop
    End If

    ' Adjust color.
    Me.BackColor = nBackClr

    ' FontBold for which controls
    strFontBold = Trim(UCase(strFontBold))
    If strFontBold = "A" Then strFontBold = "MBITD"
    If InStr(strFontBold, "M") > 0 Then lblMessage.Font.Bold = True
    If InStr(strFontBold, "B") > 0 Then
        cmdButtons.Font.Bold = True
        cmdButton(0).Font.Bold = True
        cmdButton(1).Font.Bold = True
        cmdButton(2).Font.Bold = True
    End If
    If InStr(strFontBold, "I") > 0 Then
        txtInput.Font.Bold = True
        lblInput.Font.Bold = True
    End If
    If InStr(strFontBold, "T") > 0 Then lblTimeout.Font.Bold = True
    If InStr(strFontBold, "D") > 0 Then lblDay.Font.Bold = True

    ' Message.
    If nFontSize = 0 Then nFontSize = 8 'default
    lblMessage.Alignment = lAlignment
    lblMessage.Font.Size = nFontSize
    lblMessage.Caption = ""
    If strMessage <> "" Then
        If bNoWait Then strMessage = strMessage + "|"
        Do While True
            iPos = InStr(strMessage, "|")
            If iPos = 0 Then Exit Do
            strMessage = Left(strMessage, iPos - 1) + Chr(13) + Mid(strMessage, iPos + 1)
        Loop
        lblMessage.Caption = strMessage
        strMessage = ""
    End If

    ' Icons.
    'ico.BackColor = &HFFFFFF 'Me.BackColor
    'ico.Visible = False
    'ico.ZOrder 1
    Set icoUni.Picture = LoadPicture("")
    'lblStop.Visible = False
    If bNoWait And strIcon = "" Then strIcon = "Timer"
    ShowIcon strIcon

    ' Calculate "height" of Window so far.
    lblMessage.Left = lblTimeout.Left
    lblMessage.Width = lblTimeout.Width + barTimeout.Width
    lHeight = lblMessage.Top + lblMessage.Height + lHorizSpace
    lMaxHeight = Screen.Height - (Me.Height - Me.ScaleHeight) - 1500
    If lHeight > lMaxHeight Then
        lHeight = lMaxHeight
    End If
    If Len(strIcon) > 0 Or Len(strGetString) > 0 Then
        ' Leave room for Icon.
        'lblMessage.Left = ico.Left + ico.Width
        lblMessage.Left = icoUni.Left + icoUni.Width
        lblMessage.Width = barTimeout.Width + barTimeout.Left - lblMessage.Left
        If lHeight < icoUni.Top + icoUni.Height + lHorizSpace Then
            lHeight = icoUni.Top + icoUni.Height + lHorizSpace
        End If
        ' if one-liner, move message down a line.
        If lblMessage.Height < icoUni.Height / 2 Then
            lblMessage.Caption = Chr(13) + lblMessage.Caption
            If Len(strGetString) > 0 Then
                ' if a one-liner message, center over the input box
                lblMessage.Caption = lblMessage.Caption & Space(8)
            End If
        End If
    End If

    ' Input string.
    If strGetString = "" Then
        Disable txtInput
        txtInput.Visible = False
        lblInput.Visible = False
        lblDay.Visible = False
    Else
        ' if a one-liner message, center over the input box
        If lblMessage.Height < icoUni.Height / 2 Then
            'lblMessage.Caption = lblMessage.Caption & Space(8)
        End If
        txtInput.Top = lHeight
        Enable txtInput
        txtInput.Visible = True
        txtInput = strDefault
        If strButtons = "" Then strButtons = "+OK|-Cancel"
        lblInput.Top = lHeight + 50
        lblInput.Visible = False 'True
        lHeight = txtInput.Top + txtInput.Height + lHorizSpace
        
        ' Is it a password?
        If UCase(Left(strGetString, 1)) = "P" Then
            txtInput.PasswordChar = "*"
        Else
            txtInput.PasswordChar = ""
        End If
    
        Select Case UCase(Left(strGetString, 1))
        Case "D" ' Date (show Day label)
            lblDay.Height = txtInput.Height
            lblDay.Top = txtInput.Top
            txtInput.Left = lblTimeout.Left
            txtInput.Width = lblDay.Left - txtInput.Left
            lblDay.Caption = WeekdayName(txtInput, False)
            lblDay.Visible = True
        
        Case "N" ' Number (smaller input box)
            lblDay.Visible = False
            txtInput.Left = lblTimeout.Left + 1200
            txtInput.Width = lblDay.Left - txtInput.Left + lblDay.Width - 1200
        
        Case Else
            lblDay.Visible = False
            txtInput.Left = lblTimeout.Left
            txtInput.Width = lblDay.Left - txtInput.Left + lblDay.Width
        End Select
    End If

    ' Buttons.
    iNumButtons = 0
    Do While Left(strButtons, 1) = "|"
        strButtons = Mid(strButtons, 2)
    Loop
    If strButtons = "" And Not bNoWait Then strButtons = "+-OK"
    SetButton cmdButton(0), strButtons, lHeight
    SetButton cmdButton(1), strButtons, lHeight
    SetButton cmdButton(2), strButtons, lHeight
    Select Case iNumButtons
        Case 1
            cmdButton(0).Left = cmdButtons.Left
        Case 2
            cmdButton(0).Left = cmdButtons.Left - cmdButtons.Width * 0.6
            cmdButton(1).Left = cmdButtons.Left + cmdButtons.Width * 0.6
    End Select
    If iNumButtons > 0 Then lHeight = cmdButton(0).Top + cmdButton(0).Height + lHorizSpace

    ' Timeout/Progress
    If iTimeout <= 0 Then
        barTimeout.Visible = False
        lblTimeout.Visible = False
        Disable tmrAsk
    Else
        barTimeout.Min = 0
        barTimeout.Max = iTimeout
        barTimeout.Value = 0
        barTimeout.Visible = True
        barTimeout.Top = lHeight
        lblTimeout.Top = lHeight
        lblTimeout.Visible = True
        lHeight = lblTimeout.Top + lblTimeout.Height + lHorizSpace
        ' start timer
        tmrAsk.Interval = 1000
        Enable tmrAsk
    End If
    If lProgress > 0 Then
        barTimeout.Min = 0
        barTimeout.Max = lProgress
        barTimeout.Value = 0
        barTimeout.Visible = True
        barTimeout.Top = lHeight
        lblTimeout.Top = lHeight
        lblTimeout.Caption = "Progress:"
        lblTimeout.Visible = True
        lHeight = lblTimeout.Top + lblTimeout.Height + lHorizSpace
    End If
    
    ' Don't Ask
    If bShowDontAsk Then
        chkDontAsk.Visible = True
        chkDontAsk.Top = lHeight
        lHeight = chkDontAsk.Top + chkDontAsk.Height + lHorizSpace
        If strIcon = "?" Then
            chkDontAsk.Caption = "Don't ask this question again"
        Else
            chkDontAsk.Caption = "Don't show this message again"
        End If
    Else
        chkDontAsk.Visible = False
    End If

    ' Header
    If strHeader = "" Then
        If bNoWait Then
            strHeader = "PROCESSING:  Please Wait ..."
        ElseIf Screen.ActiveForm Is Nothing Then
            strHeader = "Message"
        Else
            ' default form caption (Parm will override).
            strHeader = Screen.ActiveForm.Caption
        End If
    End If
    Me.Caption = strHeader

    If Len(strGetString) > 0 Then
        ' Set focus to input string
        Set ctlFocus = txtInput
        SelectAll txtInput
    End If

    ' Adjust window.
    'Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - lHeight - 500) \ 2, Me.Width, lHeight + 400
    'Me.Move Me.Left, Me.Height, txtInput.Left * 2 + txtInput.Width + (Me.Width - Me.ScaleWidth), _
    '                    lHeight + 400
    Me.Move Me.Left, Me.Height, Me.Width, lHeight + 400
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
End Sub

Private Sub Finish(ctl As Control)

    Dim iPos As Long, strTemp As String

    ' Assign return value.
    If strGetString = "" Then
        iPos = InStr(ctl.Caption, "&")
        strArgs = UCase(Mid(ctl.Caption, iPos + 1, 1))
        If chkDontAsk.Value = vbChecked Then strArgs = strArgs & "-"
    Else
        If ctl.Cancel = True Then
            strArgs = ""
        Else
            ' Check if data type matches.
            Select Case UCase(Left(strGetString, 1))
                Case "N"
                    ' Check for valid number.
                    If Not IsNumeric(txtInput) Then
                        InputError
                        Exit Sub
                    End If
                Case "I"
                    ' Check for valid integer.
                    If Not IsNumeric(txtInput) Or InStr(txtInput, ".") > 0 Then
                        InputError
                        Exit Sub
                    End If
                Case "D"
                    ' Check for valid date.
                    If Not IsDate(txtInput) Then
                        InputError
                        Exit Sub
                    End If
                    ' Return date in standardized format.
                    txtInput = Format$(txtInput)
            End Select
            strArgs = Trim(txtInput)
        End If
    End If

    ' Cleanup.
    Disable tmrAsk
    txtInput = ""
    strGetString = ""
    iTimeout = 0

    Unload Me

End Sub

Private Sub cmdButton_Click(Index As Integer)

    Finish cmdButton(Index)

End Sub

Private Sub ico_Click()

    If iNumButtons = 0 Then
        Finish cmdButton(0)
    End If

End Sub

Private Sub lblMessage_Click()

    If iNumButtons = 0 Then
        Finish cmdButton(0)
    End If

End Sub

Private Sub txtInput_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case UCase(Left(strGetString, 1))
        Case "C"
            ' only get one character
            Finish cmdButton(0)
        Case "D"
            ' If date, show Day label
            lblDay.Caption = WeekdayName(txtInput, False)
            lblDay.Visible = True
    End Select

End Sub

Private Sub Form_Activate()

    If Not ctlFocus Is Nothing Then
        MoveFocus ctlFocus
        Set ctlFocus = Nothing
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    barTimeout.Value = 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    barTimeout.Value = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set ctlFocus = Nothing
    Disable tmrAsk

End Sub

Private Sub InputError()

    Beep
    'ShowIcon "Error"
    MoveFocus txtInput
    SelectAll txtInput

End Sub

Private Sub SetButton(ctl As Control, strButtons$, ByVal iTop As Long)

    Dim iPos&, strCaption$

    ctl.Default = False
    ctl.Cancel = False

    If strButtons <> "" Then
        ' Get next button imbedded in string.
        iPos = InStr(strButtons, "|")
        If iPos > 0 Then
            strCaption = Left(strButtons, iPos - 1)
            strButtons = Mid(strButtons, iPos + 1)
        Else
            strCaption = strButtons
            strButtons = ""
        End If
    End If
    If strCaption = "" Then
        Disable ctl
        ctl.Visible = False
    Else
        iNumButtons = iNumButtons + 1
        ctl.Top = iTop
        ctl.Left = lblTimeout.Left + (cmdButtons.Left - lblTimeout.Left) * (iNumButtons - 1)
        Enable ctl
        ctl.Visible = True
        ctl.Font.Size = 8 '.25  ' 9.75
        ctl.Height = 400
        If Left(strCaption, 1) = "+" Then
            ctl.Default = True
            strCaption = Mid$(strCaption, 2)
            Set ctlFocus = ctl
        End If
        If Left(strCaption, 1) = "-" Then
            ctl.Cancel = True
            strCaption = Mid$(strCaption, 2)
        End If
        iPos = InStr(strCaption, "&")
        If iPos > 0 Then
            ctl.Caption = strCaption
        Else
            ctl.Caption = "&" + strCaption
        End If
    End If

End Sub

Private Sub tmrAsk_Timer()
    
    Dim i As Integer

    If iTimeout > 0 Then
        ' see if timeout has expired
        If barTimeout.Value >= barTimeout.Max Then
            For i = 0 To iNumButtons - 1
                If cmdButton(i).Visible And cmdButton(i).Default Then
                    cmdButton_Click i
                    Exit For
                End If
            Next
        Else
            ' move bar up 1 second
            barTimeout.Value = barTimeout.Value + 1
        End If
    End If

End Sub

Public Sub ShowIcon(strIcon$)

    Dim lIconNum&, hIcon&, lRC&

    If Len(Trim(strIcon)) = 0 Then
        icoUni.Visible = False
    Else
    icoUni.Visible = True
    
    
        'ico.Visible = True
        lIconNum = 0
        Select Case Left(UCase(Trim(strIcon)), 1)
            Case "[", "E"    ' Stop icon
                'DisplayStop strIcon
                'lIconNum = 32513
                Set icoUni.Picture = icoBag.GetPicture(3) 'Stop
            Case "?"
                'lIconNum = 32514
                'Ico.Picture = QuestionIcon.Picture
                Set icoUni.Picture = icoBag.GetPicture(6) '?
            Case "!"
                'lIconNum = 32515
                'Ico.Picture = ExclamationIcon.Picture
                Set icoUni.Picture = icoBag.GetPicture(9)
            Case "H"
                'ico.Picture = icoHappy.Picture
                Set icoUni.Picture = icoBag.GetPicture(4) 'Happy
            Case "I"
                'lIconNum = 32516
                Set icoUni.Picture = icoBag.GetPicture(5) 'Information
            Case "S"
                'ico.Picture = icoSad.Picture
                Set icoUni.Picture = icoBag.GetPicture(7) 'Sad
            Case "T"
                'ico.Picture = icoTimer.Picture
                'Ico.Height = 500
                'Ico.Width = Ico.Height
                Set icoUni.Picture = icoBag.GetPicture(2) 'Timer
            Case "L"
                'ico.Picture = icoLightning.Picture
                
                Set icoUni.Picture = icoBag.GetPicture(8) 'Lightning
            Case "B"
                'ico.Picture = icoBulb.Picture
                
                Set icoUni.Picture = icoBag.GetPicture(1) 'Bulb
            Case Else
            icoUni.Visible = False
            
            
                'ico.Visible = False
        End Select
        'delete
        If lIconNum > 0 Then
            ' show one of the standard icons
            hIcon = LoadIconNum(0, lIconNum)
            If hIcon Then
                'lRC = DrawIcon(ico.hDC, 0, 0, hIcon)
                lRC = DestroyIcon(hIcon)
                
            End If
        End If
    End If

End Sub







