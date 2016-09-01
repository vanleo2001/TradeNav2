VERSION 5.00
Begin VB.Form frmTheme 
   Appearance      =   0  'Flat
   Caption         =   "Select Color Theme"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14400
   ForeColor       =   &H80000009&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   14400
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
Begin HexUniControls.ctlUniButtonImageXP cmdDisplay
      Caption         =   "Test Selection"
      Height          =   465
      Left            =   11280
      TabIndex        =   6
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
Begin HexUniControls.ctlUniButtonImageXP cmdOK
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   12960
      TabIndex        =   5
      Top             =   5280
      Width           =   1095
   End
Begin HexUniControls.ctlUniRadioXP opt2
Pressed = 0
      Caption         =   "IVORY"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   6413
      TabIndex        =   2
      Top             =   4800
      Width           =   1575
   End
Begin HexUniControls.ctlUniRadioXP opt1
      Caption         =   "CLASSIC"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   4800
Pressed=-1
      Width           =   1455
   End
Begin HexUniControls.ctlUniRadioXP opt3
Pressed = 0
      Caption         =   "CHARCOAL"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   11040
      TabIndex        =   0
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4530
      Left            =   9600
      Picture         =   "frmTheme.frx":0000
      Top             =   240
      Width           =   4530
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4530
      Left            =   240
      Picture         =   "frmTheme.frx":AA58
      Top             =   240
      Width           =   4530
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4530
      Left            =   4920
      Picture         =   "frmTheme.frx":15292
      Top             =   240
      Width           =   4530
   End
Begin HexUniControls.ctlUniLabelXP Label3
      BackStyle       =   0  'Transparent
      Caption         =   "(the Theme can be changed at any time in the Program Settings)"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   5355
      Width           =   6015
   End
Begin HexUniControls.ctlUniLabelXP Label1
      BackStyle       =   0  'Transparent
      Caption         =   "Select which Color Theme you prefer"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   5355
      Width           =   3975
   End
End
Attribute VB_Name = "frmTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_bResetWhenUnload As Boolean

Public Function ShowMe(Optional ByVal bStartup As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strTheme$
    
    m_bResetWhenUnload = False

    If Not IsAtLeastVista Then
        strTheme = "Classic"
    Else
        SetWindowTheme Me.hWnd, "", 0
        Me.Icon = Picture16("kBlank")
        
        CenterTheForm Me
        
        ' get current theme
        If g.nColorTheme = vbWhite Then
            opt2.Value = True
        ElseIf g.nColorTheme = kDarkThemeColor Then
            opt3.Value = True
        Else
            opt1.Value = True
        End If
        SetFormColors
        
        ' change labels if not at startup
        If Not bStartup Then
            'Label1.Caption = "Select which Theme you prefer:"
            'Label3.Caption = ""
            'cmdDisplay.Visible = True
        End If
        
        ' show form
        Me.Show vbModal
        
        ' get selected theme
        If opt2.Value Then
            strTheme = "Ivory"
        ElseIf opt3.Value Then
            strTheme = "Charcoal"
        Else
            strTheme = "Classic"
        End If
        
        ' if changing themes, store the INI settings
        ' (but check against what's in the INI file since they could change their mind before restarting)
        If strTheme <> GetIniFileProperty("TradenavTheme", "", "General", g.strIniFile) Then
            SetIniFileProperty "TradenavTheme", strTheme, "General", g.strIniFile
            Select Case UCase(strTheme)
            Case "IVORY"
                SetIniFileProperty "ToolbarIconStyle", 1, "Toolbars", g.strIniFile
                SetIniFileProperty "ToolbarSkin", eTbSkin_LightFlat, "Toolbars", g.strIniFile
            Case "CHARCOAL"
                SetIniFileProperty "ToolbarIconStyle", 1, "Toolbars", g.strIniFile
                SetIniFileProperty "ToolbarSkin", eTbSkin_DarkFlat, "Toolbars", g.strIniFile
            Case Else
                SetIniFileProperty "ToolbarIconStyle", 0, "Toolbars", g.strIniFile
                SetIniFileProperty "ToolbarSkin", eTbSkin_Silver, "Toolbars", g.strIniFile
            End Select
            If Not bStartup Then
                ' otherwise will take effect the next time TradeNav starts
                InfBox "The new theme will take effect the next time Trade Navigator starts.", "I"
            End If
        End If
    End If
    
    ShowMe = strTheme

ErrExit:
    Unload Me
    Exit Function

ErrSection:
    RaiseError "frmTheme.ShowMe"

End Function

Private Sub cmdDisplay_Click()

    SetFormColors

End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    Dim i&
    
    If m_bResetWhenUnload Then
        'restore current theme colors
        If g.nColorTheme = kDarkThemeColor Then
            i = geHighContrastOn(frmMain.hWnd, kDarkThemeColor, vbWhite)
            SendMessage frmMain.hWnd, WM_THEMECHANGED, 0, 0
        ElseIf g.nColorTheme = vbWhite Then
            i = geHighContrastOn(frmMain.hWnd, vbWhite, vbBlack)
        Else
            i = geHighContrastOn(frmMain.hWnd, -1, -1)
        End If
    End If

End Sub

Private Sub Image1_Click()

    opt1.Value = True

End Sub

Private Sub Image2_Click()
    
    opt2.Value = True

End Sub

Private Sub Image3_Click()

    opt3.Value = True

End Sub

Private Sub opt1_Click()
    'JM 01-14-2-16: the display button not visible means user has never explicity selected a theme
    '               since this will be the only form visible at this point, the call to test theme colors will be quick
    'If Not cmdDisplay.Visible Then cmdDisplay_Click
    SetFormColors
End Sub

Private Sub opt2_Click()
    'If Not cmdDisplay.Visible Then cmdDisplay_Click
    SetFormColors
End Sub

Private Sub opt3_Click()
    'If Not cmdDisplay.Visible Then cmdDisplay_Click
    SetFormColors
End Sub

Private Sub SetFormColors()
        
    On Error Resume Next
        
    Dim i&, c&
    
    ' first set all the colors very quickly
    If opt3.Value Then
        Me.BackColor = kDarkThemeColor
        c = vbWhite
        Image1.Appearance = 0
        Image2.Appearance = 0
        Image3.Appearance = 1
    Else
        If opt2.Value Then
            Me.BackColor = vbWhite
            c = vbBlack
            Image1.Appearance = 0
            Image2.Appearance = 1
            Image3.Appearance = 0
        Else
            Me.BackColor = &H8000000F ' button face color
            c = &H80000012  ' button text color
            Image1.Appearance = 1
            Image2.Appearance = 0
            Image3.Appearance = 0
        End If
    End If
    Label1.ForeColor = c
    Label3.ForeColor = c
    opt1.ForeColor = c
    opt2.ForeColor = c
    opt3.ForeColor = c
    opt1.BackColor = Me.BackColor
    opt2.BackColor = Me.BackColor
    opt3.BackColor = Me.BackColor
    Me.Refresh

    
    ' then if want to actually set the theme (but this is a bit time consuming)
    If Me.Visible And cmdDisplay.Visible Then
        'explicitly check value of each radio button, don't want a default catch all else clause
        If opt1.Value Then
            i = geHighContrastOn(frmMain.hWnd, -1, -1)
            m_bResetWhenUnload = True
        ElseIf opt2.Value Then
            i = geHighContrastOn(frmMain.hWnd, vbWhite, vbBlack)
            m_bResetWhenUnload = True
        ElseIf opt3.Value Then
            i = geHighContrastOn(frmMain.hWnd, kDarkThemeColor, vbWhite)
            m_bResetWhenUnload = True
        End If
        i = Screen.MousePointer
        Screen.MousePointer = vbHourglass
        DoEvents
        Screen.MousePointer = i
    End If

End Sub

