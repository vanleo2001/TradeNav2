VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAutoBreakout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto Breakout Bar Settings"
   ClientHeight    =   2235
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniCheckXP chkHide 
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1860
      Width           =   2655
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
      Caption         =   "frmAutoBreakout.frx":0000
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmAutoBreakout.frx":0062
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmAutoBreakout.frx":0082
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtNumDays 
      Height          =   315
      Left            =   2700
      TabIndex        =   4
      Top             =   660
      Width           =   555
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmAutoBreakout.frx":009E
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
      Tip             =   "frmAutoBreakout.frx":00C2
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAutoBreakout.frx":00E2
   End
   Begin HexUniControls.ctlUniTextBoxXP txtMaxBars 
      Height          =   315
      Left            =   2700
      TabIndex        =   3
      Top             =   180
      Width           =   555
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmAutoBreakout.frx":00FE
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
      Tip             =   "frmAutoBreakout.frx":0120
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAutoBreakout.frx":0140
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1260
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
      Caption         =   "frmAutoBreakout.frx":015C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmAutoBreakout.frx":018C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmAutoBreakout.frx":01AC
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   540
      TabIndex        =   0
      Top             =   1260
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
      Caption         =   "frmAutoBreakout.frx":01C8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmAutoBreakout.frx":01F6
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmAutoBreakout.frx":0216
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label2 
      Height          =   255
      Left            =   360
      Top             =   720
      Width           =   2295
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
      Caption         =   "frmAutoBreakout.frx":0232
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAutoBreakout.frx":028C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAutoBreakout.frx":02AC
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   255
      Left            =   360
      Top             =   240
      Width           =   2115
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
      Caption         =   "frmAutoBreakout.frx":02C8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAutoBreakout.frx":0320
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAutoBreakout.frx":0340
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmAutoBreakout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Me.Tag = ""
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Me.Tag = "OK"
    Me.Hide

End Sub

Private Sub Form_Load()

    Me.Icon = Picture16("kBlank")
    CenterTheForm Me
    
    g.Styler.StyleForm Me

End Sub

Public Function ShowMe(ByVal strSettings$) As String
On Error GoTo ErrSection:

    Dim i&
    
    i = Val(Parse(strSettings, vbTab, 1))
    If i <= 0 Then i = 8 '6 ' default
    txtMaxBars = Str(i)
    
    i = Val(Parse(strSettings, vbTab, 2))
    If i <= 0 Then i = 3 '20 ' default
    txtNumDays = Str(i)

    chkHide.Value = Abs(GetIniFileProperty("HideAutoBreakout#", False, "General", g.strIniFile))

    ' TLB 5/8/2013: ONLY show this form if special flag file exists
    If Trim(UCase(FileToString(App.Path & "\AutoBreakout.flg", , True))) = "PROJECTX" Then
        Me.Tag = ""
        ShowForm Me, eForm_Modal
        If Len(Me.Tag) > 0 Then
            If Val(txtMaxBars) > 0 And Val(txtNumDays) > 0 Then
                ShowMe = Str(txtMaxBars) & vbTab & Str(txtNumDays)
            End If
        End If
    ' otherwise just toggle the AutoBreakout ON/OFF using the default settings
    ElseIf g.FractZen.Allowed And Len(strSettings) = 0 Then
        ShowMe = Str(txtMaxBars) & vbTab & Str(txtNumDays)
    End If
    
    If g.bHideAutoBreakoutNumber <> -(chkHide.Value) Then
        g.bHideAutoBreakoutNumber = -(chkHide.Value)
        SetIniFileProperty "HideAutoBreakout#", g.bHideAutoBreakoutNumber, "General", g.strIniFile
    End If
    
    Unload Me

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoBreakout.ShowMe"
End Function

