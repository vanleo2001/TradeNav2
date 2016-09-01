VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HexUniControls.ctlUniFrameWL fraStyle 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3300
      Width           =   4215
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
      Caption         =   "frmPassword.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPassword.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPassword.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optForex 
         Height          =   255
         Left            =   2887
         TabIndex        =   4
         Top             =   0
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
         Caption         =   "frmPassword.frx":0068
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmPassword.frx":0094
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":00B4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optBox 
         Height          =   255
         Left            =   2047
         TabIndex        =   6
         Top             =   0
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
         Caption         =   "frmPassword.frx":00D0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmPassword.frx":00F8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":0118
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optGrid 
         Height          =   255
         Left            =   1102
         TabIndex        =   8
         Top             =   0
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
         Caption         =   "frmPassword.frx":0134
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmPassword.frx":015E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":017E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblStyle 
         Height          =   255
         Left            =   382
         Top             =   0
         Width           =   555
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
         Caption         =   "frmPassword.frx":019A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPassword.frx":01C6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":01E6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraChangePassword 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4215
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
      Caption         =   "frmPassword.frx":0202
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPassword.frx":022E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPassword.frx":024E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtConfirmPassword 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   960
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPassword.frx":026A
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
         Tip             =   "frmPassword.frx":028A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":02AA
      End
      Begin HexUniControls.ctlUniTextBoxXP txtNewPassword 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   555
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPassword.frx":02C6
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
         Tip             =   "frmPassword.frx":02E6
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":0306
      End
      Begin HexUniControls.ctlUniTextBoxXP txtOldPassword 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   150
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPassword.frx":0322
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
         Tip             =   "frmPassword.frx":0342
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":0362
      End
      Begin HexUniControls.ctlUniLabelXP lblConfirmPassword 
         Height          =   255
         Left            =   60
         Top             =   990
         Width           =   1335
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
         Caption         =   "frmPassword.frx":037E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPassword.frx":03C0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":03E0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblNewPassword 
         Height          =   255
         Left            =   60
         Top             =   585
         Width           =   1335
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
         Caption         =   "frmPassword.frx":03FC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPassword.frx":0436
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":0456
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblOldPassword 
         Height          =   255
         Left            =   60
         Top             =   180
         Width           =   1335
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
         Caption         =   "frmPassword.frx":0472
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPassword.frx":04B4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":04D4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraPassword 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
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
      Caption         =   "frmPassword.frx":04F0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPassword.frx":051C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPassword.frx":053C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtPassword 
         Height          =   285
         Left            =   0
         TabIndex        =   2
         Top             =   540
         Width           =   4095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPassword.frx":0558
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
         Tip             =   "frmPassword.frx":0578
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":0598
      End
      Begin HexUniControls.ctlUniLabelXP lblPassword 
         Height          =   435
         Left            =   7
         Top             =   0
         Width           =   4095
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
         Caption         =   "frmPassword.frx":05B4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPassword.frx":0608
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":0628
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   375
      Left            =   795
      TabIndex        =   10
      Top             =   1140
      Width           =   2895
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
      Caption         =   "frmPassword.frx":0644
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPassword.frx":0670
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPassword.frx":0690
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1260
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
         Caption         =   "frmPassword.frx":06AC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPassword.frx":06D2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":06F2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   1635
         TabIndex        =   12
         Top             =   0
         Width           =   1260
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
         Caption         =   "frmPassword.frx":070E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPassword.frx":073C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPassword.frx":075C
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmPassword.frm
'' Description: Asks the user for a Password
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Private Type mPrivate
    strPassword As String               ' Password the user typed in
    bOK As Boolean                      ' Did the user hit OK or Cancel?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user hits Cancel, unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:
    
    m.bOK = False
    Me.Hide
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPassword.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit:

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
    RaiseError "frmPassword.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form loads, initialize the necessary variables
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
       
    Screen.MousePointer = vbDefault
    
    'Center the form
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    Me.Icon = Picture16(ToolbarIcon("kSelect"))
        
    txtPassword.Text = GetIniFileProperty("LastPasswordUsed", 0, "Misc", g.strIniFile)
    m.strPassword = txtPassword.Text
    
    txtPassword.SelLength = Len(txtPassword.Text)
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPassword.Form.Load", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user hits OK, return the password
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    If fraChangePassword.Visible = True Then
        If Trim(txtNewPassword.Text) <> Trim(txtConfirmPassword.Text) Then
            Err.Raise vbObjectError + 1000, , "New Password does not match Confirm Password"
        End If
        
        m.strPassword = txtNewPassword.Text
    ElseIf fraStyle.Visible = False Then
        If Len(txtPassword.Text) < 5 Then
            Err.Raise vbObjectError + 1000, , "Password must be 5 or more characters"
        End If
    
        m.strPassword = txtPassword.Text
    Else
        If Len(txtPassword.Text) = 0 Then
            Err.Raise vbObjectError + 1000, , "Please supply a name for the new tab"
        ElseIf UCase(Trim(txtPassword.Text)) = "(FILTER)" Then
            Err.Raise vbObjectError + 1000, , txtPassword.Text & " is a reserved name." & vbCrLf & "Please supply a different name"     '4238
        End If
    End If
    
    m.bOK = True
    Me.Hide
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPassword.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hits the X in the corner, it means they are
''              cancelling out of this form, so set the cancel property to true
'' Inputs:      Whether or not to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
    
    If UnloadMode = vbFormControlMenu Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPassword.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Show the form and return the password if the user hit OK
'' Inputs:      Optional Item to show in the Label
'' Returns:     Password typed in or empty string if cancelled
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Optional ByVal strItem As String = "") As String
On Error GoTo ErrSection:

    If strItem <> "" Then
        lblPassword.Caption = "Please enter the password for:" & vbCrLf & strItem
    End If
    
    fraStyle.Visible = False
    fraChangePassword.Visible = False
    Height = 2040
    
    ShowForm Me, True
    
    ShowMe = ""
    If m.bOK Then
        ShowMe = m.strPassword
    End If
    
ErrExit:
    Unload Me
    Exit Function

ErrSection:
    Unload Me
    RaiseError "frmPassword.ShowMe", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowChange
'' Description: Allow the user to change password with confirmation
'' Inputs:      Old Password
'' Returns:     New Password on OK, Blank String otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowChange(ByVal strOldPassword As String) As String
On Error GoTo ErrSection:

    fraChangePassword.Top = fraPassword.Top
    fraButtons.Top = fraChangePassword.Height + (fraChangePassword.Top * 2)
    fraStyle.Visible = False
    Height = 2640
    Caption = "Change Password"
    
    txtOldPassword.Text = strOldPassword
    txtOldPassword.Enabled = False
    
    ShowForm Me, True
    
    If m.bOK Then ShowChange = m.strPassword

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmPassword.ShowChange", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowNewTab
'' Description: Allow the user to enter in a new tab name and style
'' Inputs:      Name and Style
'' Returns:     True on OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowNewTab(strName As String, bGrid As Boolean, bBox As Boolean) As Boolean
On Error GoTo ErrSection:

    fraChangePassword.Visible = False
    Height = 2040 + fraStyle.Height + fraPassword.Top
    fraStyle.Top = (fraPassword.Top * 2) + fraPassword.Height
    fraButtons.Top = fraStyle.Top + fraPassword.Top + fraStyle.Height
    Caption = "New Tab"
    lblPassword.Caption = "Please enter in a name for the new tab and select the style of quote board you would like it to be"
    txtPassword.PasswordChar = ""
    txtPassword.Text = ""
    
    If bGrid Then
        optGrid.Value = True
    ElseIf bBox Then
        optBox.Value = True
    Else
        optForex.Value = True
    End If
    
    ShowForm Me, True
    
    If m.bOK Then
        strName = Trim(Replace(txtPassword.Text, "|", "/"))
        bGrid = optGrid.Value
        bBox = optBox.Value
    End If
    
    ShowNewTab = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmPassword.ShowNewTab", eGDRaiseError_Raise
    
End Function

