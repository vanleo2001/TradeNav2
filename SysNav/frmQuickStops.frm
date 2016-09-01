VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmQuickStops 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quick Stops"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniCheckXP chkTrailingStop 
      Height          =   220
      Left            =   420
      TabIndex        =   4
      Top             =   1200
      Width           =   1440
      _ExtentX        =   2540
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
      Caption         =   "frmQuickStops.frx":0000
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   -1  'True
      Tip             =   "frmQuickStops.frx":003A
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmQuickStops.frx":005A
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkStopLoss 
      Height          =   220
      Left            =   420
      TabIndex        =   7
      Top             =   180
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "frmQuickStops.frx":0076
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   -1  'True
      Tip             =   "frmQuickStops.frx":00A8
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmQuickStops.frx":00C8
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkProfitTarget 
      Height          =   220
      Left            =   435
      TabIndex        =   9
      Top             =   2445
      Width           =   1485
      _ExtentX        =   2619
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
      Caption         =   "frmQuickStops.frx":00E4
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   -1  'True
      Tip             =   "frmQuickStops.frx":011E
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmQuickStops.frx":013E
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL Frame3 
      Height          =   1110
      Left            =   225
      TabIndex        =   8
      Top             =   1200
      Width           =   6870
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
      Caption         =   "frmQuickStops.frx":015A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQuickStops.frx":0186
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuickStops.frx":01A6
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtTrailingStop 
         Height          =   360
         Left            =   195
         TabIndex        =   10
         Top             =   360
         Width           =   1515
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmQuickStops.frx":01C2
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
         Tip             =   "frmQuickStops.frx":01E2
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuickStops.frx":0202
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   675
         Index           =   3
         Left            =   1845
         Top             =   330
         Width           =   4740
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
         Caption         =   "frmQuickStops.frx":021E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuickStops.frx":0398
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuickStops.frx":03B8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL Frame2 
      Height          =   900
      Left            =   225
      TabIndex        =   5
      Top             =   180
      Width           =   6870
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
      Caption         =   "frmQuickStops.frx":03D4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQuickStops.frx":0400
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuickStops.frx":0420
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtStopLoss 
         Height          =   360
         Left            =   195
         TabIndex        =   6
         Top             =   345
         Width           =   1515
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmQuickStops.frx":043C
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
         Tip             =   "frmQuickStops.frx":045C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuickStops.frx":047C
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   465
         Index           =   1
         Left            =   1830
         Top             =   345
         Width           =   4740
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
         Caption         =   "frmQuickStops.frx":0498
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuickStops.frx":05AA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuickStops.frx":05CA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   900
      Left            =   240
      TabIndex        =   2
      Top             =   2415
      Width           =   6870
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
      Caption         =   "frmQuickStops.frx":05E6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQuickStops.frx":0612
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuickStops.frx":0632
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtProfitTarget 
         Height          =   360
         Left            =   195
         TabIndex        =   3
         Top             =   360
         Width           =   1515
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmQuickStops.frx":064E
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
         Tip             =   "frmQuickStops.frx":066E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuickStops.frx":068E
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   420
         Index           =   0
         Left            =   1815
         Top             =   345
         Width           =   4900
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
         Caption         =   "frmQuickStops.frx":06AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuickStops.frx":07F4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuickStops.frx":0814
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Height          =   420
      Left            =   7320
      TabIndex        =   1
      Top             =   270
      Width           =   1140
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
      Caption         =   "frmQuickStops.frx":0830
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmQuickStops.frx":0854
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmQuickStops.frx":0874
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7320
      TabIndex        =   0
      Top             =   825
      Width           =   1140
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
      Caption         =   "frmQuickStops.frx":0890
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmQuickStops.frx":08BC
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmQuickStops.frx":08DC
      RightToLeft     =   0   'False
   End
End
Attribute VB_Name = "frmQuickStops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmQuickStops.frm
'' Description: Allow the user to add Stop Loss/Profit Target rules to system
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkProfitTarget_Click
'' Description: Enable/Disable the Profit Target Text Box appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkProfitTarget_Click()
On Error GoTo ErrSection:

    Enable txtProfitTarget, chkProfitTarget = vbChecked

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.chkProfitTarget.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkStopLoss_Click
'' Description: Enable/Disable the Stop Loss Text Box appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkStopLoss_Click()
On Error GoTo ErrSection:

    Enable txtStopLoss, chkStopLoss = vbChecked

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.chkStopLoss.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkTrailingStop_Click
'' Description: Enable/Disable the Trailing Stop Text Box appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkTrailingStop_Click()
On Error GoTo ErrSection:

    Enable txtTrailingStop, chkTrailingStop = vbChecked

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.chkTrailingStop.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Hide the form and return the results
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    MoveFocus cmdOK
    m.bOK = True
    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.cmdOk.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Hide the form and don't return the results
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    MoveFocus cmdCancel
    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.cmdCancel.Click", eGDRaiseError_Show
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
    RaiseError "frmQuickStops.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize and size the controls and the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error Resume Next
    Dim frm As Form
    Set frm = Screen.ActiveForm
    If frm Is Nothing Then
        CenterTheForm Me
    Else
        Me.Move frm.Left + (frm.Width - Me.Width) / 2, frm.Top + frm.Height - frm.ScaleHeight
        Set frm = Nothing
    End If
    
    g.Styler.StyleForm Me
    
    On Error GoTo ErrSection:
    Me.Icon = Picture16(ToolbarIcon("kSelect"))

    txtProfitTarget.Text = Format(GetIniFileProperty("DefaultTargetProfit", 0, _
            "Systems", g.strIniFile), "#,##0.00")
    txtStopLoss.Text = Format(GetIniFileProperty("DefaultStopLoss", 0, _
            "Systems", g.strIniFile), "#,##0.00")
    txtTrailingStop.Text = Format(GetIniFileProperty("DefaultTrailingStop", 0, _
            "Systems", g.strIniFile), "#,##0.00")
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hits the 'X', unload the form and don't return results
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        m.bOK = False
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save the setting to the INI file upon unloading the form
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:
    
    If ValOfText(txtProfitTarget) <> 0 Then
        SetIniFileProperty "DefaultTargetProfit", ValOfText(txtProfitTarget), "Systems", g.strIniFile
    End If
    If ValOfText(txtStopLoss) <> 0 Then
        SetIniFileProperty "DefaultStopLoss", ValOfText(txtStopLoss), "Systems", g.strIniFile
    End If
    If ValOfText(txtTrailingStop) <> 0 Then
        SetIniFileProperty "DefaultTrailingStop", ValOfText(txtTrailingStop), "Systems", g.strIniFile
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStopLoss_LostFocus
'' Description: Format what the user entered upon the control losing focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStopLoss_LostFocus()
On Error GoTo ErrSection:

    txtStopLoss = Format(txtStopLoss, "#,###.00")
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.txtStopLoss.LostFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStopLoss_Validate
'' Description: Validate and format what the user entered
'' Inputs:      Whether to Cancel the Update
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStopLoss_Validate(Cancel As Boolean)
On Error GoTo ErrSection:
    
    If Not IsNumeric(txtStopLoss.Text) Then
        Cancel = True
        Err.Raise vbObjectError + 1000, , "Please enter in a valid Numeric Value"
    ElseIf ValOfText(txtStopLoss.Text) < 0 Then
        txtStopLoss.Text = Format(ValOfText(txtStopLoss.Text) * -1, "#,##0.00")
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.txtStopLoss.Validate", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTrailingStop_LostFocus
'' Description: Format what the user entered upon the control losing focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTrailingStop_LostFocus()
On Error GoTo ErrSection:

    txtTrailingStop = Format(txtTrailingStop, "#,###.00")
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.txtTrailingStop.LostFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTrailingStop_Validate
'' Description: Validate and format what the user entered
'' Inputs:      Whether to Cancel the Update
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTrailingStop_Validate(Cancel As Boolean)
On Error GoTo ErrSection:
    
    If Not IsNumeric(txtTrailingStop.Text) Then
        Cancel = True
        Err.Raise vbObjectError + 1000, , "Please enter in a valid Numeric Value"
    ElseIf ValOfText(txtTrailingStop.Text) < 0 Then
        txtTrailingStop.Text = Format(ValOfText(txtTrailingStop.Text) * -1, "#,##0.00")
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.txtTrailingStop.Validate", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtProfitTarget_LostFocus
'' Description: Format what the user entered upon the control losing focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtProfitTarget_LostFocus()
On Error GoTo ErrSection:

    txtProfitTarget = Format(txtProfitTarget, "#,###.00")
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.txtProfitTarget.LostFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtProfitTarget_Validate
'' Description: Validate and format what the user entered
'' Inputs:      Whether to Cancel the Update
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtProfitTarget_Validate(Cancel As Boolean)
On Error GoTo ErrSection:
    
    If Not IsNumeric(txtProfitTarget.Text) Then
        Cancel = True
        Err.Raise vbObjectError + 1000, , "Please enter in a valid Numeric Value"
    ElseIf ValOfText(txtProfitTarget.Text) < 0 Then
        txtProfitTarget.Text = Format(ValOfText(txtProfitTarget.Text) * -1, "#,##0.00")
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuickStops.txtProfitTarget.Validate", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and Show the form
'' Inputs:      Profit Target, Stop Loss, Trailing Stop
'' Returns:     True if the user pressed OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(dProfitTarget As Double, dStopLoss As Double, dTrailingStop As Double) As Boolean
On Error GoTo ErrSection:

    If dProfitTarget <> 0 Then
        txtProfitTarget.Text = Format(dProfitTarget, "#,##0.00")
        chkProfitTarget.Value = vbChecked
    Else
        chkProfitTarget.Value = vbUnchecked
    End If
    
    If dStopLoss <> 0 Then
        txtStopLoss.Text = Format(dStopLoss, "#,##0.00")
        chkStopLoss.Value = vbChecked
    Else
        chkStopLoss.Value = vbUnchecked
    End If
    
    If dTrailingStop <> 0 Then
        txtTrailingStop.Text = Format(dTrailingStop, "#,##0.00")
        chkTrailingStop.Value = vbChecked
    Else
        chkTrailingStop.Value = vbUnchecked
    End If
    
    ' If all quick stops are unchecked them check the stop loss by default
    If chkStopLoss = vbUnchecked And chkTrailingStop = vbUnchecked And chkProfitTarget = vbUnchecked Then
        chkStopLoss = vbChecked
    End If
    
    ShowForm Me, True
    
    If m.bOK Then
        dProfitTarget = 0#
        dStopLoss = 0#
        dTrailingStop = 0#
        
        If chkProfitTarget = vbChecked Then dProfitTarget = ValOfText(txtProfitTarget.Text)
        If chkStopLoss = vbChecked Then dStopLoss = ValOfText(txtStopLoss.Text)
        If chkTrailingStop = vbChecked Then dTrailingStop = ValOfText(txtTrailingStop.Text)
    End If
    
    ShowMe = m.bOK
    Unload Me
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuickStops.ShowMe", eGDRaiseError_Show
    Resume ErrExit

End Function

