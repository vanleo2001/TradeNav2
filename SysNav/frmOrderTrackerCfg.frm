VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOrderTrackerCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraAlert 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
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
      Caption         =   "frmOrderTrackerCfg.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOrderTrackerCfg.frx":004A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderTrackerCfg.frx":006A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdPlay 
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   915
         Width           =   615
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
         Caption         =   "frmOrderTrackerCfg.frx":0086
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderTrackerCfg.frx":00B0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderTrackerCfg.frx":00D0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdBrowse 
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   915
         Width           =   855
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
         Caption         =   "frmOrderTrackerCfg.frx":00EC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderTrackerCfg.frx":011A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderTrackerCfg.frx":013A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtSoundFile 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   960
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmOrderTrackerCfg.frx":0156
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
         Tip             =   "frmOrderTrackerCfg.frx":0176
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOrderTrackerCfg.frx":0196
      End
      Begin HexUniControls.ctlUniCheckXP chkAlarm 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1935
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
         Caption         =   "frmOrderTrackerCfg.frx":01B2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOrderTrackerCfg.frx":01FC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOrderTrackerCfg.frx":021C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkPopUp 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
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
         Caption         =   "frmOrderTrackerCfg.frx":0238
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOrderTrackerCfg.frx":027A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOrderTrackerCfg.frx":029A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   1208
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
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
      Caption         =   "frmOrderTrackerCfg.frx":02B6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOrderTrackerCfg.frx":02E2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderTrackerCfg.frx":0302
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   0
         Width           =   1215
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
         Caption         =   "frmOrderTrackerCfg.frx":031E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderTrackerCfg.frx":034C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderTrackerCfg.frx":036C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1215
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
         Caption         =   "frmOrderTrackerCfg.frx":0388
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderTrackerCfg.frx":03AE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderTrackerCfg.frx":03CE
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmOrderTrackerCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmOrderTrackerCfg.frm
'' Description: Allow the user to configure alerts for the order tracker
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
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe() As Boolean
On Error GoTo ErrSection:

    chkPopUp.Value = GetIniFileProperty("PopUp", vbUnchecked, "OrderTracker", g.strIniFile)
    chkAlarm.Value = GetIniFileProperty("Sound", vbUnchecked, "OrderTracker", g.strIniFile)
    txtSoundFile.Text = GetIniFileProperty("SoundFile", "", "OrderTracker", g.strIniFile)
    EnableControls
    
    ShowForm Me, eForm_Modal
    PlaySoundFile ' to stop sound
    
    If m.bOK = True Then
        SetIniFileProperty "PopUp", chkPopUp.Value, "OrderTracker", g.strIniFile
        SetIniFileProperty "Sound", chkAlarm.Value, "OrderTracker", g.strIniFile
        SetIniFileProperty "SoundFile", Trim(txtSoundFile.Text), "OrderTracker", g.strIniFile
    End If

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmOrderTrackerCfg.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdBrowse_Click
'' Description: Allow the user to browse for a sound file to play
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdBrowse_Click()
On Error GoTo ErrSection:

    Dim strInit$

    strInit = txtSoundFile.Text
    If Len(strInit) = 0 Then
        strInit = AddSlash(WindowsPath) & "Media"
    End If
    
    txtSoundFile.Text = CommonDialogFile(frmMain.CommonDialog1, False, "Wave Files (*.wav)|*.wav", strInit)
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderTrackerCfg.cmdBrowse.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form without saving anything
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
    RaiseError "frmOrderTrackerCfg.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Save the user's changes and unload
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If chkAlarm.Value = vbChecked And Len(Trim(txtSoundFile.Text)) = 0 Then
        MoveFocus txtSoundFile
        InfBox "Please specify a sound file to play", "!", , "Error"
        GoTo ErrExit
    End If

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderTrackerCfg.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPlay_Click
'' Description: Play a preview of the sound selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPlay_Click()
On Error GoTo ErrSection:

    PlaySoundFile txtSoundFile.Text

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderTrackerCfg.cmdPlay.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Caption = "Configuration for the Order Tracker"
    Me.Icon = Picture16(ToolbarIcon("ID_Alerts"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderTrackerCfg.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Hide the form when the user clicks on the 'X'
'' Inputs:      Whether or not to Cancel Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.bOK = False
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderTracker.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Enable cmdPlay, FileExist(txtSoundFile.Text)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderTracker.EnableControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSoundFile_Change
'' Description: Enable/Disable controls as text changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSoundFile_Change()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderTracker.txtSoundFile.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

