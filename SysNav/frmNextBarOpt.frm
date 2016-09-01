VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmNextBarOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orders for Next Bar"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   1215
      Left            =   180
      TabIndex        =   0
      Top             =   1140
      Width           =   4035
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
      Caption         =   "frmNextBarOpt.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmNextBarOpt.frx":0040
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmNextBarOpt.frx":0060
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkVerifyFrom 
         Height          =   225
         Left            =   3060
         TabIndex        =   4
         Top             =   210
         Visible         =   0   'False
         Width           =   2835
         _ExtentX        =   5001
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
         Caption         =   "frmNextBarOpt.frx":007C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmNextBarOpt.frx":00E0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNextBarOpt.frx":0100
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkIgnoreOpen 
         Height          =   255
         Left            =   300
         TabIndex        =   6
         Top             =   780
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
         Caption         =   "frmNextBarOpt.frx":011C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmNextBarOpt.frx":018A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNextBarOpt.frx":01AA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkNeutral 
         Height          =   255
         Left            =   300
         TabIndex        =   5
         Top             =   480
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
         Caption         =   "frmNextBarOpt.frx":01C6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmNextBarOpt.frx":022C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNextBarOpt.frx":024C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   225
         Left            =   420
         Top             =   210
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
         Caption         =   "frmNextBarOpt.frx":0268
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmNextBarOpt.frx":02D4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNextBarOpt.frx":02F4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   600
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
      Caption         =   "frmNextBarOpt.frx":0310
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmNextBarOpt.frx":033E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmNextBarOpt.frx":035E
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   120
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
      Caption         =   "frmNextBarOpt.frx":037A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmNextBarOpt.frx":03A0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmNextBarOpt.frx":03C0
      RightToLeft     =   0   'False
   End
   Begin gdOCX.gdSelectDate DateNextBar 
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      AllowWeekends   =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   255
      Left            =   180
      Top             =   180
      Width           =   1515
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
      Caption         =   "frmNextBarOpt.frx":03DC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmNextBarOpt.frx":0426
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmNextBarOpt.frx":0446
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmNextBarOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmNextBarOpt.frm
'' Description: Allow the user to change options for the next bar report
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
'' Function:    chkVerifyFrom_Click
'' Description: If the Verify From is on then turn the Ignore Open on
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkVerifyFrom_Click()
On Error GoTo ErrSection:

    If chkVerifyFrom Then chkIgnoreOpen = 1
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBarOpt.chkVerifyFrom.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form without saving
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
    RaiseError "frmNextBarOpt.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Unload the form and save the changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBarOpt.cmdOK.Click", eGDRaiseError_Show
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
    RaiseError "frmNextBarOpt.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form and controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    chkVerifyFrom.Visible = False
    chkVerifyFrom.Left = chkNeutral.Left
    chkVerifyFrom.Value = 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBarOpt.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Intialize and show the form
'' Inputs:      Date, Whether to Show Time
'' Returns:     True if OK, False if Cancelled
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(dDate As Double, ByVal bShowTime As Boolean, bVerifyFrom As Boolean, _
                        bAssumeNoPosition As Boolean, bIgnoreNextBar As Boolean) As Boolean
On Error GoTo ErrSection:

    If bShowTime Then
        DateNextBar.ShowTime = HourMinute
        DateNextBar.Width = 2895
    Else
        DateNextBar.ShowTime = NoTime
        DateNextBar.Width = 2235
    End If
    
    chkVerifyFrom.Visible = FileExist(AddSlash(App.Path) & "VerifyNBR")
    DateNextBar.Value = dDate
    
    ShowForm Me, True
    
    If m.bOK Then
        dDate = DateNextBar.Value
        bAssumeNoPosition = (chkNeutral = vbChecked)
        bIgnoreNextBar = (chkIgnoreOpen = vbChecked)
        bVerifyFrom = (chkVerifyFrom = vbChecked)
    End If

ErrExit:
    ShowMe = m.bOK
    Exit Function
    
ErrSection:
    RaiseError "frmNextBarOpt.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hits the 'X', unload the form without saving
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBarOpt.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

