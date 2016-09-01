VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmNotes 
   Caption         =   "Notes"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   3615
      Left            =   8820
      TabIndex        =   1
      Top             =   100
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
      Caption         =   "frmNotes.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmNotes.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmNotes.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   420
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1200
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
         Caption         =   "frmNotes.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmNotes.frx":008E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmNotes.frx":00AE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   420
         Left            =   0
         TabIndex        =   2
         Top             =   555
         Width           =   1200
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
         Caption         =   "frmNotes.frx":00CA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmNotes.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmNotes.frx":0118
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txtNotes 
      Height          =   5685
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   8580
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmNotes.frx":0134
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
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   2
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmNotes.frx":0154
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmNotes.frx":0174
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmNotes.frm
'' Description: Allow the user to enter in notes about a system
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
    RaiseError "frmNotes.cmdCancel.Click", eGDRaiseError_Show
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
    RaiseError "frmNotes.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated, move the focus to the notes box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:
    
    MoveFocus txtNotes

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNotes.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Show the help form if the user presses F1
'' Inputs:      Code of Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNotes.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Intialize the form and controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Dim strText As String               ' Placement for the form from the ini file
    
    g.Styler.StyleForm Me
    
    Icon = Picture16(ToolbarIcon("ID_News"))
    strText = GetIniFileProperty("SysNotes", "", "Placement", g.strIniFile)
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText, "LHTW"
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmNotes.Form.Load", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Unload the form without saving if the user hits the 'X'
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
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
    RaiseError "frmNotes.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move the controls as the form resizes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    If LimitFormSize(Me, fraButtons.Width * 3, fraButtons.Height + (fraButtons.Top * 2)) Then Exit Sub
    
    With fraButtons
        .Move ScaleWidth - .Width - txtNotes.Left
    End With
    
    With txtNotes
        .Move .Left, .Top, ScaleWidth - fraButtons.Width - (.Left * 3), ScaleHeight - (.Top * 2)
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save the Placement and size of the form on unload
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "SysNotes", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNotes.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Notes to Show
'' Returns:     New Notes on OK, Old Notes on Cancel
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal strNotes As String, Optional ByVal strCaption As String = "Notes") As String
On Error GoTo ErrSection:

    Caption = strCaption
    txtNotes.Text = strNotes
    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK Then
        ShowMe = txtNotes.Text
    Else
        ShowMe = strNotes
    End If
    
ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmNotes.ShowMe", eGDRaiseError_Raise

End Function

