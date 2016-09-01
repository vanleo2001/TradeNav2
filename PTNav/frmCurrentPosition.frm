VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmCurrentPosition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
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
      Caption         =   "frmCurrentPosition.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmCurrentPosition.frx":0026
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmCurrentPosition.frx":0046
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtQuantity 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   1185
      Width           =   975
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmCurrentPosition.frx":0062
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
      Tip             =   "frmCurrentPosition.frx":0082
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCurrentPosition.frx":00A2
   End
   Begin HexUniControls.ctlUniComboImageXP cboDirection 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   1170
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ButtonBackColor =   -2147483633
      ButtonForeColor =   -2147483630
      ButtonStyle     =   -1
      SelectorStyle   =   -1
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
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
      Tip             =   "frmCurrentPosition.frx":00BE
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmCurrentPosition.frx":00DE
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin gdOCX.gdScrollBar sbQty 
      Height          =   360
      Left            =   4080
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1147
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   635
   End
   Begin HexUniControls.ctlUniLabelXP lblQuantity 
      Height          =   255
      Left            =   2400
      Top             =   1200
      Width           =   735
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
      Caption         =   "frmCurrentPosition.frx":00FA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCurrentPosition.frx":012E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCurrentPosition.frx":014E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblDirection 
      Height          =   255
      Left            =   120
      Top             =   1200
      Width           =   735
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
      Caption         =   "frmCurrentPosition.frx":016A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCurrentPosition.frx":01A0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCurrentPosition.frx":01C0
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblDescription 
      Height          =   855
      Left            =   120
      Top             =   120
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
      Caption         =   "frmCurrentPosition.frx":01DC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCurrentPosition.frx":0386
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCurrentPosition.frx":03A6
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmCurrentPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmCurrentPosition.frm
'' Description: Allow the user to enter in their current "real" position
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Qty As cPriceEditor                 ' Quantity editor
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up and show the form
'' Inputs:      Default Current Position
'' Returns:     User's Current Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal lCurrentPosition As Long) As Long
On Error GoTo ErrSection:

    Set m.Qty = New cPriceEditor
    m.Qty.Init sbQty, txtQuantity, Nothing, , , , , True
    
    Select Case lCurrentPosition
        Case Is = 0
            cboDirection.Text = "None"
        Case Is > 0
            cboDirection.Text = "Long"
        Case Is < 0
            cboDirection.Text = "Short"
    End Select
    m.Qty.Price = Abs(lCurrentPosition)
    
    EnableControls
    ShowForm Me, eForm_Modal, frmMain
    
    Select Case cboDirection.Text
        Case "None"
            ShowMe = 0&
        Case "Long"
            ShowMe = m.Qty.Price
        Case "Short"
            ShowMe = m.Qty.Price * -1&
    End Select

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmCurrentPosition.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboDirection_Click
'' Description: Enable/Disable the quantity controls as the combo changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboDirection_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCurrentPosition.cboDirection_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Hide the form and let ShowMe unload
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCurrentPosition.cmdOK_Click"
    
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

    Caption = "Current Position"
    Icon = Picture16("kBlank")
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    With cboDirection
        .AddItem "None"
        .AddItem "Long"
        .AddItem "Short"
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCurrentPosition.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Hide the form and let ShowMe unload it
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCurrentPosition.Form_QueryUnload"
    
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

    Enable lblQuantity, (cboDirection.Text <> "None")
    Enable txtQuantity, (cboDirection.Text <> "None")
    Enable sbQty, (cboDirection.Text <> "None")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCurrentPosition.EnableControls"
    
End Sub

