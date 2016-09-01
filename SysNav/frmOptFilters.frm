VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOptFilters 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   435
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5505
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmMDIChild"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   5505
   Begin HexUniControls.ctlUniButtonImageXP cmdClear 
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Top             =   60
      Width           =   450
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
      Caption         =   "frmOptFilters.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmOptFilters.frx":0030
      Style           =   1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmOptFilters.frx":0050
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP Corner 
      Height          =   240
      Left            =   4800
      TabIndex        =   4
      Top             =   210
      Visible         =   0   'False
      Width           =   720
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
      Caption         =   "frmOptFilters.frx":006C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmOptFilters.frx":0098
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmOptFilters.frx":00B8
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   315
      Left            =   4590
      TabIndex        =   3
      Top             =   60
      Width           =   825
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
      Caption         =   "frmOptFilters.frx":00D4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmOptFilters.frx":0100
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmOptFilters.frx":0120
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Height          =   315
      Left            =   3615
      TabIndex        =   2
      Top             =   60
      Width           =   465
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
      Caption         =   "frmOptFilters.frx":013C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmOptFilters.frx":0166
      Style           =   1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmOptFilters.frx":0186
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtOperVal 
      Height          =   315
      Left            =   2175
      TabIndex        =   1
      Top             =   60
      Width           =   1350
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmOptFilters.frx":01A2
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
      Alignment       =   1
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmOptFilters.frx":01C2
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptFilters.frx":01E2
   End
   Begin HexUniControls.ctlUniComboBoxXP txtOper 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      Tip             =   "frmOptFilters.frx":01FE
      Sorted          =   0   'False
      HScroll         =   0   'False
      Style           =   2
      ButtonBackColor =   -2147483633
      ButtonForeColor =   0
      ButtonWidth     =   17
      Locked          =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      TrapTab         =   0   'False
      ButtonStyle     =   -1
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmOptFilters.frx":021E
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
      MaxLength       =   0
      RightToLeft     =   0   'False
      LeftMargin      =   0
      RightMargin     =   0
      SelectOnFocus   =   0   'False
   End
End
Attribute VB_Name = "frmOptFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    strFormat As String
    bOK As Boolean
End Type
Private m As mPrivate

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptFilters.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdClear_Click()
On Error GoTo ErrSection:

    txtOperVal.Text = ""
    txtOper.Text = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptFilters.cmdClear.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptFilters.cmdOK.Click", eGDRaiseError_Show
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
    RaiseError "frmOptFilters.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
    
    g.Styler.StyleForm Me
    
    'Load operators
    txtOper.AddItem "Not Equal"
    txtOper.AddItem "Greater than or Equal to"
    txtOper.AddItem "Greater than"
    txtOper.AddItem "Less than or Equal to"
    txtOper.AddItem "Less than"
    txtOper.AddItem "Equal to"
    txtOper.Text = "Greater than"
    txtOperVal.Text = 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptFilters.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtOperVal_LostFocus()
On Error GoTo ErrSection:
    
    txtOperVal.Text = Format(txtOperVal.Text, m.strFormat)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptFilters.txtOperVal.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Public Function ShowMe(ByVal X As Long, ByVal Y As Long, strStat As String, _
                        strOper As String, strOperVal As String, strFormat As String) As Boolean
On Error GoTo ErrSection:

    Move X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
                
    m.strFormat = strFormat
            
    If strOper = "N" Or strOper = "" Then
        txtOper = "Greater than"
        txtOperVal = 0
    Else
        txtOper.Text = strOper
        txtOperVal.Text = Format(strOperVal, m.strFormat)
    End If
    
    Select Case strOper
        Case "<>": txtOper.Text = "Not Equal"
        Case ">=": txtOper.Text = "Greater than or Equal to"
        Case ">": txtOper.Text = "Greater than"
        Case "<=": txtOper.Text = "Less than or Equal to"
        Case "<": txtOper.Text = "Less than"
        Case "=": txtOper.Text = "Equal to"
        Case Else: txtOper.Text = ""
    End Select
    
    ShowForm Me, True
    
    If m.bOK Then
        If frmOptCustomize.Visible Then
            Select Case txtOper.Text
                Case "Not Equal": strOper = "<>"
                Case "Greater than or Equal to": strOper = ">="
                Case "Greater than": strOper = ">"
                Case "Less than or Equal to": strOper = "<="
                Case "Less than": strOper = "<"
                Case "Equal to": strOper = "="
                Case Else: strOper = ""
            End Select
            strOperVal = txtOperVal.Text
        End If
    End If
    
    ShowMe = m.bOK
    Unload Me
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptFilters.ShowMe", eGDRaiseError_Raise
    
End Function

