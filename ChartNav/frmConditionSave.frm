VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmConditionSave 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   435
      Left            =   2865
      TabIndex        =   10
      Top             =   4260
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
      Caption         =   "frmConditionSave.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmConditionSave.frx":002E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmConditionSave.frx":004E
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Height          =   435
      Left            =   1651
      TabIndex        =   9
      Top             =   4260
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
      Caption         =   "frmConditionSave.frx":006A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmConditionSave.frx":0094
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmConditionSave.frx":00B4
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraCopyTo 
      Height          =   4185
      Left            =   38
      TabIndex        =   0
      Top             =   0
      Width           =   5475
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
      Caption         =   "frmConditionSave.frx":00D0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmConditionSave.frx":00F0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmConditionSave.frx":0110
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtAlert 
         Height          =   315
         Left            =   1748
         TabIndex        =   8
         Top             =   2320
         Width           =   3495
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmConditionSave.frx":012C
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
         Tip             =   "frmConditionSave.frx":014C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmConditionSave.frx":016C
      End
      Begin HexUniControls.ctlUniRadioXP optNew 
         Height          =   315
         Index           =   6
         Left            =   548
         TabIndex        =   11
         Top             =   2320
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
         Caption         =   "frmConditionSave.frx":0188
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmConditionSave.frx":01BA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmConditionSave.frx":01DA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optNew 
         Height          =   255
         Index           =   5
         Left            =   548
         TabIndex        =   7
         Top             =   1950
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
         Caption         =   "frmConditionSave.frx":01F6
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmConditionSave.frx":0268
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmConditionSave.frx":0288
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optNew 
         Height          =   315
         Index           =   4
         Left            =   548
         TabIndex        =   6
         Top             =   2750
         Width           =   4575
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
         Caption         =   "frmConditionSave.frx":02A4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmConditionSave.frx":033C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmConditionSave.frx":035C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optNew 
         Height          =   255
         Index           =   0
         Left            =   548
         TabIndex        =   5
         Top             =   840
         Width           =   1575
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
         Caption         =   "frmConditionSave.frx":0378
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmConditionSave.frx":03B0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmConditionSave.frx":03D0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optNew 
         Height          =   255
         Index           =   1
         Left            =   548
         TabIndex        =   4
         Top             =   1210
         Width           =   2175
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
         Caption         =   "frmConditionSave.frx":03EC
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmConditionSave.frx":0434
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmConditionSave.frx":0454
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optNew 
         Height          =   255
         Index           =   2
         Left            =   548
         TabIndex        =   3
         Top             =   1580
         Width           =   1395
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
         Caption         =   "frmConditionSave.frx":0470
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmConditionSave.frx":04A8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmConditionSave.frx":04C8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optNew 
         Height          =   255
         Index           =   3
         Left            =   548
         TabIndex        =   2
         Top             =   3780
         Width           =   2595
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
         Caption         =   "frmConditionSave.frx":04E4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmConditionSave.frx":0526
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmConditionSave.frx":0546
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboSystem 
         Height          =   315
         Left            =   788
         TabIndex        =   1
         Top             =   3180
         Width           =   4455
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
         Tip             =   "frmConditionSave.frx":0562
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmConditionSave.frx":0582
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   735
         Left            =   308
         Top             =   120
         Width           =   4455
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
         Caption         =   "frmConditionSave.frx":059E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmConditionSave.frx":06D4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmConditionSave.frx":06F4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmConditionSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmConditionSave.frm
'' Description: Allow the user to save a Trade Sense expression in various ways
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kFormHeight = 5250

Private Enum eGDSaveOption
    eGDSaveOption_Function = 0
    eGDSaveOption_Scoring
    eGDSaveOption_Criteria
    eGDSaveOption_Clipboard
    eGDSaveOption_Rule
    eGDSaveOption_HighlightBar
    eGDSaveOption_Alert
End Enum

Private Type mPrivate
    eSaveType As eExprType              ' Expression type
    lComboIndex As Long                 ' Index for the combo box
    strAlert As String                  ' Alert string
    strSymbol As String                 ' Symbol
End Type
Private m As mPrivate

Private Property Get SaveOption(ByVal nSaveOption As eGDSaveOption) As OptionButton
    Set SaveOption = optNew(nSaveOption)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Expression Type, System List, Selected System, Alert, Symbol,
''              Expression has Assignment?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(eSaveAsType As eExprType, tblSystemList As cGdTable, lSelectedSystem As Long, strAlert As String, ByVal strSymbol As String, Optional ByVal bHasAssignment As Boolean = False)
On Error GoTo ErrSection:

    cmdOK.Enabled = False
    m.strSymbol = strSymbol
    
    SaveOption(eGDSaveOption_Scoring).Enabled = Not bHasAssignment
    SaveOption(eGDSaveOption_Criteria).Enabled = Not bHasAssignment
    SaveOption(eGDSaveOption_Alert).Enabled = Not bHasAssignment
    
    If ExtremeCharts = 1 Then
        txtAlert.Visible = False
        cboSystem.Visible = False
        SaveOption(eGDSaveOption_Rule).Visible = False
        SaveOption(eGDSaveOption_Alert).Visible = False
        SaveOption(eGDSaveOption_Clipboard).Top = SaveOption(eGDSaveOption_Alert).Top
        cmdOK.Top = SaveOption(eGDSaveOption_Rule).Top + 100
        cmdCancel.Top = cmdOK.Top
        
        Me.Height = kFormHeight - (SaveOption(eGDSaveOption_Rule).Height + cboSystem.Height) * 2
    Else
        Me.Height = kFormHeight
        LoadSystemCombo tblSystemList
    End If
    
    CenterTheForm Me
    ShowForm Me, eForm_Modal, frmMain
    
    eSaveAsType = m.eSaveType
    lSelectedSystem = m.lComboIndex
    strAlert = m.strAlert
    
ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmConditionSave.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboSystem_Click
'' Description: Handle the user changing the system combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboSystem_Click()
On Error GoTo ErrSection:

    m.lComboIndex = cboSystem.ListIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionSave.cboSystem_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Hide the form and allow the ShowMe to unload it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.eSaveType = eType_Undefined
    m.lComboIndex = -1
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionSave.cmdCancel_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Hide the form and allow the ShowMe to unload it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.strAlert = txtAlert.Text
    Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionSave.cmdOK_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Show help if appropriate
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Nothing
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionSave.Form_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form when loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    g.Styler.StyleForm Me
    
    Me.Icon = Picture16(ToolbarIcon("ID_ConditionBuilder"), , True)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionSave.Form_Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Determine whether to continue with the unloading
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.eSaveType = eType_Undefined
        m.lComboIndex = -1
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionSave.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optNew_Click
'' Description: Handle the user clicking on one of the option buttons
'' Inputs:      Index for the option button
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optNew_Click(Index As Integer)
On Error GoTo ErrSection:

    Select Case Index
        Case eGDSaveOption_Function:
            m.eSaveType = eType_Function
        
        Case eGDSaveOption_Scoring:
            m.eSaveType = eType_Scoring
        
        Case eGDSaveOption_Criteria:
            m.eSaveType = eType_Criteria
        
        Case eGDSaveOption_Clipboard:
            m.eSaveType = eType_Clipboard
            
        Case eGDSaveOption_Rule:
            m.eSaveType = eType_Rule
            
        Case eGDSaveOption_HighlightBar:
            m.eSaveType = eType_HighlightBars
            
        Case eGDSaveOption_Alert:
            If Not HasGold(True, "Adding a chart alert") Then
                SaveOption(eGDSaveOption_Alert).Value = False
                SaveOption(eGDSaveOption_Alert).Enabled = False
                txtAlert.Enabled = False
                Exit Sub
            End If
            m.eSaveType = eType_Alert
            
        Case Else:
            m.eSaveType = eType_Undefined
    
    End Select
    
    If Index = eGDSaveOption_Alert Then
        txtAlert.Enabled = True
        If Len(m.strSymbol) > 0 Then
            txtAlert.Text = m.strSymbol & ": Chart Alert"
        Else
            txtAlert.Text = "Enter alert name here"
        End If
    Else
        txtAlert.Enabled = False
        txtAlert.Text = ""
    End If
    
    If (Index >= eGDSaveOption_Function) And (Index <= eGDSaveOption_Alert) Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionSave.optNew_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSystemCombo
'' Description: Load up the system combo box
'' Inputs:      System List
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadSystemCombo(tblSystemList As cGdTable)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    cboSystem.Clear
    
    For lIndex = 0 To tblSystemList.NumRecords - 1
        cboSystem.AddItem tblSystemList(0, lIndex)
    Next
    
    cboSystem.ListIndex = 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionSave.LoadSystemCombo"
    
End Sub

