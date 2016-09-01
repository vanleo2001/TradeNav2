VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTradeRule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   900
      TabIndex        =   0
      Top             =   1440
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
      Caption         =   "frmTradeRule.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeRule.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeRule.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   1440
         TabIndex        =   2
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
         Caption         =   "frmTradeRule.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeRule.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeRule.frx":00B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   4
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
         Caption         =   "frmTradeRule.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeRule.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeRule.frx":0118
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txtDescription 
      Height          =   315
      Left            =   1260
      TabIndex        =   5
      Top             =   840
      Width           =   3135
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmTradeRule.frx":0134
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
      Tip             =   "frmTradeRule.frx":0154
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeRule.frx":0174
   End
   Begin HexUniControls.ctlUniTextBoxXP txtAbbreviation 
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   480
      Width           =   795
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmTradeRule.frx":0190
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
      Tip             =   "frmTradeRule.frx":01B0
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeRule.frx":01D0
   End
   Begin HexUniControls.ctlUniTextBoxXP txtName 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   120
      Width           =   2355
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmTradeRule.frx":01EC
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
      Tip             =   "frmTradeRule.frx":020C
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeRule.frx":022C
   End
   Begin HexUniControls.ctlUniLabelXP lblDescription 
      Height          =   255
      Left            =   180
      Top             =   900
      Width           =   1035
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
      Caption         =   "frmTradeRule.frx":0248
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTradeRule.frx":0282
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeRule.frx":02A2
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAbbreviation 
      Height          =   255
      Left            =   180
      Top             =   540
      Width           =   1035
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
      Caption         =   "frmTradeRule.frx":02BE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTradeRule.frx":02FA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeRule.frx":031A
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblName 
      Height          =   255
      Left            =   180
      Top             =   180
      Width           =   1035
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
      Caption         =   "frmTradeRule.frx":0336
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTradeRule.frx":0362
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeRule.frx":0382
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmTradeRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTradeRule.frm
'' Description: Allow the user to edit or create a new trade rule
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on the OK button?
    TradeRule As cTradeRule             ' Trade rule object
    astrRules As cGdArray               ' Array of trade rules for checking uniqueness
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Trade Rule
'' Returns:     True if OK clicked, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(TradeRule As cTradeRule) As Boolean
On Error GoTo ErrSection:

    Set m.TradeRule = TradeRule
    txtName.Text = TradeRule.Name
    txtAbbreviation.Text = TradeRule.Abbreviation
    txtDescription.Text = TradeRule.Description
    
    Set m.astrRules = New cGdArray
    If TradeRule.RuleType = eGDTradeRuleType_Entry Then
        m.astrRules.FromFile AddSlash(App.Path) & "Provided\ErFilter.TXT"
        m.astrRules.FromFile AddSlash(App.Path) & "Custom\ErFilter.TXT", True
        SetEditorCaption Me, "Entry Trade Rule", TradeRule.Name
    Else
        m.astrRules.FromFile AddSlash(App.Path) & "Provided\XrFilter.TXT"
        m.astrRules.FromFile AddSlash(App.Path) & "Custom\XrFilter.TXT", True
        SetEditorCaption Me, "Exit Trade Rule", TradeRule.Name
    End If
    
    MoveFocus txtName
    ShowForm Me, eForm_ActModal, frmMain
    
    If m.bOK Then
        If TradeRule.ID = 0 Then
            TradeRule.ID = NextCustomTradeRuleID(TradeRule.RuleType)
        End If
        TradeRule.Name = txtName.Text
        TradeRule.Abbreviation = txtAbbreviation.Text
        TradeRule.Description = txtDescription.Text
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTradeRule.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Close the form without saving information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRule.cmdCancel_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Close the form and save the information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    MoveFocus cmdOK
    
    If Len(Trim(txtName.Text)) = 0 Then
        MoveFocus txtName
        InfBox "You must supply a Trade Rule name", "e", , "Trade Rule Error"
    ElseIf Len(Trim(txtName.Text)) > 30 Then
        MoveFocus txtName
        InfBox "Trade Rule name must be 30 characters in length or less", "e", , "Trade Rule Error"
    ElseIf InStr(Trim(txtName.Text), "-") <> 0 Then
        MoveFocus txtName
        InfBox "Trade Rule name cannot contain a dash ('-')", "e", , "Trade Rule Error"
    ElseIf CheckName = False Then
        MoveFocus txtName
        InfBox "Trade Rule name must be unique", "e", , "Trade Rule Error"
    ElseIf Len(Trim(txtAbbreviation.Text)) = 0 Then
        MoveFocus txtAbbreviation
        InfBox "You must supply a Trade Rule abbreviation", "e", , "Trade Rule Error"
    ElseIf Len(Trim(txtAbbreviation.Text)) > 5 Then
        MoveFocus txtAbbreviation
        InfBox "Trade Rule abbreviation must be 5 characters in length or less", "e", , "Trade Rule Error"
    ElseIf InStr(Trim(txtAbbreviation.Text), "-") <> 0 Then
        MoveFocus txtAbbreviation
        InfBox "Trade Rule abbreviation cannot contain a dash ('-')", "e", , "Trade Rule Error"
    ElseIf CheckAbbreviation = False Then
        MoveFocus txtAbbreviation
        InfBox "Trade Rule abbreviation must be unique", "e", , "Trade Rule Error"
    Else
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRule.cmdOK_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement of the form
    
    strPlacement = GetIniFileProperty("frmTradeRule", "", "Placement", g.strIniFile)
    If Len(strPlacement) > 0 Then
        SetFormPlacement Me, strPlacement, "LTHW"
    Else
        CenterTheForm Me
    End If
    
    g.Styler.StyleForm Me
    
    Icon = Picture16("kBlank")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRule.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the 'X', unload form without saving
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = vbFormCode Then
        m.bOK = False
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRule.Form_QueryUnload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    fraButtons.Move (ScaleWidth - fraButtons.Width) / 2

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save information and clean up when form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "frmTradeRule", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRule.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAbbreviation_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAbbreviation_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtAbbreviation

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRule.txtAbbreviation_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAbbreviation_LostFocus
'' Description: When the control loses the focus, change it to upper case
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAbbreviation_LostFocus()
On Error GoTo ErrSection:

    txtAbbreviation.Text = Trim(UCase(txtAbbreviation.Text))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRule.txtAbbreviation_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDescription_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDescription_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtDescription

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRule.txtDescription_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtName_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtName_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRule.txtName_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckName
'' Description: Check to see if name is unique
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckName() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    For lIndex = 0 To m.astrRules.Size - 1
        If UCase(Trim(txtName.Text)) = UCase(Parse(Parse(m.astrRules(lIndex), vbTab, 2), "-", 2)) Then
            If m.TradeRule.ID <> CLng(Val(Parse(m.astrRules(lIndex), vbTab, 1))) Then
                bReturn = False
                Exit For
            End If
        End If
    Next lIndex
    
    CheckName = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeRule.CheckName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckAbbreviation
'' Description: Check to see if abbreviation is unique
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckAbbreviation() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    For lIndex = 0 To m.astrRules.Size - 1
        If UCase(Trim(txtAbbreviation.Text)) = UCase(Parse(Parse(m.astrRules(lIndex), vbTab, 2), "-", 1)) Then
            If m.TradeRule.ID <> CLng(Val(Parse(m.astrRules(lIndex), vbTab, 1))) Then
                bReturn = False
                Exit For
            End If
        End If
    Next lIndex
    
    CheckAbbreviation = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeRule.CheckAbbreviation"
    
End Function

