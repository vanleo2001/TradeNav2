VERSION 5.00
Object = "{C0F09D2D-1125-11D7-8DA9-0004757A4B66}#1.9#0"; "NavTradeSenseOCXV3.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmCustomCondition 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraNumDays 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5235
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
      Caption         =   "frmCustomCondition.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCustomCondition.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCustomCondition.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtNumBars 
         Height          =   315
         Left            =   4080
         TabIndex        =   7
         Top             =   60
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmCustomCondition.frx":005C
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
         Tip             =   "frmCustomCondition.frx":007E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCustomCondition.frx":009E
      End
      Begin HexUniControls.ctlUniTextBoxXP txtOverride 
         Height          =   315
         Left            =   4140
         TabIndex        =   8
         Top             =   180
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmCustomCondition.frx":00BA
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
         Tip             =   "frmCustomCondition.frx":00DC
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCustomCondition.frx":00FC
      End
      Begin HexUniControls.ctlUniRadioXP optOverride 
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   300
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
         Caption         =   "frmCustomCondition.frx":0118
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmCustomCondition.frx":0156
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmCustomCondition.frx":0176
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAutoDetect 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   300
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
         Caption         =   "frmCustomCondition.frx":0192
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmCustomCondition.frx":01E4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmCustomCondition.frx":0204
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblNumBars1 
         Height          =   195
         Left            =   0
         Top             =   60
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
         Caption         =   "frmCustomCondition.frx":0220
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmCustomCondition.frx":02B4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCustomCondition.frx":02D4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraTop 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
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
      Caption         =   "frmCustomCondition.frx":02F0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCustomCondition.frx":031C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCustomCondition.frx":033C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboDefaultPeriod 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   0
         Width           =   1575
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   0
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
         Tip             =   "frmCustomCondition.frx":0358
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmCustomCondition.frx":0378
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblDefaultPeriod 
         Height          =   255
         Left            =   0
         Top             =   30
         Width           =   1155
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
         Caption         =   "frmCustomCondition.frx":0394
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmCustomCondition.frx":03D4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCustomCondition.frx":03F4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin NavTradeSenseV3.Editor tsCondition 
      Height          =   2655
      Left            =   180
      TabIndex        =   9
      Top             =   2100
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   4683
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1695
      Left            =   5580
      TabIndex        =   10
      Top             =   120
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
      Caption         =   "frmCustomCondition.frx":0410
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCustomCondition.frx":043C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCustomCondition.frx":045C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdVerify 
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   1200
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
         Caption         =   "frmCustomCondition.frx":0478
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCustomCondition.frx":04A6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCustomCondition.frx":04C6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   540
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
         Caption         =   "frmCustomCondition.frx":04E2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCustomCondition.frx":0510
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCustomCondition.frx":0530
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   11
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
         Caption         =   "frmCustomCondition.frx":054C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCustomCondition.frx":0572
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCustomCondition.frx":0592
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmCustomCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmCustomCondition.frm
'' Description: Allow the user to create a custom condition for an order
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/19/2010   DAJ         Fixed RunExpressions when secondary markets (#5437)
'' 03/11/2010   DAJ         Added numbars required stuff for conditional orders (#5580)
'' 07/15/2013   DAJ         Allow 'Of Monthly' in conditional order expression
'' 11/04/2013   DAJ         Allow system and criteria functions for expression
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user OK the dialog
    lSymbolID As Long                   ' Symbol ID
    strSymbol As String                 ' Symbol
    
    ListLoading As cListLoading         ' Lists of stuff for TradeSense
End Type
Private m As mPrivate

Private Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID = 0& Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If
End Property
Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Current Expression, Symbol, Default Period, Override,
''              Num Bars Calc, Num Bars Override
'' Returns:     New Expression if OK, Blank string otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal strExpression As String, ByVal vSymbolOrSymbolID As Variant, strDefaultPeriod As String, Optional bOverride As Boolean = False, Optional lNumBarsCalc As Long = 0&, Optional lNumBarsOverride As Long = 0&) As String
On Error GoTo ErrSection:

    optOverride.Value = bOverride
    txtNumBars.Text = Str(lNumBarsCalc)
    txtOverride.Text = Str(lNumBarsOverride)
    
    SymbolOrSymbolID = vSymbolOrSymbolID

    tsCondition.Text = strExpression
    cboDefaultPeriod.Text = strDefaultPeriod
    Verify False
 
    EnableControls
    MoveFocus tsCondition
    
    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK = True Then
        strDefaultPeriod = cboDefaultPeriod.Text
        bOverride = optOverride.Value
        lNumBarsCalc = CLng(ValOfText(txtNumBars.Text))
        lNumBarsOverride = CLng(ValOfText(txtOverride.Text))
        
        ShowMe = tsCondition.Text
    Else
        strDefaultPeriod = strDefaultPeriod
        ShowMe = ""
    End If

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmCustomCondition.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboDefaultPeriod_KeyPress
'' Description: Fix the period that the user typed in to the correct one
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboDefaultPeriod_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        MoveFocus tsCondition
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.cboDefaultPeriod_KeyPress"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboDefaultPeriod_LostFocus
'' Description: Fix the period that the user typed in to the correct one
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboDefaultPeriod_LostFocus()
On Error GoTo ErrSection:

    Dim strPeriod As String             ' Adjusted version of period typed in
    
    strPeriod = GetPeriodStr(cboDefaultPeriod.Text)
    If strPeriod <> cboDefaultPeriod.Text Then
        cboDefaultPeriod.Text = strPeriod
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.cboDefaultPeriod_LostFocus"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Close the form without saving the changes
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
    RaiseError "frmCustomCondition.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Close the form and save the changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If Verify Then
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdVerify_Click
'' Description: Verify the condition the user has typed in
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdVerify_Click()
On Error GoTo ErrSection:

    Verify

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.cmdVerify_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated, give the editor the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    MoveFocus tsCondition

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement for the form from ini file
    
    g.Styler.StyleForm Me

    Icon = Picture16("kBlank")
    Caption = "Custom Condition for Order"

    strPlacement = GetIniFileProperty("frmCustomCondition", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If

    'Load internally generated TradeSense lists (Symbols, etc.)
    ' (when activate, in case list has changed)
    Set m.ListLoading = New cListLoading
    m.ListLoading.Load
    
    InitializeEditor
    
    With cboDefaultPeriod
        .AddItem "Daily"
        .AddItem "60 Minute"
        .AddItem "30 Minute"
        .AddItem "10 Minute"
        .AddItem "5 Minute"
    End With

    txtNumBars.Locked = True
    txtNumBars.Enabled = False
    txtOverride.Move txtNumBars.Left, txtNumBars.Top

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If code did not unload the form, allow ShowMe to finish
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.Form_QueryUnload"
    
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

    Dim lMinWidth As Long               ' Minimum scale width allowed
    Dim lMinHeight As Long              ' Minimum scale height allowed
    Dim lTop As Long                    ' Top of the condition editor
    
    'lMinWidth = (fraButtons.Width * 5) + 120
    lMinWidth = fraNumDays.Width + fraButtons.Width + 180
    lMinHeight = fraButtons.Height + 120
    
    If LimitFormSize(Me, lMinWidth, lMinHeight) Then Exit Sub
    
    With fraButtons
        .Move ScaleWidth - .Width - 60, 60
    End With
    
    With tsCondition
        'lTop = fraTop.Top + fraTop.Height + 60
        lTop = fraNumDays.Top + fraNumDays.Height + 60
        .Move 60, lTop, ScaleWidth - fraButtons.Width - 180, ScaleHeight - lTop - 60
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save things and clean up when the form is unloaded
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "frmCustomCondition", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAutoDetect_Click
'' Description: Enable/Disable the controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAutoDetect_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.optAutoDetect_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optOverride_Click
'' Description: Enable/Disable the controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optOverride_Click()
On Error GoTo ErrSection:

    If Visible Then
        If (Len(Trim(txtOverride.Text)) = 0) And (Len(Trim(txtNumBars.Text)) > 0) Then
            txtOverride.Text = Trim(txtNumBars.Text)
        End If
        
        EnableControls
        MoveFocus txtOverride
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.optOverride_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsCondition_Change
'' Description: As the condition changes, make sure controls are enabled
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsCondition_Change()
On Error GoTo ErrSection:

    EnableControls

    ' Don't allow the user to use an assignment in the expression for now...
    If InStr(tsCondition.Text, ":=") <> 0 Then
        InfBox "You cannot have an assignment operator in this expression.", "!", , "Expression Error"
        tsCondition.Text = Replace(tsCondition.Text, ":=", "")
        If Len(tsCondition.Text) > 0 Then
            tsCondition.SelStart = Len(tsCondition.Text)
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.tsCondition_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsCondition_GotFocus
'' Description: Reinitialize the control when it gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsCondition_GotFocus()
On Error GoTo ErrSection:

    Set g.ActiveEditor = tsCondition
    InitializeEditor
    
    If Len(Trim(tsCondition.Text)) = 0 Then
        tsCondition.Text = ""
        SendKeys " "
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.tsCondition_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsCondition_LostFocus
'' Description: Clean up after the control loses the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsCondition_LostFocus()
On Error GoTo ErrSection:

    Set g.ActiveEditor = Nothing
    tsCondition.RemoveTradeSense

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.tsCondition_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Verify
'' Description: Verify the condition
'' Inputs:      None
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Verify(Optional ByVal bShowMsg As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim Expr As New cExpression         ' Expression to verify condition
    Dim Func As New cFunction           ' Temporary function object
    Dim Inputs As New cInputs           ' Collection of inputs for the expression
    Dim bExtraInputs As Boolean         ' Does the expression have extra inputs?
    Dim strNotKnown As String           ' Inputs that are not known
    Dim lIndex As Long                  ' Index into a for loop
    Dim strParmName As String           ' Parameter Name
    Dim lNumBars As Long                ' Number of bars necessary
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Len(Trim(tsCondition.Text)) = 0 Then
        If bShowMsg Then InfBox "Must specify an expression", "!", , "Custom Condition Error"
    Else
        With Expr
            .PortfolioNavigator = False
            .Functions = g.Functions
            .ValidateFunctionRule tsCondition.Text
            
            If m.ListLoading Is Nothing Then
                Set m.ListLoading = New cListLoading
                m.ListLoading.Load
            End If
            
            tsCondition.TurnOffEditing
            tsCondition.TextRTF = Func.GetRTF(.EditText)
            tsCondition.ExprIsFormatted = True
            
            bExtraInputs = False
            strNotKnown = ""
            If Not Expr.Inputs Is Nothing Then
                Set Inputs = Expr.Inputs
                For lIndex = 1 To Inputs.Count
                    strParmName = Inputs.Item(lIndex).ParmName
                    If Inputs.Item(lIndex).ParmTypeID <> 5 Then
                        strNotKnown = strNotKnown & "|" & strParmName
                        bExtraInputs = True
                    ElseIf UCase(Left(strParmName, 7)) <> "MARKET1" Then
                        If (UCase(strParmName) <> "DAILY") And (UCase(strParmName) <> "WEEKLY") And (UCase(strParmName) <> "MONTHLY") And _
                                Left(strParmName, 1) <> Chr(34) And Right(strParmName, 1) <> Chr(34) Then
                            strNotKnown = strNotKnown & "|" & strParmName
                            bExtraInputs = True
                        End If
                    End If
                Next lIndex
            End If
            
            If bExtraInputs Then
                MoveFocus tsCondition
                If bShowMsg Then InfBox "There are unrecognized inputs in your expression:|" & strNotKnown & "|", "!", , "Custom Condition Error"
            ElseIf .FunctionReturnType <> 3 And .FunctionReturnType <> 6 Then
                MoveFocus tsCondition
                If bShowMsg Then InfBox "Expression must be a Boolean Expression", "!", , "Custom Condition Error"
            ElseIf EngineVerify(.CodedText) Then
                lNumBars = AutoDetect(.CodedText)
                If lNumBars > 0 Then
                    bReturn = True
                ElseIf (optOverride.Value = True) And (ValOfText(txtOverride.Text) > 0#) Then
                    bReturn = True
                End If
            End If
        End With
    End If
    
    Verify = bReturn

ErrExit:
    Set Expr = Nothing
    Set Func = Nothing
    Exit Function
    
ErrSection:
    Set Expr = Nothing
    Set Func = Nothing
    RaiseError "frmCustomCondition.Verify"
    
End Function

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

    cmdOK.Enabled = (Len(Trim(tsCondition.Text)) > 0)
    cmdVerify.Enabled = (Len(Trim(tsCondition.Text)) > 0)

    If optOverride.Value = True Then
        txtNumBars.Visible = False
        txtOverride.Visible = True
    Else
        txtNumBars.Visible = True
        txtOverride.Visible = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoDetect
'' Description: Auto Detect how many bars the criteria needs
'' Inputs:      Expression
'' Returns:     Number of Bars needed, -1 if not detected
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AutoDetect(ByVal strExpression As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim AD As New cAutoDetect           ' Auto detection object
    
    lReturn = AD.AutoDetect(strExpression, SymbolOrSymbolID, cboDefaultPeriod.Text)
    txtNumBars.Text = Str(lReturn)
    
    If (ValOfText(txtOverride.Text) < lReturn) And (optOverride.Value = True) Then
        InfBox "Trade Navigator has determined that your custom condition needs at least " & _
            Trim(CStr(lReturn)) & " bars to run properly.  " & _
            "The value has been set accordingly.", _
            "i", , "Custom Condition"
        optAutoDetect = True
        txtOverride.Text = Str(lReturn)
    End If
    
    If (lReturn = -1&) And ((optAutoDetect.Value = True) Or (ValOfText(txtOverride.Text) <= 0)) Then
        InfBox "Trade Navigator could not automatically determine how many bars are needed to calculate " & _
                " the custom condition.  Please specify an override for the number of necessary bars.", _
                "!", , "Custom Condition Error"
        optOverride = True
        MoveFocus txtOverride
    End If
    
    AutoDetect = lReturn

#If 0 Then
    Dim lAutoDetect As Long             ' Auto detected number of bars
    
    '11936: IBM
    '41180: SP-067
    '50: $DJIA
        
    ' 1 year of IBM
    lAutoDetect = RunExpression(strExpression, 11936)
    If lAutoDetect > 0 Then
        If RunExpression(strExpression, 11936, 365, True) <> lAutoDetect Then
            lAutoDetect = -1
        End If
    ElseIf lAutoDetect = -99 Then
        lAutoDetect = -99
        GoTo ErrExit
    End If
    
    ' 10 years of IBM
    If lAutoDetect = 0& Then
        lAutoDetect = RunExpression(strExpression, 11936, 3650)
        If lAutoDetect > 0 Then
            If RunExpression(strExpression, 11936, 3650, True) <> lAutoDetect Then
                lAutoDetect = -1
            End If
        ElseIf lAutoDetect = -99 Then
            lAutoDetect = -99
            GoTo ErrExit
        End If
    End If
    
    ' Full History of IBM
    If lAutoDetect = 0& Then
        lAutoDetect = RunExpression(strExpression, 11936)
        If lAutoDetect > 0 Then
            If RunExpression(strExpression, 11936, -1, True) <> lAutoDetect Then
                lAutoDetect = -1
            End If
        ElseIf lAutoDetect = -99 Then
            lAutoDetect = -99
            GoTo ErrExit
        End If
    End If
    
    ' 1 year of SP-067
    If lAutoDetect = 0& Then
        lAutoDetect = RunExpression(strExpression, 41180, 365)
        If lAutoDetect > 0 Then
            If RunExpression(strExpression, 41180, 365, True) <> lAutoDetect Then
                lAutoDetect = -1
            End If
        ElseIf lAutoDetect = -99 Then
            lAutoDetect = -99
            GoTo ErrExit
        End If
    End If
    
    ' 10 years of SP-067
    If lAutoDetect = 0& Then
        lAutoDetect = RunExpression(strExpression, 41180, 3650)
        If lAutoDetect > 0 Then
            If RunExpression(strExpression, 41180, 3650, True) <> lAutoDetect Then
                lAutoDetect = -1
            End If
        ElseIf lAutoDetect = -99 Then
            lAutoDetect = -99
            GoTo ErrExit
        End If
    End If
    
    ' Full History of SP-067
    If lAutoDetect = 0& Then
        lAutoDetect = RunExpression(strExpression, 41180)
        If lAutoDetect > 0 Then
            If RunExpression(strExpression, 41180, -1, True) <> lAutoDetect Then
                lAutoDetect = -1
            End If
        ElseIf lAutoDetect = -99 Then
            lAutoDetect = -99
            GoTo ErrExit
        End If
    End If
    
    ' 1 year of $DJIA
    If lAutoDetect = 0& Then
        lAutoDetect = RunExpression(strExpression, 50, 365)
        If lAutoDetect > 0 Then
            If RunExpression(strExpression, 50, 365, True) <> lAutoDetect Then
                lAutoDetect = -1
            End If
        ElseIf lAutoDetect = -99 Then
            lAutoDetect = -99
            GoTo ErrExit
        End If
    End If
    
    ' 10 years of $DJIA
    If lAutoDetect = 0& Then
        lAutoDetect = RunExpression(strExpression, 50, 3650)
        If lAutoDetect > 0 Then
            If RunExpression(strExpression, 50, 3650, True) <> lAutoDetect Then
                lAutoDetect = -1
            End If
        ElseIf lAutoDetect = -99 Then
            lAutoDetect = -99
            GoTo ErrExit
        End If
    End If
    
    ' Full History of $DJIA
    If lAutoDetect = 0& Then
        lAutoDetect = RunExpression(strExpression, 50)
        If lAutoDetect > 0 Then
            If RunExpression(strExpression, 50, -1, True) <> lAutoDetect Then
                lAutoDetect = -1
            End If
        ElseIf lAutoDetect = -99 Then
            lAutoDetect = -99
            GoTo ErrExit
        End If
    End If
    
    AutoDetect = lAutoDetect
#End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCustomCondition.AutoDetect"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunExpression
'' Description: Run the expression to determine how many bars are needed
'' Inputs:      Expression, Symbol ID, Num Bars to Load, Delay One Bar?
'' Returns:     Number of Bars Necessary
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RunExpression(ByVal strExpression As String, ByVal lSymbolID As Long, _
        Optional ByVal lNumBarsToLoad As Long = -1&, Optional ByVal bDelayOneBar As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim i&, ii&, rc&, d#, hArray&, nAutoDetect&
    Dim nRecord&, nCount&, nStartDate&
    Dim strCodedText$
    Dim Bars As New cGdBars
    Dim Daily As New cGdBars
    Dim Weekly As New cGdBars
    Dim Monthly As New cGdBars
    Dim GC As New cGdBars
    
    Dim astrParms As New cGdArray, astrBarNames As New cGdArray
    Dim aScanExpr As New cGdArray, aArrayOfResults As New cGdArray
    Dim aArrayOfBars As New cGdArray
    Dim aScanArrays As New cGdArray
    Dim aMinBarsReq As New cGdArray
    
    Dim iDayOfWeek As Integer
    
    Dim SecondaryMarkets As New cGdTree ' Bars collection of secondary markets
    Dim lBars As Long                   ' Index into a for loop

    ' Get coded text and handle of values array from each Criteria
    aScanExpr.Create eGDARRAY_Strings
    aScanArrays.Create eGDARRAY_Longs
    aArrayOfResults.Create eGDARRAY_Longs
    aMinBarsReq.Create eGDARRAY_Longs
    strCodedText = Trim(strExpression)

    nRecord = g.SymbolPool.PoolRecForSymbolID(lSymbolID)
    If Len(strCodedText) > 0 And nRecord >= 0 Then
        If lNumBarsToLoad = -1& Then
            nStartDate = 0
        Else
            nStartDate = LastDailyDownload - lNumBarsToLoad
            iDayOfWeek = Weekday(nStartDate)
            nStartDate = nStartDate - (iDayOfWeek - vbMonday)
        End If
        
        'nSymbolID = g.SymbolPool.SymbolID(nRecord)

        aArrayOfBars.Create eGDARRAY_Longs
        Bars.Size = 0
            
        If lSymbolID <> 0 Then
            ' load a year's worth of data
            If Not DM_GetBars(Bars, lSymbolID, , nStartDate, 0, , , , False) Then
                Bars.Size = 0
            ElseIf bDelayOneBar Then
                ' need to start with next bar
                nStartDate = Int(Bars(eBARS_DateTime, 0)) + 1
                'If optWeekly.Value = True Then
                '    nStartDate = nStartDate + 6
                'End If
                If Not DM_GetBars(Bars, lSymbolID, , nStartDate, 0, , , , False) Then
                    Bars.Size = 0
                End If
            End If
            nStartDate = Bars(eBARS_DateTime, 0)
        End If
            
        If Bars.Size > 0 Then
            aScanExpr.Add strCodedText
            
            Daily.BuildBars "Daily", Bars.BarsHandle
            Weekly.BuildBars "Weekly", Bars.BarsHandle
            Monthly.BuildBars "Monthly", Bars.BarsHandle
            SecondaryMarkets.Add Bars
            SecondaryMarkets.Add Daily
            SecondaryMarkets.Add Weekly
            SecondaryMarkets.Add Monthly
            
            astrBarNames.Add "Market1"
            astrBarNames.Add "Daily"
            astrBarNames.Add "Weekly"
            astrBarNames.Add "Monthly"
            
            MarketsInExpressions aScanExpr, nStartDate, False, astrBarNames, SecondaryMarkets, "Daily"
            
            ' create a temporary result array to be used
            ' by the expression evaluator
            hArray = gdCreateArray(eGDARRAY_Doubles, Bars.Size)
            aArrayOfResults.Add hArray
            
            ' Init the expression evaluator with list of scan expressions
            astrParms(0) = "CriteriaRunExp"
            If Not SetupExpressions(astrParms, astrBarNames, aScanExpr) Then
                InfBox "i=[] ; h=Criteria ; An error exists in a Criteria expression."
                RunExpression = -99
                Exit Function
            End If
    
            ' run engine to evaluate expressions for this symbol
            aArrayOfBars.Num(0) = Bars.BarsHandle '(in case changed)
            aArrayOfBars.Num(1) = Daily.BarsHandle
            aArrayOfBars.Num(2) = Weekly.BarsHandle
            aArrayOfBars.Num(3) = Monthly.BarsHandle
            
            'aArrayOfBars.Num(2) = GC.BarsHandle
            For lBars = 4 To astrBarNames.Size - 1
                aArrayOfBars.Num(lBars) = SecondaryMarkets(lBars + 1).BarsHandle
            Next lBars
            astrParms.Size = 1
            rc = RunExpressions(astrParms.ArrayHandle, _
                astrBarNames.ArrayHandle, aArrayOfBars.ArrayHandle, _
                aArrayOfResults.ArrayHandle, aMinBarsReq.ArrayHandle, ByVal 0&)
            If rc = 0 Then
                ' see if last value is not null
                If aMinBarsReq.Size > 0 Then
                    ' new method (engine calculates the number)
                    If aMinBarsReq(0) < Bars.Size Then
                        nAutoDetect = aMinBarsReq(0) + 1
                        'If optWeekly = False And InStr(UCase(strExpression), "~07006WEEKLY") <> 0 Then
                        '    If nAutoDetect = 0 Then
                        '        nAutoDetect = 5
                        '    Else
                                ' figure number of daily bars for full weeks
                                d = Bars(eBARS_DateTime, nAutoDetect - 1) - Bars(eBARS_DateTime, 0)
                                nAutoDetect = Int((d + 6) / 7) * 5
                                If nAutoDetect < 5 Then
                                    nAutoDetect = 5
                                End If
                        '    End If
                        'End If
                    End If
                Else
                    hArray = aArrayOfResults.Num(0)
                    d = gdGetNum(hArray, gdGetSize(hArray) - 1)
                    If d <> gdNullValue(hArray) Then
                        ' if so, find first non-null item
                        For i = 0 To gdGetSize(hArray) - 1
                            d = gdGetNum(hArray, i)
                            If d <> gdNullValue(hArray) Then
                                nAutoDetect = i + 1
                                If InStr(UCase(strExpression), "~07006WEEKLY") <> 0 Then
                                    gdCopy Bars.ArrayHandle(eBARS_Close), hArray
                                    Bars.BuildBars "Weekly"
                                    For ii = 0 To Bars.Size - 1
                                        If Bars(eBARS_Close, ii) <> gdNullValue(Bars.ArrayHandle(eBARS_Close)) Then
                                            nAutoDetect = Int((ii + 4) / 5)
                                            Exit For
                                        End If
                                    Next ii
                                End If
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
            ' clear the expression evaluator when done with it
            SetupExpressions astrParms '(clear expressions)
        End If
            
    End If
    
    ' destroy all the temporary result arrays
    For i = 0 To aArrayOfResults.Size - 1
        gdDestroyArray aArrayOfResults(i)
    Next
    aArrayOfResults.Size = 0
    
    RunExpression = nAutoDetect

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCustomCondition.RunExpression"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EngineVerify
'' Description: Verify the expression with the engine
'' Inputs:      None
'' Returns:     True if verifies through the engine, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EngineVerify(ByVal strCodedText As String) As Boolean
On Error GoTo ErrSection:

    Dim astrParms As New cGdArray       ' Parameters to pass to the engine
    Dim astrBarNames As New cGdArray    ' List of Bar Names to pass to the engine
    Dim aScanExpr As New cGdArray       ' List of expressions to pass to the engine
    Dim strError As String              ' Error message back from the engine
    Dim bInvalidSecondaryPeriod As Boolean ' Invalid secondary period?
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Len(strCodedText) > 0 Then
        ' Init the expression evaluator with list of scan expressions
        aScanExpr.Add strCodedText
        
        MarketsInExpressions aScanExpr, 0#, False, astrBarNames, Nothing, cboDefaultPeriod.Text, m.strSymbol, bInvalidSecondaryPeriod
        
        If bInvalidSecondaryPeriod Then
            InfBox "You cannot have a secondary market have intraday data when the default period is not intraday", "!", , "Custom Condition Error"
        Else
            astrParms(0) = "OrderConditionVerify"
            If SetupExpressions(astrParms, astrBarNames, aScanExpr, strError) Then
                bReturn = True
            Else
                InfBox "An error occured with the engine verification:||" & strError & "|", , , "Engine Verification Error", , , , , , , , eGDAlign_Left
            End If
            
            ' Clear the expression evaluator when done with it
            SetupExpressions astrParms
        End If
    End If
    
    EngineVerify = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCustomCondition.EngineVerify"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitializeEditor
'' Description: Initialize the TradeSense editor
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitializeEditor()
On Error GoTo ErrSection:

    With tsCondition
        .AppPath = App.Path
        .FunctionsRef = g.Functions
        .Lists = m.ListLoading.Lists
        .DisableEnterKey = True
        .ShowNewFunction = False
        .Usage = 10 '8 ' 1=MM; 2=System; 4=Charting; 8=Criteria
        .TurnOnEditing
        .Refresh
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomCondition.InitializeEditor"
    
End Sub

