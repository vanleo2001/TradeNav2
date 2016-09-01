VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmNextBar 
   Caption         =   "Orders for the Next Bar"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8415
   Icon            =   "frmNextBar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8415
   Begin HexUniControls.ctlUniFrameWL fraStrategies 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6075
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
      Caption         =   "frmNextBar.frx":000C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmNextBar.frx":0038
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmNextBar.frx":0058
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboStrategies 
         Height          =   315
         Left            =   1980
         TabIndex        =   2
         Top             =   0
         Width           =   2955
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
         Tip             =   "frmNextBar.frx":0074
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmNextBar.frx":0094
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblStrategies 
         Height          =   195
         Left            =   0
         Top             =   60
         Width           =   2055
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
         Caption         =   "frmNextBar.frx":00B0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmNextBar.frx":0104
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNextBar.frx":0124
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5100
      Width           =   8175
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
      Caption         =   "frmNextBar.frx":0140
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmNextBar.frx":016C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmNextBar.frx":018C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Cancel          =   -1  'True
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   0
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
         Caption         =   "frmNextBar.frx":01A8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmNextBar.frx":01CE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmNextBar.frx":01EE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPrint 
         Height          =   375
         Left            =   1260
         TabIndex        =   8
         Top             =   0
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
         Caption         =   "frmNextBar.frx":020A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmNextBar.frx":0236
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmNextBar.frx":0256
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSave 
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   0
         Width           =   1395
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
         Caption         =   "frmNextBar.frx":0272
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmNextBar.frx":02AC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmNextBar.frx":02CC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClipboard 
         Height          =   375
         Left            =   4140
         TabIndex        =   10
         Top             =   0
         Width           =   1575
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
         Caption         =   "frmNextBar.frx":02E8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmNextBar.frx":032C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmNextBar.frx":034C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkDisplayInUnits 
         Height          =   255
         Left            =   6060
         TabIndex        =   11
         Top             =   60
         Width           =   2055
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
         Caption         =   "frmNextBar.frx":0368
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmNextBar.frx":03BA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNextBar.frx":03DA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsIndexTab vst 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7858
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   0
      FrontTabForeColor=   8388608
      Caption         =   "&Consolidated Orders|Signals for each &Rule"
      Align           =   0
      Appearance      =   1
      CurrTab         =   1
      FirstTab        =   0
      Style           =   0
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   0   'False
      TabsPerPage     =   2
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      Begin vsOcx6LibCtl.vsElastic vseRuleBased 
         Height          =   4080
         Left            =   45
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   7197
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   1
         BorderWidth     =   2
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin RichTextLib.RichTextBox rtbRuleBased 
            Height          =   4020
            Left            =   30
            TabIndex        =   5
            Top             =   30
            Width           =   7965
            _ExtentX        =   14049
            _ExtentY        =   7091
            _Version        =   393217
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            TextRTF         =   $"frmNextBar.frx":03F6
         End
      End
      Begin vsOcx6LibCtl.vsElastic vseConsolidated 
         Height          =   4080
         Left            =   -8670
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   330
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   7197
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   1
         BorderWidth     =   2
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin RichTextLib.RichTextBox rtbConsolidated 
            Height          =   4020
            Left            =   30
            TabIndex        =   4
            Top             =   30
            Width           =   7965
            _ExtentX        =   14049
            _ExtentY        =   7091
            _Version        =   393217
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            TextRTF         =   $"frmNextBar.frx":047C
         End
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmNextBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmNextBar.frm
'' Description: Show the user the orders for the next bar
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Rules As cRules
    strSymbol As String
    strNextBarFile As String
    strSystemName As String
    strSecType As String
    Orders As cRichText
    Signals As cRichText
    dTickMove As Double
    dMinMoveInTicks As Double
    dTickValue As Double
    bIsIntraday As Boolean
    strTimeZoneInfo As String
    
    astrNextBarFiles As cGdArray
    astrStrategies As cGdArray
    
    lIndex As Long
    strFont As String
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_Position = 2
    eGDCol_RuleID = 3
    eGDCol_OrderType = 4
    eGDCol_Price1 = 5
    eGDCol_Offset1 = 6
    eGDCol_Price2 = 7
    eGDCol_Offset2 = 8
    eGDCol_NumContracts = 9
End Enum

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboStrategies_Click
'' Description: Only show the strategy (or all strategies) the the user selects
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboStrategies_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        If cboStrategies.Text = " ALL Strategies" Then
            m.strNextBarFile = ""
            DoReport
        Else
            m.strNextBarFile = m.astrNextBarFiles(cboStrategies.ItemData(cboStrategies.ListIndex))
            DoReport
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.cboStrategies.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkDisplayInUnits_Click
'' Description: If the user clicks in the Display In Units check box, toggle
''              the display between trading units and decimal
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkDisplayInUnits_Click()
On Error GoTo ErrSection:
    
    If Me.Visible Then
        DoReport
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.chkDisplayUnits.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClipboard_Click
'' Description: If the user clicks on the Clipboard button, copy the contents
''              of the current rich text box to the clipboard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClipboard_Click()
On Error GoTo ErrSection:
    
    Dim strMsg As String                ' Message to display to the user
    
    Clipboard.Clear
    If vst.CurrTab = 0 Then
        Clipboard.SetText rtbConsolidated.TextRTF, vbCFRTF
    Else
        Clipboard.SetText rtbRuleBased.TextRTF, vbCFRTF
    End If
    
    strMsg = "You can now paste the report into |another application by selecting |'Edit-Paste'  (or hit 'Ctrl-V')."
    InfBox strMsg, "i"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.cmdClipboard.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user clicks on the OK button, unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPrint_Click
'' Description: If the user clicks on the Print button, bring up the print
''              preview screen with the contents of the current rich text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPrint_Click()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "SNV Report", frmNextBar, 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.cmdPrint.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSave_Click
'' Description: If the user clicks on the Save button, save the current rich
''              text box to file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSave_Click()
On Error GoTo ErrSection:
    
    Dim strFile As String               ' Filename to save the file as
    
    'strFile = AddSlash(App.Path) & "ORDERS.RTF"
    strFile = AddSlash(App.Path) & StripStr(m.strSystemName, ":\/*?|><" & Chr(34)) & ".RTF"
    strFile = CommonDialogFile(frmMain.CommonDialog1, True, "RTF Files (*.rtf)|*.rtf", strFile)
    If Len(strFile) > 0 Then
        If vst.CurrTab = 0 Then
            rtbConsolidated.SaveFile strFile
        Else
            rtbRuleBased.SaveFile strFile
        End If
        
        InfBox "i=i ; h=Next Bar Report ; Saved to file ...|" & strFile
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.cmdSave.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Whenever the form gets activated, make sure that the OK button
''              has the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:
    
    MoveFocus cmdOK
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.Form.Activate", eGDRaiseError_Show
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
    RaiseError "frmNextBar.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form gets loaded, place the form where it is supposed
''              to go, make sure that the first tab is showing, and get the
''              value for the Display In Units check box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Dim strText As String               ' Value from the ini file
    
    Me.Icon = Picture16(ToolbarIcon("ID_Orders"), , True)
    
    g.Styler.StyleForm Me
    
    strText = GetIniFileProperty("NextBar", "", "Placement", g.strIniFile)
    m.strFont = GetIniFileProperty("NextBar", "", "Fonts", g.strIniFile)
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText, "LHTW"
    End If
    
    vst.CurrTab = 0
    vst.FrontTabForeColor = &H8000000D        'flag this to be set to blue for classic & light color scheme
    
    chkDisplayInUnits.Value = GetIniFileProperty("InUnits", vbChecked, "Misc", g.strIniFile)
    
    mnuPopUp.Visible = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, save some properties and clear out
''              some variables
'' Inputs:      Whether or not to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:
    
    SetIniFileProperty "NextBar", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "NextBar", m.strFont, "Fonts", g.strIniFile
    SetIniFileProperty "InUnits", chkDisplayInUnits.Value, "Misc", g.strIniFile
    SetIniFileProperty "NextBarTab", Me.vst.CurrTab, "Misc", g.strIniFile
    
    Set m.Orders = Nothing
    Set m.Signals = Nothing
    Set m.astrNextBarFiles = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: When the form is resized, resize the controls accordingly
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinWidth&, lMinHeight&
    
    lMinWidth = fraButtons.Width + (fraStrategies.Left * 2)
    lMinHeight = fraButtons.Height * 10
    If LimitFormSize(Me, lMinWidth, lMinHeight) Then Exit Sub
    
    With fraButtons
        .Move fraStrategies.Left, ScaleHeight - fraStrategies.Top - .Height
    End With
    
    With vst
        If fraStrategies.Visible Then
            .Move .Left, fraStrategies.Height + (fraStrategies.Top * 2), _
                        ScaleWidth - .Left * 2, _
                        ScaleHeight - fraButtons.Height - fraStrategies.Height - (fraStrategies.Top * 4)
        Else
            .Move .Left, fraStrategies.Top, ScaleWidth - .Left * 2, ScaleHeight - fraButtons.Height - (fraStrategies.Top * 3)
        End If
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitRT
'' Description: Initialize rich text objects
'' Inputs:      Rich Text to initialze, Rich Text Box to put it in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitRT(RT As cRichText, rtb As RichTextBox)
On Error GoTo ErrSection:
    
    Dim strText As String               ' Text to put into the rich text
    
    With RT
        'assign control
        .RTBox = rtb
        With .RTBox
            If Len(m.strFont) > 0 Then
                FontFromString .Font, m.strFont
            Else
                .Font.Name = "Arial"
                .Font.Size = 8
            End If
        End With
        
        ' disclaimer
        strText = "WARNING:  Information provided is for educational and informational purposes only.  Displayed information is not to be construed as trading recommendations by Genesis Financial Data Services, it's employees or affiliates.  Utilization of displayed information shall be at the user's own risk." & vbCrLf & vbCrLf
        .CreateCustomFormat 1
        With .RTBox
            .SelFontName = "Arial"
            .SelFontSize = 8
            .SelItalic = True
            .SelBold = False
            .SelColor = RGB(64, 64, 64)
        End With
        .AddText strText, rtfUseCustomFormat + 1
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.InitRT", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPosition
'' Description: Parse the "Current Position" line from the report
'' Inputs:      Rich Text, Buffer from the report
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CurrentPosition(RT As cRichText, ByVal strBuffer As String)
On Error GoTo ErrSection:
    
    Dim strField1 As String             ' First field in the string
    Dim strField2 As String             ' Second field in the string
    Dim strField3 As String             ' Third field in the string
    Dim strRule As String               ' Rule Name
    Dim strNumContracts As String       ' Number of contracts entered
    Dim dDate As Double                 ' Date to display
    
    With RT
        '1: Current position - this will be "L" or "S" or "N"
        '2: Entry Date - #YYYY-MM-DD {HH:MM:SS}#
        '3: Entry Price - (%g format)
        '4: Entry Rule Name - "Rule Name"
        strRule = Parse(strBuffer, vbTab, 5)
        strField2 = Parse(strBuffer, vbTab, 3)
        'strField2 = Format(Val(StripStr(strField2, "#")))
        dDate = Val(strField2)
        If m.bIsIntraday And g.bShowInLocalTimeZone Then
            dDate = ConvertTimeZone(dDate, m.strTimeZoneInfo, "")
        End If
        strField2 = DateFormat(dDate, MM_DD_YYYY, H_MM, AMPM_UPPER, True)
        strField3 = Parse(strBuffer, vbTab, 4)
        strField3 = PriceDisplay(strField3)
        strNumContracts = NumContracts(Parse(strBuffer, vbTab, 2))
        .AddText "  Current Position:   ", rtfItalic
        Select Case UCase(Left(strBuffer, 1))
            Case "L":
                .AddText "Entered LONG " & strNumContracts & " at " & strField3 & " on " & strField2 & vbCrLf
            Case "S":
                .AddText "Entered SHORT " & strNumContracts & " at " & strField3 & " on " & strField2 & vbCrLf
            Case "N":
                .AddText "NONE" & vbCrLf
                strRule = ""
        End Select
        If Len(strRule) > 0 Then
            .AddText vbTab & vbTab & "          ("
            .AddText "Rule:  ", rtfItalic
            .AddText Chr(34) & strRule & Chr(34) & ")" & vbCrLf
        End If
    
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.CurrentPosition", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DatesText
'' Description: Reads in the dates line of the report
'' Inputs:      Rich Text, Buffer from the report
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DatesText(RT As cRichText, ByVal strBuffer As String)
On Error GoTo ErrSection:
    
    Dim strField As String              ' Field from the string
    Dim dDate As Double                 ' Date to display
    
    With RT
        strField = Parse(strBuffer, vbTab, 1)
        dDate = Val(strField)
        If m.bIsIntraday And g.bShowInLocalTimeZone Then
            dDate = ConvertTimeZone(dDate, m.strTimeZoneInfo, "")
        End If
        strField = DateFormat(dDate, MM_DD_YYYY, H_MM, AMPM_UPPER, True)
        .AddText "  " & strField, rtfBold
        
        strField = Parse(strBuffer, vbTab, 2)
        dDate = Val(strField)
        If m.bIsIntraday And g.bShowInLocalTimeZone Then
            dDate = ConvertTimeZone(dDate, m.strTimeZoneInfo, "")
        End If
        strField = DateFormat(dDate, MM_DD_YYYY, H_MM, AMPM_UPPER, True)
        .AddText "    (last complete bar: " & strField
        
        strField = Parse(strBuffer, vbTab, 3)
        strField = PriceDisplay(strField)
        .AddText "  O=" & strField
        
        strField = Parse(strBuffer, vbTab, 4)
        strField = PriceDisplay(strField)
        .AddText "  H=" & strField
        
        strField = Parse(strBuffer, vbTab, 5)
        strField = PriceDisplay(strField)
        .AddText "  L=" & strField
        
        strField = Parse(strBuffer, vbTab, 6)
        strField = PriceDisplay(strField)
        .AddText "  C=" & strField
        
        .AddText ")" & vbCrLf & vbCrLf
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.DatesText", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoReport
'' Description: Perform and show the report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoReport(Optional ByVal bFillCombo As Boolean = False)
On Error GoTo ErrSection:

    Dim bOrders As Boolean              ' Do we have a consolidated report to show?
    Dim lIndex As Long                  ' Index into a for loop

    ' Initialize rich text objects...
    Set m.Orders = New cRichText
    InitRT m.Orders, rtbConsolidated
    Set m.Signals = New cRichText
    InitRT m.Signals, rtbRuleBased
    
    ' If we only need to run one, then do it here...
    If Len(m.strNextBarFile) > 0 Then
        ParseRuleBased m.strNextBarFile
        If Not ParseOrders(m.strNextBarFile) Then
            vst.CurrTab = 1
        End If
        
    ' Otherwise step through the array and process each one...
    Else
        bOrders = False
        For lIndex = 0 To m.astrNextBarFiles.Size - 1
            m.lIndex = lIndex
            ParseRuleBased m.astrNextBarFiles(lIndex), bFillCombo, lIndex < m.astrNextBarFiles.Size - 1
            If ParseOrders(m.astrNextBarFiles(lIndex), lIndex < m.astrNextBarFiles.Size - 1) Then
                bOrders = True
            End If
        Next lIndex
        If Not bOrders Then vst.CurrTab = 1
    End If
    
    ' Now generate the RTF for the controls...
    m.Orders.BuildRTF
    m.Signals.BuildRTF

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.DoReport", eGDRaiseError_Raise

End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParseOrders
'' Description: Parse the Orders file
'' Inputs:      Next Bar File
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ParseOrders(strNextBarFile As String) As Boolean
On Error GoTo ErrSection:
    
    Dim fhInput As Integer              ' File handle for the next bar file
    Dim strBuffer As String             ' Buffer from the input file
    Dim strField1$                      ' Fields from the input file
    Dim strField2$, strField3$          ' Fields from the input file
    Dim bPrinted As Boolean             ' Has the #2 header been printed?
    Dim strText$                        ' Display text to put in rich text box
    Dim lCounter As Long                ' Number of orders
    Dim strBottomPrice$, strTopPrice$   ' Bottom/top of range that was checked
    Dim strTemp As String               ' Temporary string variable
    Dim lIndex As Long                  ' Temporary index variable
    Dim strFrontMonth As String
    
    With m.Orders
    
        If FileLength(strNextBarFile) < 9 Then
            'unavailable (not able to run it)
            .Clear
            .AddText vbCrLf & vbCrLf
            .AddText "   NOTE:  Must see the 'Signals for each Rule' report for potential orders." & vbCrLf, rtfBold
            .AddText "   The orders cannot be consolidated when a potential signal references an 'unknown price'" & vbCrLf
            .AddText "   (e.g. the Next Bar's High, Low or Close, or a secondary market's Next Bar Open)." & vbCrLf
            lCounter = -1
            ParseOrders = False
        Else
            ParseOrders = True
            
            .AddText vbCrLf & "  Strategy: "
            .AddText m.strSystemName & vbCrLf, rtfBold
    
            ' Try to open the next bar report file
            fhInput = FreeFile
            Open strNextBarFile For Input As #fhInput
            
            ' Get info at beginning of file
            lCounter = 0
            Do While Not EOF(fhInput)
                Line Input #fhInput, strBuffer
                strBuffer = Trim(strBuffer)
                If Len(strBuffer) > 0 Then
                    lCounter = lCounter + 1
                    Select Case lCounter
                        Case 1:
                            .AddText "  Symbol Tested:   "
                            strFrontMonth = FrontMonth(m.strSymbol, Val(Parse(strBuffer, vbTab, 1)))
                            If strFrontMonth <> m.strSymbol Then
                                .AddText m.strSymbol & "  (Active Contract: " & strFrontMonth & ")" & vbCrLf & vbCrLf, rtfBold
                            Else
                                .AddText strFrontMonth & vbCrLf & vbCrLf, rtfBold
                            End If
                            
                            .AddText "  ORDERS for: ", rtfBold
                            DatesText m.Orders, strBuffer
                    
                        Case 2:
                            lIndex = 1&
                            strTemp = Parse(strBuffer, "|", lIndex)
                            Do While strTemp <> ""
                                CurrentPosition m.Orders, strTemp
                                lIndex = lIndex + 1&
                                strTemp = Parse(strBuffer, "|", lIndex)
                            Loop
                    
                        Case 3:
                            strBottomPrice = Parse(strBuffer, vbTab, 1)
                            strBottomPrice = PriceDisplay(strBottomPrice)
                            strTopPrice = Parse(strBuffer, vbTab, 2)
                            strTopPrice = PriceDisplay(strTopPrice)
                            Exit Do
                    End Select
                End If
            Loop
            
            ' Walk through the next bar report file
            lCounter = 0&
            Do While Not EOF(fhInput)
                Line Input #fhInput, strBuffer
                strField1 = Parse(strBuffer, vbTab, 1)
                Select Case Trim(strField1)
                    Case "0"
                        lCounter = lCounter + 1&
                        strField2 = Parse(strBuffer, vbTab, 2)
                        strField3 = Parse(strBuffer, vbTab, 3)
                        If Len(strField3) = 0 Then strField3 = strField2
                        strField2 = PriceDisplay(strField2)
                        strField3 = PriceDisplay(strField3)
                        strText = ""
                        If strBottomPrice = strTopPrice Then
                            ' in this case there is no "If ..." - just place the order!
                            .AddText vbCrLf & vbCrLf & "   The bar to place the orders for opened at:  " & strBottomPrice & ""
                        ElseIf strField2 = strField3 Then
                            strText = strField2
                        ElseIf strField2 = strBottomPrice And strField3 = strTopPrice Then
                            ' in this case there is no "If ..." - just place the order!
                        ElseIf strField2 = strBottomPrice Then
                            strText = "less than or equal to " & strField3
                        ElseIf strField3 = strTopPrice Then
                            strText = "greater than or equal to " & strField2
                        Else
                            strText = strField2 & " through " & strField3
                        End If
                        If Len(strText) > 0 Then
                            .AddText vbCrLf & vbCrLf & "  If the "
                            .AddText "Next Bar's Open", rtfBold
                            .AddText " is "
                            .AddText strText, rtfBold
                            .AddText " then..."
                        Else
                            .AddText vbCrLf
                        End If
                        .AddText vbCrLf & vbTab & "Place the following order(s)..." & vbCrLf & vbCrLf, rtfItalic
                    
                    Case "1"
                        bPrinted = False
                        OrderText strBuffer, vbTab
                        
                    Case "2"
                        If bPrinted = False Then
                            .AddText vbTab & vbTab & "If entry is filled, cancel all other existing " _
                                & "orders and place the following order(s)..." & vbCrLf, rtfItalic
                            bPrinted = True
                        End If
                        OrderText strBuffer, vbTab & vbTab
                End Select
            Loop
            If lCounter = 0 Then
                .AddText vbCrLf & vbCrLf & vbTab & "(No orders to place)" & vbCrLf
            End If
            Close #fhInput
        End If
    End With
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmNextBar.ParseOrders", eGDRaiseError_Raise
    Resume ErrExit
    
End Function
#Else
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParseOrders
'' Description: Parse the Orders file
'' Inputs:      Next Bar File
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ParseOrders(strNextBarFile As String, Optional ByVal bAddSep As Boolean = False) As Boolean
On Error GoTo ErrSection:
    
    Dim strField1$                      ' Fields from the input file
    Dim strField2$, strField3$          ' Fields from the input file
    Dim bPrinted As Boolean             ' Has the #2 header been printed?
    Dim strText$                        ' Display text to put in rich text box
    Dim lCounter As Long                ' Number of orders
    Dim strBottomPrice$, strTopPrice$   ' Bottom/top of range that was checked
    Dim strTemp As String               ' Temporary string variable
    Dim lIndex As Long                  ' Temporary index variable
    Dim strFrontMonth As String
    Dim strDataContract As String
    Dim astrFile As New cGdArray        ' Consolidated orders file
    Dim bSkip As Boolean
    
    astrFile.Create eGDARRAY_Strings
    
    With m.Orders
        If astrFile.FromFile(strNextBarFile) Then
            ' TLB 4/24/2015: when doing orders for a big group (i.e. multiple symbols or basket),
            ' let's just skip the ones which are "empty" (i.e. if in no position AND if no orders).
            bSkip = False
            If Len(m.strNextBarFile) = 0 And astrFile.Size <= 4 Then
                If UCase(Parse(astrFile(2), vbTab, 5) = "NONE") Then
                    bSkip = True
                End If
            End If
            
            If astrFile(0) = "N/A" Or astrFile(1) = "N/A" Then
                bSkip = False
                'unavailable (not able to run it)
                If Not bAddSep Then
                    .Clear
                Else
                    .AddText vbCrLf & "  Strategy: "
                    .AddText m.strSystemName & vbCrLf, rtfBold
                End If
                
                .AddText vbCrLf & vbCrLf
                .AddText "   NOTE:  Must see the 'Signals for each Rule' report for potential orders." & vbCrLf, rtfBold
                .AddText "   The orders cannot be consolidated when a potential signal references an 'unknown price'" & vbCrLf
                .AddText "   (e.g. the Next Bar's High, Low or Close, or a secondary market's Next Bar Open)." & vbCrLf
                lCounter = -1
                ParseOrders = False
            ElseIf Not bSkip Then
                ParseOrders = True
                
                .AddText vbCrLf & "  Strategy: "
                .AddText m.strSystemName & vbCrLf, rtfBold
        
                ' First Line: Symbol information
                If Len(astrFile(1)) > 0 Then
                    .AddText "  Symbol Tested:   "
                    'strFrontMonth = FrontMonth(m.strSymbol, Val(Parse(astrFile(1), vbTab, 1)))
                    strFrontMonth = RollSymbolForDate(m.strSymbol, Val(Parse(astrFile(1), vbTab, 1)))
                    strDataContract = RollSymbolForDate(m.strSymbol, Val(Parse(astrFile(1), vbTab, 2)))
                    If strFrontMonth <> m.strSymbol Then
                        If strDataContract = strFrontMonth Then
                            .AddText m.strSymbol & "  (Active Contract: " & strFrontMonth & ")" & vbCrLf & vbCrLf, rtfBold
                        Else
                            .AddText m.strSymbol & "  (Old Contract: " & strDataContract & ", New Active Contract: " & strFrontMonth & ")" & vbCrLf & vbCrLf, rtfBold
                            .AddText "    *** " & m.strSymbol & " is rolling.  Signals have been generated for " & strFrontMonth & ". ***" & vbCrLf & vbCrLf, rtfBold
                        End If
                    Else
                        .AddText strFrontMonth & vbCrLf, rtfBold
                    End If
                    
                    .AddText "  ORDERS for: ", rtfBold
                    DatesText m.Orders, astrFile(1)
                End If
                
                ' Second Line: Current Position information
                If Len(astrFile(2)) > 0 Then
                    lIndex = 1&
                    strTemp = Parse(astrFile(2), "|", lIndex)
                    Do While strTemp <> ""
                        CurrentPosition m.Orders, strTemp
                        lIndex = lIndex + 1&
                        strTemp = Parse(astrFile(2), "|", lIndex)
                    Loop
                End If
                        
                ' Third Line: Price information
                If Len(astrFile(3)) > 0 Then
                    strBottomPrice = Parse(astrFile(3), vbTab, 1)
                    strBottomPrice = PriceDisplay(strBottomPrice)
                    strTopPrice = Parse(astrFile(3), vbTab, 2)
                    strTopPrice = PriceDisplay(strTopPrice)
                End If
                
                ' Rest of the file: Orders information
                For lCounter = 4 To astrFile.Size - 1
                    strField1 = Parse(astrFile(lCounter), vbTab, 1)
                    
                    Select Case Trim(strField1)
                        Case "0"
                            strField2 = Parse(astrFile(lCounter), vbTab, 2)
                            strField3 = Parse(astrFile(lCounter), vbTab, 3)
                            If Len(strField3) = 0 Then strField3 = strField2
                            strField2 = PriceDisplay(strField2)
                            strField3 = PriceDisplay(strField3)
                            strText = ""
                            If strBottomPrice = strTopPrice Then
                                ' in this case there is no "If ..." - just place the order!
                                .AddText vbCrLf & vbCrLf & "   The bar to place the orders for opened at:  " & strBottomPrice & ""
                            ElseIf strField2 = strField3 Then
                                strText = strField2
                            ElseIf strField2 = strBottomPrice And strField3 = strTopPrice Then
                                ' in this case there is no "If ..." - just place the order!
                            ElseIf strField2 = strBottomPrice Then
                                strText = "less than or equal to " & strField3
                            ElseIf strField3 = strTopPrice Then
                                strText = "greater than or equal to " & strField2
                            Else
                                strText = strField2 & " through " & strField3
                            End If
                            If Len(strText) > 0 Then
                                .AddText vbCrLf & vbCrLf & "  If the "
                                .AddText "Next Bar's Open", rtfBold
                                .AddText " is "
                                .AddText strText, rtfBold
                                .AddText " then..."
                            Else
                                .AddText vbCrLf
                            End If
                            .AddText vbCrLf & vbTab & "Place the following order(s)..." & vbCrLf & vbCrLf, rtfItalic
                        
                        Case "1"
                            bPrinted = False
                            OrderText astrFile(lCounter), vbTab
                            
                        Case "2"
                            If bPrinted = False Then
                                .AddText vbTab & vbTab & "If entry is filled, cancel all other existing " _
                                    & "orders and place the following order(s)..." & vbCrLf, rtfItalic
                                bPrinted = True
                            End If
                            OrderText astrFile(lCounter), vbTab & vbTab
                    
                    End Select
                Next lCounter
                
                If astrFile.Size <= 4 Then
                    .AddText vbCrLf & vbCrLf & vbTab & "(No orders to place)" & vbCrLf & vbCrLf
                End If
            End If
        
            ' Add some spacing in case we are doing multiple files...
            If bAddSep And Not bSkip Then
                .AddText vbCrLf
                .AddText "********************************************************************************************************************************"
                .AddText vbCrLf
            End If
        
        End If
    End With
    
ErrExit:
    Set astrFile = Nothing
    Exit Function
    
ErrSection:
    Set astrFile = Nothing
    RaiseError "frmNextBar.ParseOrders", eGDRaiseError_Raise
    Resume ErrExit
    
End Function
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderText
'' Description: Make a display string out of a coded string for a order line
'' Input:       Coded string
'' Returns:     Display string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrderText(ByVal strString As String, ByVal strMargin As String)
On Error GoTo ErrSection:
    
    Dim strPosition As String           ' Exit Long, Enter Short, etc.
    Dim strAction As String             ' BUY, SELL
    Dim strPrice As String              ' Price to Buy or Sell at
    Dim strDirection As String          ' Offset direction
    Dim strOffsetFrom As String         ' What to offset from
    Dim strPrice2 As String             ' Price to Buy or Sell at
    Dim strDirection2 As String         ' Offset direction
    Dim strOffsetFrom2 As String        ' What to offset from
    Dim strOrderType As String          ' MARKET, LIMIT, STOP, etc.
    Dim strRule As String
    Dim strText As String               ' Temporary string
    Dim strNumContracts As String       ' Number of contracts to enter/exit
    
    ' Determine the Position to take and the action
    strText = Parse(strString, vbTab, GDCol(eGDCol_Position))
    Select Case Trim(strText)
        Case "EL"
            strPosition = "Enter Long"
            strAction = "BUY"
        Case "ES"
            strPosition = "Enter Short"
            strAction = "SELL"
        Case "XL"
            strPosition = "Exit Long"
            strAction = "SELL"
        Case "XS"
            strPosition = "Exit Short"
            strAction = "BUY"
        Case Else
            strPosition = ""
            strAction = ""
    End Select
    
    ' Get name of rule
    strText = Parse(strString, vbTab, GDCol(eGDCol_RuleID))
    If Not m.Rules.Found(strText) Then
        strRule = ""
    Else
        strRule = m.Rules.Item(strText).Name
    End If
    
    ' Determine the Price and the Direction
    strText = Parse(strString, vbTab, GDCol(eGDCol_Price1))
    If Val(strText) = -999999 Then strText = ""
    strPrice = PriceDisplay(strText, False, True)
    If Val(strText) = 0 Then
        strPrice = ""
        strDirection = ""
    ElseIf Val(strText) < 0 Then
        strDirection = " below"
    Else
        strDirection = " above"
    End If
    ' Determine what to offset from, if anything
    strText = Parse(strString, vbTab, GDCol(eGDCol_Offset1))
    Select Case Trim(strText)
        Case "ONB"
            strOffsetFrom = " the Next Bar's Open"
        Case "EP"
            strOffsetFrom = " the Entry Price"
        Case Else
            strOffsetFrom = ""
            strDirection = ""
    End Select
    
    ' Determine the Price2 and the Direction2
    strText = Parse(strString, vbTab, GDCol(eGDCol_Price2))
    If Val(strText) = -999999 Then strText = ""
    strPrice2 = PriceDisplay(strText, False, True)
    If Val(strText) = 0 Then
        strPrice2 = ""
        strDirection2 = ""
    ElseIf Val(strText) < 0 Then
        strDirection2 = " below"
    Else
        strDirection2 = " above"
    End If
    ' Determine what to offset from, if anything
    strText = Parse(strString, vbTab, GDCol(eGDCol_Offset2))
    Select Case Trim(strText)
        Case "ONB"
            strOffsetFrom2 = " the Next Bar's Open"
        Case "EP"
            strOffsetFrom2 = " the Entry Price"
        Case Else
            strOffsetFrom2 = ""
            strDirection2 = ""
    End Select
    
    ' Determine the order type
    strText = Parse(strString, vbTab, GDCol(eGDCol_OrderType))
    Select Case Trim(strText)
        Case "S"
            strOrderType = "STOP"
            'If Right(strString, 6) = vbTab & "0" & vbTab & "ONB" Then
            If Parse(strString, vbTab, GDCol(eGDCol_Price1)) = "0" And Parse(strString, vbTab, GDCol(eGDCol_Offset1)) = "ONB" Then
                If Val(Parse(strString, vbTab, 1)) = 1 Then strOrderType = "MARKET"
            End If
        Case "L"
            strOrderType = "LIMIT"
            'If Right(strString, 6) = vbTab & "0" & vbTab & "ONB" Then
            If Parse(strString, vbTab, GDCol(eGDCol_Price1)) = "0" And Parse(strString, vbTab, GDCol(eGDCol_Offset1)) = "ONB" Then
                If Val(Parse(strString, vbTab, 1)) = 1 Then strOrderType = "MARKET"
            End If
        Case "M"
            strOrderType = "MARKET"
        Case "MOC"
            strOrderType = "MARKET ON CLOSE"
        Case "SCO"
            strOrderType = "STOP CLOSE ONLY"
        Case "LCO"
            strOrderType = "LIMIT CLOSE ONLY"
        Case "SWL"
            strOrderType = "STOP with LIMIT"
            If (Parse(strString, vbTab, GDCol(eGDCol_Price1)) = "0" And Parse(strString, vbTab, GDCol(eGDCol_Offset1)) = "ONB") Or _
                (Parse(strString, vbTab, GDCol(eGDCol_Price2)) = "0" And Parse(strString, vbTab, GDCol(eGDCol_Offset2)) = "ONB") Then
                If Val(Parse(strString, vbTab, 1)) = 1 Then strOrderType = "MARKET"
            End If
        Case "SWLCO"
            strOrderType = "STOP with LIMIT CLOSE ONLY"
        Case Else
            strOrderType = ""
    End Select
    If UCase(Left(strOrderType, 6)) = "MARKET" Then
        strOffsetFrom = ""
        strDirection = ""
        strPrice = ""
    End If
    If UCase(Left(strOrderType, 15)) <> "STOP WITH LIMIT" Then
        strPrice2 = ""
        strDirection2 = ""
        strOffsetFrom2 = ""
    End If
    
    strNumContracts = NumContracts(Parse(strString, vbTab, GDCol(eGDCol_NumContracts)))
    
    ' Put the string together
    With m.Orders
        .AddText strMargin & "To " & strPosition & ":     " & strAction
        If Len(strNumContracts) > 0 Then .AddText " " & strNumContracts
        .AddText " at"
        If Len(strPrice) > 0 Then .AddText " " & strPrice
        If Len(strDirection) > 0 Then .AddText strDirection ' & strOffsetFrom
        If Len(strOffsetFrom) > 0 Then .AddText strOffsetFrom
        If strPrice2 = "" Then
            .AddText " " & strOrderType
        Else
            .AddText " " & "STOP, with LIMIT at " & strPrice2
            If Len(strDirection2) > 0 Then .AddText strDirection2 & strOffsetFrom2
            If InStr(UCase(strOrderType), "CLOSE") > 0 Then .AddText ", CLOSE ONLY"
        End If
    
        If Len(strRule) > 0 Then
            .AddText vbCrLf & strMargin & "          ("
            .AddText "Rule:  ", rtfItalic
            .AddText Chr(34) & strRule & Chr(34) & ")"
        End If
    
        .AddText vbCrLf & vbCrLf
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.OrderText", eGDRaiseError_Raise
    Resume ErrExit

End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParseRuleBased
'' Description: Parse the Rule Based Report
'' Inputs:      Rule Based Report File
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ParseRuleBased(ByVal strNextBarFile As String)
On Error GoTo ErrSection:
    
    Dim fhInput As Integer              ' File handle for the next bar file
    Dim strBuffer As String             ' Buffer from the input file
    Dim lCounter As Long                ' Number of orders
    Dim i&, strPosition$, nRule&, strOrderType$, strCondition$, strPrice$
    Dim strText$, strHeader$, strRule$, iRuleStart&
    Dim strAction$, strMargin$, strStop$, strLimit$
    Dim nSection&, bLongEntries As Boolean, bShortEntries As Boolean
    Dim strEntryRules As String
    Dim astrEntries As New cGdArray
    Dim iIndex As Integer
    Dim strTemp As String               ' Temporary string variable
    Dim lIndex As Long                  ' Temporary index variable
    Dim strNumContracts As String       ' Number of contracts to enter/exit
    Dim strFrontMonth As String
    
    astrEntries.Create eGDARRAY_Strings
    
    With m.Signals
    
        .AddText vbCrLf & "  Strategy: "
        .AddText m.strSystemName & vbCrLf, rtfBold
    
        i = Len(AddSlash(FilePath(strNextBarFile)))
        Mid(strNextBarFile, i + 1, 1) = "R"
        
        If FileLength(strNextBarFile) < 9 Then
            'unavailable
            .AddText " (currently unavailable)"
            lCounter = -1
        Else
            ' Try to open the next bar report file
            fhInput = FreeFile
            Open strNextBarFile For Input As #fhInput
            
            ' Get info at beginning of file
            lCounter = 0
            Do While Not EOF(fhInput)
                Line Input #fhInput, strBuffer
                strBuffer = Trim(strBuffer)
                If Len(strBuffer) > 0 Then
                    lCounter = lCounter + 1
                    Select Case lCounter
                        Case 1:
                            .AddText "  Symbol Tested:   "
                            strFrontMonth = FrontMonth(m.strSymbol, Val(Parse(strBuffer, vbTab, 1)))
                            If strFrontMonth <> m.strSymbol Then
                                .AddText m.strSymbol & "  (Active Contract: " & strFrontMonth & ")" & vbCrLf & vbCrLf, rtfBold
                            Else
                                .AddText strFrontMonth & vbCrLf & vbCrLf, rtfBold
                            End If
                            
                            .AddText "  SIGNALS for: ", rtfBold
                            DatesText m.Signals, strBuffer
                        Case 2:
                            lIndex = 1&
                            strTemp = Parse(strBuffer, "|", lIndex)
                            Do While strTemp <> ""
                                CurrentPosition m.Signals, strTemp
                                lIndex = lIndex + 1&
                                strTemp = Parse(strBuffer, "|", lIndex)
                            Loop
                            Exit Do
                    End Select
                End If
            Loop
            
            ' Walk through the next bar report file
            lCounter = 0&
            nRule = 0
            Do While Not EOF(fhInput)
                Line Input #fhInput, strBuffer
                Select Case Parse(strBuffer, vbTab, 1)
                    Case "0" 'Section
                        strHeader = ""
                        Select Case Parse(strBuffer, vbTab, 2)
                            Case "CP":
                                strHeader = "CURRENT SIGNALS ..."
                                nSection = 0
                                'set flags to see if any long and/or short entries
                                'in "Current" section so know whether to print
                                'the conditional sections or not
                                bLongEntries = False
                                bShortEntries = False
                            Case "EL":
                                strHeader = "SIGNALS only after a new LONG ENTRY in this bar ..."
                                nSection = 1
                            Case "ES":
                                strHeader = "SIGNALS only after a new SHORT ENTRY in this bar ..."
                                nSection = -1
                        End Select
                        If Len(strHeader) > 0 Then
                            strHeader = vbCrLf & vbCrLf & "  " & strHeader & vbCrLf
                        End If
                    
                    Case "1" 'Rule
                        strPosition = Parse(strBuffer, vbTab, 2)
                        nRule = Val(Parse(strBuffer, vbTab, 3))
                        strOrderType = Parse(strBuffer, vbTab, 4)
                        strNumContracts = Parse(strBuffer, vbTab, 5)
                        
                    Case "2"
                        strCondition = Trim(Parse(strBuffer, vbTab, 2))
                    
                    Case "3"
                        strPrice = Trim(Parse(strBuffer, vbTab, 2))
                        
                        'get name of rule
                        strRule = ""
                        If UCase(strCondition) = "FALSE" Then
                            strCondition = "FALSE"
                            If 1 Then
                                nRule = 0 'ignore this rule
                            End If
                        End If
                        If Len(strCondition) > 0 And nRule <> 0 Then
                            If m.Rules.Found(CStr(nRule)) Then
                                strRule = m.Rules.Item(CStr(nRule)).Name
                            End If
                        End If
                        
                        'don't bother showing conditional sections if
                        'did not show any entries in that direction
                        If nSection = 1 And Not bLongEntries Then strRule = ""
                        If nSection = -1 And Not bShortEntries Then strRule = ""
                    
                        'if we're actually going to print this ...
                        If Len(strRule) > 0 Then
                            lCounter = lCounter + 1
                        
                            'header (if first rule of this section)
                            If Len(strHeader) > 0 Then
                                .AddText strHeader, rtfBold
                                strHeader = ""
                            End If
                            iRuleStart = .TextLength
                            
                            'strip Market1 out of condition and price text
                            strText = " OF " & UCase(m.strSymbol)
                            Do
                                i = InStr(UCase(strCondition), strText)
                                If i <= 0 Then Exit Do
                                strCondition = Left(strCondition, i - 1) & Mid(strCondition, i + Len(strText))
                            Loop
                            Do
                                i = InStr(UCase(strPrice), strText)
                                If i <= 0 Then Exit Do
                                strPrice = Left(strPrice, i - 1) & Mid(strPrice, i + Len(strText))
                            Loop
                            
                            'determine the Position to take and the action
                            Select Case Trim(strPosition)
                                Case "EL"
                                    If nSection = 0 Then bLongEntries = True
                                    strPosition = "Enter Long"
                                    strAction = "BUY"
                                Case "ES"
                                    If nSection = 0 Then bShortEntries = True
                                    strPosition = "Enter Short"
                                    strAction = "SELL"
                                Case "XL"
                                    strPosition = "Exit Long"
                                    strAction = "SELL"
                                Case "XS"
                                    strPosition = "Exit Short"
                                    strAction = "BUY"
                                Case Else '(error?)
                                    strPosition = ""
                                    strAction = ""
                            End Select
                            
                            'rule
                            strMargin = vbTab
                            .AddText vbCrLf & strMargin & "Rule:  ", rtfItalic
                            .AddText Chr(34) & strRule & Chr(34) & vbCrLf
                            
                            'condition
                            strMargin = strMargin & "     "
                            If UCase(strCondition) <> "TRUE" Then
                                .AddText strMargin & "IF " & PriceDisplay(strCondition, True) & vbCrLf
                            End If
                            
                            'order
                            strStop = ""
                            strLimit = ""
                            Select Case strOrderType
                                Case "M":
                                    strText = "Market"
                                Case "MOC", "MCO":
                                    strText = "Market on CLOSE"
                                Case "S", "SCO":
                                    strText = "STOP"
                                    strStop = strPrice
                                Case "L", "LCO":
                                    strText = "LIMIT"
                                    strLimit = strPrice
                                Case "SWL", "SWLCO":
                                    strText = "STOP with LIMIT"
                                    strStop = Parse(strPrice, vbTab, 1)
                                    strLimit = Parse(strPrice, vbTab, 2)
                            End Select
                            If InStr(strOrderType, "CO") > 0 Then
                                strText = strText & "CLOSE ONLY"
                            End If
                            .AddText strMargin & "To " & strPosition & ":   " ', rtfItalic
                            .AddText strAction & " " & strNumContracts & " at " & strText & vbCrLf
                            strMargin = strMargin & "     "
                            If Len(strStop) > 0 Then
                                .AddText strMargin & "STOP price:  " & PriceDisplay(strStop, True) & vbCrLf
                            End If
                            If Len(strLimit) > 0 Then
                                .AddText strMargin & "LIMIT price:  " & PriceDisplay(strLimit, True) & vbCrLf
                            End If
                            
                            If strCondition = "FALSE" Then
                                .SetFormat rtfCustomColor, rtfTurnOn, iRuleStart
                            End If
                        End If
                        
                        nRule = 0 'so won't accidentally repeat this rule
                    
                    Case "4"    ' Linked entries
                        If Len(strRule) > 0 Then
                            strMargin = strMargin & "     "
                            .AddText strMargin & "Only Exit if Rule of Entry is one of the following: " & vbCrLf
                            strEntryRules = Parse(strBuffer, vbTab, 2)
                            astrEntries.SplitFields strEntryRules, ","
                            strMargin = strMargin & "     "
                            For iIndex = 0 To astrEntries.Size - 1
                                If astrEntries(iIndex) <> "" Then
                                    .AddText strMargin & m.Rules.Item(CStr(astrEntries(iIndex))).Name & vbCrLf
                                End If
                            Next iIndex
                        End If
                    
                End Select
            Loop
            
            If lCounter = 0 Then
                .AddText vbCrLf & vbCrLf & vbTab & "(No signals for this bar)" & vbCrLf
            End If
            Close #fhInput
        End If
        
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.ParseRuleBased", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub
#Else
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParseRuleBased
'' Description: Parse the Rule Based Report
'' Inputs:      Rule Based Report File
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ParseRuleBased(ByVal strNextBarFile As String, Optional ByVal bFillCombo As Boolean = False, Optional ByVal bAddSep As Boolean = False)
On Error GoTo ErrSection:
    
    Dim lCounter As Long                ' Number of orders
    Dim i&, strPosition$, nRule&, strOrderType$, strCondition$, strPrice$
    Dim strText$, strHeader$, strRule$, iRuleStart&
    Dim strAction$, strMargin$, strStop$, strLimit$
    Dim nSection&, bLongEntries As Boolean, bShortEntries As Boolean
    Dim strEntryRules As String
    Dim astrEntries As New cGdArray
    Dim iIndex As Integer
    Dim strTemp As String               ' Temporary string variable
    Dim lIndex As Long                  ' Temporary index variable
    Dim strNumContracts As String       ' Number of contracts to enter/exit
    Dim strFrontMonth As String
    Dim strDataContract As String
    Dim astrFile As New cGdArray        ' Rule Based file
    Dim astrHeader As New cGdArray
    
    astrEntries.Create eGDARRAY_Strings
    astrFile.Create eGDARRAY_Strings
    astrHeader.Create eGDARRAY_Strings
    
    With m.Signals
    
        i = Len(AddSlash(FilePath(strNextBarFile)))
        Mid(strNextBarFile, i + 1, 1) = "R"
        
        If FileLength(strNextBarFile) < 9 Then
            'unavailable
            .AddText " (currently unavailable)"
            lCounter = -1
        Else
            If Not astrFile.FromFile(strNextBarFile) Then
                .AddText " (currently unavailable)"
                lCounter = -1
            Else
                astrHeader.SplitFields astrFile(0), vbTab
                m.strSymbol = astrHeader(TradesHdrField(eTradesHeader_Symbol))
                m.strSecType = astrHeader(TradesHdrField(eTradesHeader_SecurityType))
                m.strSystemName = astrHeader(TradesHdrField(eTradesHeader_SystemName))
                m.dTickMove = Val(astrHeader(TradesHdrField(eTradesHeader_TickMove)))
                m.dMinMoveInTicks = Val(astrHeader(TradesHdrField(eTradesHeader_MinMoveInTicks)))
                m.dTickValue = Val(astrHeader(TradesHdrField(eTradesHeader_TickValue)))
                m.bIsIntraday = IsIntraday(GetPeriodicity(astrHeader(TradesHdrField(eTradesHeader_BarTimeFrame))))
                m.strTimeZoneInfo = astrHeader(TradesHdrField(eTradesHeader_TimeZoneInfo))
                
                .AddText vbCrLf & "  Strategy: "
                .AddText m.strSystemName & vbCrLf, rtfBold
                
                If bFillCombo Then
                    cboStrategies.AddItem m.strSystemName & " (" & m.strSymbol & ")"
                    cboStrategies.ItemData(cboStrategies.NewIndex) = m.lIndex
                End If
    
                ' First line: Symbol information
                .AddText "  Symbol Tested:   "
                'strFrontMonth = FrontMonth(m.strSymbol, Val(Parse(astrFile(1), vbTab, 1)))
                strFrontMonth = RollSymbolForDate(m.strSymbol, Val(Parse(astrFile(1), vbTab, 1)))
                strDataContract = RollSymbolForDate(m.strSymbol, Val(Parse(astrFile(1), vbTab, 2)))
                If strFrontMonth <> m.strSymbol Then
                    If strDataContract = strFrontMonth Then
                        .AddText m.strSymbol & "  (Active Contract: " & strFrontMonth & ")" & vbCrLf & vbCrLf, rtfBold
                    Else
                        .AddText m.strSymbol & "  (Old Contract: " & strDataContract & ", New Active Contract: " & strFrontMonth & ")" & vbCrLf & vbCrLf, rtfBold
                        .AddText "    *** " & m.strSymbol & " is rolling.  Signals have been generated for " & strFrontMonth & ". ***" & vbCrLf & vbCrLf, rtfBold
                    End If
                Else
                    .AddText strFrontMonth & vbCrLf, rtfBold
                End If
                
                .AddText "  SIGNALS for: ", rtfBold
                DatesText m.Signals, astrFile(1)
            
                ' Second line: Current position(s)
                lIndex = 1&
                strTemp = Parse(astrFile(2), "|", lIndex)
                Do While strTemp <> ""
                    CurrentPosition m.Signals, strTemp
                    lIndex = lIndex + 1&
                    strTemp = Parse(astrFile(2), "|", lIndex)
                Loop
            
                ' Rest of file: Next Bar Orders
                nRule = 0
                For lCounter = 3 To astrFile.Size - 1
                    strText = astrFile(lCounter)
                    strText = Replace(strText, "Bonds of TQ-067", "Close of TQ-067")
                    astrFile(lCounter) = Replace(strText, "Gold of GC-067", "Close of GC-067")
                    Select Case Parse(astrFile(lCounter), vbTab, 1)
                        Case "0" 'Section
                            strHeader = ""
                            Select Case Parse(astrFile(lCounter), vbTab, 2)
                                Case "CP":
                                    strHeader = "CURRENT SIGNALS ..."
                                    nSection = 0
                                    'set flags to see if any long and/or short entries
                                    'in "Current" section so know whether to print
                                    'the conditional sections or not
                                    bLongEntries = False
                                    bShortEntries = False
                                Case "EL":
                                    strHeader = "SIGNALS only after a new LONG ENTRY in this bar ..."
                                    nSection = 1
                                Case "ES":
                                    strHeader = "SIGNALS only after a new SHORT ENTRY in this bar ..."
                                    nSection = -1
                            End Select
                            If Len(strHeader) > 0 Then
                                strHeader = vbCrLf & vbCrLf & "  " & strHeader & vbCrLf
                            End If
                    
                        Case "1" 'Rule
                            strPosition = Parse(astrFile(lCounter), vbTab, 2)
                            nRule = Val(Parse(astrFile(lCounter), vbTab, 3))
                            strOrderType = Parse(astrFile(lCounter), vbTab, 4)
                            strNumContracts = NumContracts(Parse(astrFile(lCounter), vbTab, 5))
                        
                        Case "2"
                            strCondition = Trim(Parse(astrFile(lCounter), vbTab, 2))
                    
                        Case "3"
                            strPrice = Trim(Parse(astrFile(lCounter), vbTab, 2))
                            
                            'get name of rule
                            strRule = ""
                            If UCase(strCondition) = "FALSE" Then
                                strCondition = "FALSE"
                                If 1 Then
                                    nRule = 0 'ignore this rule
                                End If
                            End If
                            If Len(strCondition) > 0 And nRule <> 0 Then
                                If m.Rules.Found(CStr(nRule)) Then
                                    strRule = m.Rules.Item(CStr(nRule)).Name
                                End If
                            End If
                            
                            'don't bother showing conditional sections if
                            'did not show any entries in that direction
                            If nSection = 1 And Not bLongEntries Then strRule = ""
                            If nSection = -1 And Not bShortEntries Then strRule = ""
                        
                            'if we're actually going to print this ...
                            If Len(strRule) > 0 Then
                                'header (if first rule of this section)
                                If Len(strHeader) > 0 Then
                                    .AddText strHeader, rtfBold
                                    strHeader = ""
                                End If
                                iRuleStart = .TextLength
                                
                                'strip Market1 out of condition and price text
                                strText = " OF " & UCase(m.strSymbol)
                                Do
                                    i = InStr(UCase(strCondition), strText)
                                    If i <= 0 Then Exit Do
                                    strCondition = Left(strCondition, i - 1) & Mid(strCondition, i + Len(strText))
                                Loop
                                Do
                                    i = InStr(UCase(strPrice), strText)
                                    If i <= 0 Then Exit Do
                                    strPrice = Left(strPrice, i - 1) & Mid(strPrice, i + Len(strText))
                                Loop
                                
                                'determine the Position to take and the action
                                Select Case Trim(strPosition)
                                    Case "EL"
                                        If nSection = 0 Then bLongEntries = True
                                        strPosition = "Enter Long"
                                        strAction = "BUY"
                                    Case "ES"
                                        If nSection = 0 Then bShortEntries = True
                                        strPosition = "Enter Short"
                                        strAction = "SELL"
                                    Case "XL"
                                        strPosition = "Exit Long"
                                        strAction = "SELL"
                                    Case "XS"
                                        strPosition = "Exit Short"
                                        strAction = "BUY"
                                    Case Else '(error?)
                                        strPosition = ""
                                        strAction = ""
                                End Select
                                
                                'rule
                                strMargin = vbTab
                                .AddText vbCrLf & strMargin & "Rule:  ", rtfItalic Or rtfBold
                                .AddText Chr(34) & strRule & Chr(34) & vbCrLf, rtfBold
                                
                                'condition
                                strMargin = strMargin & "     "
                                If UCase(strCondition) <> "TRUE" Then
                                    .AddText strMargin & "IF " & PriceDisplay(strCondition, True) & vbCrLf
                                End If
                                
                                'order
                                strStop = ""
                                strLimit = ""
                                Select Case strOrderType
                                    Case "M":
                                        strText = "MARKET"
                                    Case "MOC", "MCO":
                                        strText = "MARKET ON CLOSE"
                                    Case "S", "SCO":
                                        strText = "STOP"
                                        strStop = strPrice
                                    Case "L", "LCO":
                                        strText = "LIMIT"
                                        strStop = strPrice
                                    Case "SWL", "SWLCO":
                                        strText = "STOP with LIMIT"
                                        strStop = Parse(astrFile(lCounter), vbTab, 2)
                                        strLimit = Parse(astrFile(lCounter), vbTab, 3)
                                End Select
                                If InStr(strOrderType, "CO") > 0 Then
                                    strText = strText & " CLOSE ONLY"
                                End If
                                .AddText strMargin & "To " & strPosition & ":   " ', rtfItalic
                                '.AddText strAction & " " & strNumContracts & " at " & strText & vbCrLf
                                'strMargin = strMargin & "     "
                                'If Len(strStop) > 0 Then
                                '    .AddText strMargin & "STOP price:  " & PriceDisplay(strStop, True) & vbCrLf
                                'End If
                                'If Len(strLimit) > 0 Then
                                '    .AddText strMargin & "LIMIT price:  " & PriceDisplay(strLimit, True) & vbCrLf
                                'End If
                                .AddText strAction & " " & strNumContracts & " at"
                                If Len(strStop) > 0 Then .AddText " " & PriceDisplay(strStop, False, True)
                                If Len(strLimit) = 0 Then
                                    .AddText " " & strText & vbCrLf
                                Else
                                    .AddText " STOP, with LIMIT at " & PriceDisplay(strLimit, False, True)
                                    If InStr(strOrderType, "CO") Then .AddText ", CLOSE ONLY"
                                    .AddText vbCrLf
                                End If
                                
                                If strCondition = "FALSE" Then
                                    .SetFormat rtfCustomColor, rtfTurnOn, iRuleStart
                                End If
                            End If
                            
                            nRule = 0 'so won't accidentally repeat this rule
                        
                        Case "4"    ' Linked entries
                            If Len(strRule) > 0 Then
                                strMargin = strMargin & "     "
                                .AddText strMargin & "Only Exit if Rule of Entry is one of the following: " & vbCrLf
                                strEntryRules = Parse(astrFile(lCounter), vbTab, 2)
                                astrEntries.SplitFields strEntryRules, ","
                                strMargin = strMargin & "     "
                                For iIndex = 0 To astrEntries.Size - 1
                                    If astrEntries(iIndex) <> "" Then
                                        .AddText strMargin & m.Rules.Item(CStr(astrEntries(iIndex))).Name & vbCrLf
                                    End If
                                Next iIndex
                            End If
                        
                    End Select
                Next lCounter
            End If
            
            If astrFile.Size <= 3 Then
                .AddText vbCrLf & vbCrLf & vbTab & "(No signals for this bar)" & vbCrLf & vbCrLf
            End If
        End If
        
        ' Add some spacing in case we are doing multiple files...
        If bAddSep Then
            .AddText vbCrLf
            .AddText "********************************************************************************************************************************"
            .AddText vbCrLf
        End If
        
    End With

ErrExit:
    Set astrFile = Nothing
    Exit Sub
    
ErrSection:
    Set astrFile = Nothing
    RaiseError "frmNextBar.ParseRuleBased", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Generate the print preview screen for the orders reports
'' Inputs:      Argument(s) passed back and forth
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    With frmPrintPreview.vp
        If frmPrintPreview.GoingToFile And .ExportFormat = vpxRTF Then
            If vst.CurrTab = 0 Then
                rtbConsolidated.SaveFile .ExportFile
            Else
                rtbRuleBased.SaveFile .ExportFile
            End If
        Else
            .StartDoc
            DoPrintHeader
            
            If frmPrintPreview.GoingToFile = False Then
                If vst.CurrTab = 0 Then
                    .Text = rtbConsolidated.TextRTF
                Else
                    .Text = rtbRuleBased.TextRTF
                End If
            ElseIf frmPrintPreview.GoingToFile = True And .ExportFormat <> vpxRTF Then
                If vst.CurrTab = 0 Then
                    .Text = rtbConsolidated.Text
                Else
                    .Text = rtbRuleBased.Text
                End If
            End If
            .Text = vbCrLf
            
            .EndDoc
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.GenerateReport", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PriceDisplay
'' Description: Format the price to display
'' Inputs:      String to Format, Units, Absolute Value
'' Returns:     Formatted String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PriceDisplay(ByVal pstrString As String, _
        Optional ByVal bInUnitsOnlyIfEvenTick As Boolean = False, _
        Optional ByVal bAbs As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strPrice As String
    Dim strPrev As String
    Dim aFields As New cGdArray
    Dim lIndex As Long
    Dim bInUnits As Boolean
    Dim bConvert As Boolean
    Dim dMinMove As Double
    Dim dPrice As Double
    
'TLB 10/15/03: I now think this should be on all the time?
bInUnitsOnlyIfEvenTick = True
    
    If chkDisplayInUnits = vbChecked Then bInUnits = True
    dMinMove = m.dTickMove * m.dMinMoveInTicks
    
    aFields.SplitFields pstrString, " "
    
    For lIndex = 0 To aFields.Size - 1
        strPrice = aFields(lIndex)
        If Left(strPrice, 1) = "(" Then strPrice = Mid(strPrice, 2)
        If Right(strPrice, 1) = ")" Then strPrice = Left(strPrice, Len(strPrice) - 1)
        If TextIsNumeric(strPrice) Then
            dPrice = Val(strPrice)
            strPrice = CStr(dPrice)
            If bInUnits Then
                bConvert = False
                ' don't convert a number just before or after a multiplication or division operator
                If aFields(lIndex - 1) = "*" Or aFields(lIndex - 1) = "/" Then
                    bConvert = False
                ElseIf aFields(lIndex + 1) = "*" Or aFields(lIndex + 1) = "/" Then
                    bConvert = False
                ElseIf bInUnitsOnlyIfEvenTick Then
                    'only if an even tick
                    If dMinMove <> 0 Then
                        If dPrice / dMinMove = Int(dPrice / dMinMove) Then
                            bConvert = True
                        End If
                    End If
                Else
                    bConvert = True
                End If
                If bConvert Then
                    strPrice = gdFormatPrice(dPrice, m.dTickMove, m.dMinMoveInTicks, 0)
                End If
            End If
            If bAbs And Left(strPrice, 1) = "-" Then strPrice = Mid(strPrice, 2)
            If Left(aFields(lIndex), 1) = "(" Then strPrice = "(" & strPrice
            If Right(aFields(lIndex), 1) = ")" Then strPrice = strPrice & ")"
            aFields(lIndex) = strPrice
        End If
    Next lIndex
    
    PriceDisplay = aFields.JoinFields(" ")
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmNextBar.PriceDisplay", eGDRaiseError_Raise
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsNumeric
'' Description: Determines whether the string passed in is numeric or not
'' Inputs:      String to test
'' Returns:     True if numeric, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsNumeric(ByVal pstrString) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    IsNumeric = True
    For lIndex = 1 To Len(pstrString)
        If InStr(".-0123456789", Mid(pstrString, lIndex, 1)) = 0 Then
            IsNumeric = False
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmNextBar.IsNumeric", eGDRaiseError_Raise
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FrontMonth
'' Description: Determine the front month for the continuous contract
'' Inputs:      Symbol being Traded, Date being Calculated for
'' Returns:     Front Month Contract
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FrontMonth(ByVal strSymbol As String, ByVal dDate As Double) As String
On Error GoTo ErrSection:

    Dim strPath As String
    Dim astrFile As New cGdArray
    Dim iPos As Long
    Dim strReturn As String
    Dim lIndex As Long
    Dim Rolls As New cGdTable
    Dim SymInf As vbSymbolInfo
    Dim Bars As New cGdBars
    
    Bars.Prop(eBARS_Symbol) = strSymbol
    strSymbol = Bars.Prop(eBARS_Symbol)

    Set Rolls = GetRollsTable(strSymbol)
    iPos = -1
    For lIndex = 0 To Rolls.NumRecords - 1
        If Rolls(1, lIndex) > dDate Then
            If lIndex = 0 Then
                iPos = 0
            Else
                iPos = lIndex - 1
            End If
            Exit For
        End If
    Next lIndex
    If iPos = -1 Then iPos = Rolls.NumRecords - 1
    If SU_GetSymbolInf(Rolls(0, iPos), SymInf) Then
        strReturn = Parse(SymInf.Symbol, "-", 2)
    End If
    
    If strReturn <> "" Then
        If InStr(strSymbol, "-") Then
            Bars.Prop(eBARS_Symbol) = Parse(strSymbol, "-", 1) & "-" & strReturn
        ElseIf InStr(strSymbol, "/") Then
            Bars.Prop(eBARS_Symbol) = Parse(strSymbol, "/", 1) & "/" & strReturn
        End If
        FrontMonth = Bars.Prop(eBARS_Symbol)
    Else
        FrontMonth = strSymbol
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmNextBar.FrontMonth", eGDRaiseError_Raise
    Resume ErrExit

End Function

Public Function ShowMe(ByVal strNextBarFile As String, Optional Rules As cRules = Nothing) As Boolean
On Error GoTo ErrSection:

    Screen.MousePointer = vbHourglass

    Set m.Rules = New cRules
    If Rules Is Nothing Then
        m.Rules.Load
    Else
        Set m.Rules = Rules
    End If
    
    Set m.astrNextBarFiles = New cGdArray

    m.strNextBarFile = strNextBarFile
    DoReport
    
    If GetIniFileProperty("NextBarTab", 0, "Misc", g.strIniFile) = 0 Then
        vst.CurrTab = 0
    Else
        vst.CurrTab = 1
    End If
    
    SetEditorCaption Me, "Next Bar Report", m.strSystemName
    
    fraStrategies.Visible = False
    Screen.MousePointer = 0
    ShowForm Me, False, frmMain ' True
    ShowMe = True
    
ErrExit:
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmNextBar.ShowMe", eGDRaiseError_Raise
    
End Function

Private Function NumContracts(ByVal strNumContracts As String) As String
On Error GoTo ErrSection:

    If Right(strNumContracts, 1) = "%" Then
        NumContracts = strNumContracts & " of Position"
    Else
        Select Case UCase(m.strSecType)
            Case "S"
                NumContracts = Str(CLng(Val(strNumContracts)) * CLng(m.dTickValue / m.dTickMove)) & " Shares"
            
            Case "I"
                NumContracts = Str(CLng(Val(strNumContracts)) * 100)
            
            Case Else
                If CLng(Val(strNumContracts)) = 1 Then
                    NumContracts = strNumContracts & " Contract"
                Else
                    NumContracts = strNumContracts & " Contracts"
                End If
        
        End Select
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmNextBar.NumContracts", eGDRaiseError_Raise
    
End Function

Public Function ShowMeMult(ByVal astrNextBarFiles As cGdArray, Optional Rules As cRules = Nothing, Optional ByVal strCaption As String = "Multiple") As Boolean
On Error GoTo ErrSection:

    Screen.MousePointer = vbHourglass

    Set m.Rules = New cRules
    If Rules Is Nothing Then
        m.Rules.Load
    Else
        Set m.Rules = Rules
    End If
    
    Set m.astrNextBarFiles = New cGdArray
    Set m.astrNextBarFiles = astrNextBarFiles.MakeCopy
    Set m.astrStrategies = New cGdArray
    m.astrStrategies.Create eGDARRAY_Strings
    m.strNextBarFile = ""
    cboStrategies.Clear
    cboStrategies.AddItem " ALL Strategies"
    DoReport True
    cboStrategies.ListIndex = 0
    
    If GetIniFileProperty("NextBarTab", 0, "Misc", g.strIniFile) = 0 Then
        vst.CurrTab = 0
    Else
        vst.CurrTab = 1
    End If
    
    SetEditorCaption Me, "Next Bar Report", strCaption
    
    fraStrategies.Visible = m.astrNextBarFiles.Size > 1
    Form_Resize
    
    Screen.MousePointer = 0
    ShowForm Me, False, frmMain
    ShowMeMult = True
    
ErrExit:
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmNextBar.ShowMemult", eGDRaiseError_Raise
    
End Function

Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    Dim Font As StdFont

    Set Font = rtbConsolidated.Font
    If CommonDialogFont(frmMain.CommonDialog1, Font) = True Then
#If 0 Then
        With rtbConsolidated
            .SelStart = 0
            .SelLength = Len(rtbConsolidated.Text)
            .SelFontName = Font.Name
            .SelFontSize = Font.Size
            .Refresh
            .SelLength = 0
            Set .Font = Font
        End With
#End If
        m.strFont = FontToString(Font)
        DoReport
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub rtbConsolidated_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.rtbConsolidated.MouseDown", eGDRaiseError_Raise
    
End Sub

Private Sub rtbRuleBased_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNextBar.rtbRuleBased.MouseDown", eGDRaiseError_Raise
    
End Sub


