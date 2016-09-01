VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl Editor 
   BackColor       =   &H80000014&
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   ScaleHeight     =   1020
   ScaleWidth      =   7350
   Begin RichTextLib.RichTextBox RuleEditor 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   1191
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"Editor.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "FunctionOptions"
      Visible         =   0   'False
      Begin VB.Menu optCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu optPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu optLine 
         Caption         =   "-"
      End
      Begin VB.Menu optEdit 
         Caption         =   "&Edit Function"
      End
      Begin VB.Menu optRules 
         Caption         =   "Function &Info"
      End
   End
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Public Event Change()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event EditFunction(FunctionID As Long, FunctionName As String, Found As Boolean)
Public Event NewFunction(ByVal lCategoryID As Long)

Private Type mPrivate
    bEditingOn As Boolean
    bExprIsFormatted As Boolean
    bDisableEnterKey As Boolean
    iTabCount As Integer
    lTabWidth As Long
    bHasFocus As Boolean
    
    'Used to control when the current function is bold or not (Right click)
    lFunctionBeg As Long
    lFunctionLength As Long
    
    bFunctionPropertiesShown As Boolean
    bShowNewFunctionButton As Boolean
End Type
Private m As mPrivate

Const C_LastModifiedDate = 1
Const C_Rules = 2
Const C_Desc = 3

Property Let EnterDoesntSelect(pData As Boolean)
    If gEditingArea Is Nothing Then Set gEditingArea = New cEditingArea
    gEditingArea.EnterDoesntSelect = pData
End Property

Property Get Text() As String
    Text = RuleEditor.Text
End Property
Property Get TextRtf() As String
    TextRtf = RuleEditor.TextRtf
End Property

Property Let Text(pData As String)
    RuleEditor.Text = pData
End Property
Property Let TextRtf(pData As String)
    RuleEditor.TextRtf = pData
    m.bExprIsFormatted = True
    SetTabs m.iTabCount, m.lTabWidth
End Property

Property Get TabCnt() As Integer
    TabCnt = m.iTabCount
End Property
Property Get TabWidth() As Long
    TabWidth = m.lTabWidth
End Property
Property Let TabCnt(pData As Integer)
    m.iTabCount = pData
End Property
Property Let TabWidth(pData As Long)
    m.lTabWidth = pData
End Property

'Set these in caller for newly initiated OCX...
Property Let FunctionsRef(pData As cFunctions)
    If gEditingArea Is Nothing Then Set gEditingArea = New cEditingArea
    With gEditingArea
        .FunctionsRef = pData
        .RtfTextBox = RuleEditor
        .EditorCtl = Me
    End With
End Property

'Pass TradeSense lists...
Property Let Lists(pData As cLists)
    If gEditingArea Is Nothing Then Set gEditingArea = New cEditingArea
    gEditingArea.Lists = pData
End Property

Property Let DisableEnterKey(pData As Boolean)
    m.bDisableEnterKey = pData
End Property
Property Let Usage(pData As Byte)
    gEditingArea.Usage = pData
End Property

Property Let Enabled(pData As Boolean)
    RuleEditor.Enabled = pData
End Property
Property Get Enabled() As Boolean
    Enabled = RuleEditor.Enabled
End Property

Property Let Locked(pData As Boolean)
    RuleEditor.Locked = pData
End Property
Property Get Locked() As Boolean
    Locked = RuleEditor.Locked
End Property

Property Get AppPath() As String
    AppPath = g.strAppPath
End Property
Property Let AppPath(ByVal strPath As String)
    g.strAppPath = AppPath
End Property

Property Get ShowNewFunction() As Boolean
    ShowNewFunction = m.bShowNewFunctionButton
End Property
Property Let ShowNewFunction(ByVal bShowNewFunction As Boolean)
    m.bShowNewFunctionButton = bShowNewFunction
End Property

'Sets the default tabs for the editor
Public Sub SetTabs(pTabs As Integer, pTabWidth As Long)
On Error GoTo ErrSection:
    
    Dim X           As Long
    Dim svStart     As Long
    Dim sveditingon As Boolean
    Dim TabWidth    As Long
    Dim retval      As Variant
    
    retval = LockWindowUpdate(RuleEditor.hWnd)
    
    'This prevents Change event from triggering as each tab is set...
    sveditingon = m.bEditingOn
    m.bEditingOn = False
    
    'If Number of Tabs < 0 then default to 25 with width of 250
    If m.iTabCount <= 0 Then
        m.iTabCount = 25
        m.lTabWidth = 250
    End If
    
    'Set tab stops within Editor
    With RuleEditor
        svStart = .SelStart
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelTabCount = m.iTabCount
        For X = 0 To .SelTabCount - 1
            TabWidth = TabWidth + m.lTabWidth
           .SelTabs(X) = TabWidth
        Next X
        .SelLength = 0
        .SelStart = svStart
    End With
    
    m.bEditingOn = sveditingon
      
ErrExit:
    retval = LockWindowUpdate(0)
    Exit Sub

ErrSection:
    RaiseError "TSOCX.Editor.SetTabs", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Private Sub optCopy_Click()
On Error GoTo ErrEnd
    Clipboard.Clear
    If RuleEditor.SelLength > 0 Then
        Clipboard.SetText RuleEditor.SelText, rtfCFText
    End If
Exit Sub
ErrEnd:
End Sub

Private Sub optPaste_Click()
On Error GoTo ErrEnd
    If Clipboard.GetFormat(rtfCFText) Then
        RuleEditor.SelText = Clipboard.GetText
    End If
Exit Sub
ErrEnd:
End Sub

Private Sub RuleEditor_Keydown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
    
    'Exit if user pressed:
    '- a function key
    '- the <alt> key or <ctrl> key (Shift = 2 or 4)
    '- Shift <tab>
    '- Control key (17)
    If (KeyCode >= 112 And KeyCode <= 123) Or Shift = 2 Or Shift = 4 Or _
       (KeyCode = 9 And Shift = 1) Then
        RaiseEvent KeyDown(KeyCode, Shift)
        'Keycode = 0
        Exit Sub
    End If
    
    'If Function list is visible then allow enter key to accept selection,
    'otherwise, if the AllowEntryKey flag is set off then disable
    If m.bDisableEnterKey Then
        If KeyCode = 13 And Not frmFunctionList.Visible Then
            KeyCode = 0
            Exit Sub
        End If
    End If
    
    Select Case KeyCode
        
        'Tilda
        Case 192
            If Shift = 1 Then
                KeyCode = 0
            End If
        
        '222=Single Quote
        'Case 222
            'If Shift = 0 Then
            '    Keycode = 0
            'End If
            
        '8=BackSpace, 9=Tab, 13=Enter, 27=Esc, 38=UpArrow, 40=DownArray
        '33=PgDown,34=Pgup, 32=Space
        Case 8, 33, 34, 27, 34, 38, 40, 32
            gEditingArea.ProcessKeyDown KeyCode
            'Disable tab so that the focus doesn't leave the editor
            If gEditingArea.DisableKey Then
                KeyCode = 0
            End If
    
        '9=Tab, 13=Enter
        Case 9, 13
            gEditingArea.ProcessKeyDown KeyCode
            'Disable tab so that the focus doesn't leave the editor
            If gEditingArea.DisableKey Then
                KeyCode = 0
            End If
        
        '188=comma, 48=RParen
        Case 188, 48
            'RParen requires the Shift Key as well
            If KeyCode = 48 And Shift <> 1 Then Exit Sub
            
            'Comma requires the shift key to not be pressed
            If KeyCode = 188 And Shift <> 0 Then Exit Sub
            
            'if there is a default selection and user hits comma or right paren
            'we don't want to overwrite the selection, rather append the char
            If RuleEditor.SelLength > 0 Then
                'there is a selection - remove it
                Dim selLen As Integer
                selLen = RuleEditor.SelLength
                RuleEditor.SelStart = RuleEditor.SelStart + selLen
                
            End If
    End Select
   
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.Editor.RuleEditor.KeyDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Private Sub RuleEditor_Keyup(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
    
    'Exit if user pressed:
    '- a function key
    '- the <alt> key
    '- Shift <tab>
    '- Control key (17)
    If (KeyCode >= 112 And KeyCode <= 123) Or Shift = 2 Or Shift = 4 Or _
       (KeyCode = 9 And Shift = 1) Then
        RaiseEvent KeyUp(KeyCode, Shift)
        Exit Sub
    End If
    
    '35-36  Home/End
    '37,39  L/R  Arrow keys
    Select Case KeyCode
        Case 35, 36
            If Shift = 0 Then
                gEditingArea.ProcessKeyUp KeyCode
            Else
                'Shift Home/End should mark text for deleting/copying, etc.
            End If
            
        Case 37, 39
            gEditingArea.ProcessKeyUp KeyCode
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.Editor.RuleEditor.KeyUp", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Private Sub RuleEditor_Change()
On Error GoTo ErrSection:

    If Not m.bEditingOn Then Exit Sub
    
    RaiseEvent Change
    
    If m.bExprIsFormatted Or m.bFunctionPropertiesShown Then
        TurnOffFormatting
    End If
    
    frmFunctionInfo.Visible = False
    
    gEditingArea.ProcessChg
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.Editor.RuleEditor.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Private Sub RuleEditor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
    
    Dim svState As Boolean
    
    Dim svEditingOnFlag As Boolean
    
    'focus control to get the focus (so those events trigger) before continuing
    If Not m.bHasFocus Then Exit Sub
    
    If Button = 2 Then
    
        'Force left click event to make sure focus in under correction
        'location
'        DoEvents
'        Call MouseEvent(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
'        Call MouseEvent(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
'        DoEvents
        
        'MT Dec/2001
        'Disable for now.  When an option is chosen from the list.  The rule
        'is no longer colored.  Also, the description line shows up over the
        'inputs line.  Wierd "event" bugs.
        
        'prior to popup, enable/disable appropriatly
        optPaste.Enabled = Clipboard.GetFormat(vbCFText) Or Clipboard.GetFormat(rtfCFText)
        optCopy.Enabled = RuleEditor.SelLength > 0
        
        'from here - we need to know the function in question
        Dim bValidFunction As Boolean, bSecurity As Byte
        If RuleEditor.SelLength > 0 Then
            'user has selected something - use their selection
            bValidFunction = gEditingArea.IsValidFunction(RuleEditor.SelText, bSecurity)
        Else
            'user hasn't selected anything - lets find the closest function
            bValidFunction = gEditingArea.IsValidFunction(RuleEditor.SelText, bSecurity)
        End If
        'user can only edit function based on security level & validity
        optEdit.Enabled = bValidFunction And bSecurity < 2
        'user can only view the info if its a valid function
        optRules.Enabled = bValidFunction
        
        PopupMenu mnuOptions
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.Editor.RuleEditor.MouseDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Private Sub RuleEditor_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Not m.bHasFocus Then Exit Sub
    
'    gEditingArea.ProcessMouseClick
            
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.Editor.RuleEditor.MouseUp", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'This must be run before using control
Public Sub Refresh()
On Error GoTo ErrSection:

    gEditingArea.Refresh
    SetTabs m.iTabCount, m.lTabWidth
   
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.Editor.Refresh", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'Attributes that can be set in Rich text box...
Property Get SelText() As String
    SelText = RuleEditor.SelText
End Property
Property Let SelStart(pData As Long)
    RuleEditor.SelStart = pData
End Property
Property Let SelLength(pData As Long)
    RuleEditor.SelLength = pData
End Property
Property Let SelBold(pData As Boolean)
    RuleEditor.SelBold = pData
End Property
Property Let SelItalic(pData As Boolean)
    RuleEditor.SelItalic = pData
End Property
Property Let SelUnderline(pData As Boolean)
    RuleEditor.SelUnderline = pData
End Property
Property Let SelColor(pData As Long)
    RuleEditor.SelColor = pData
End Property
Property Let SelFontSize(pData As Long)
    RuleEditor.SelFontSize = pData
End Property

'Set after validating a rule
Property Let ExprIsFormatted(pData As Boolean)
    m.bExprIsFormatted = pData
End Property
Property Get ExprIsFormatted() As Boolean
    ExprIsFormatted = m.bExprIsFormatted
End Property

Public Sub RemoveTradeSense()
On Error GoTo ErrSection:

    If gEditingArea Is Nothing Then Exit Sub
    With gEditingArea
        .MakeListInvisible
        .MakeParmLineInvisible
    End With
    frmFunctionInfo.Visible = False
    
    If m.bFunctionPropertiesShown Then
        TurnOffFormatting
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.Editor.RemoveTradeSense", eGDRaiseError_Raise, g.strAppPath

End Sub

'Method to turn off keystroking editing during verification in TradeSense.
Public Sub TurnOffEditing()
On Error GoTo ErrSection:

    m.bEditingOn = False
    RemoveTradeSense
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.Editor.TurnOffEditing", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Public Sub TurnOnEditing()
On Error GoTo ErrSection:
    
    m.bEditingOn = True
            
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.Editor.TurnOnEditing", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Private Sub optEdit_Click()
On Error GoTo ErrSection:
    
    Dim FunctionName        As String
    Dim FunctionID          As Long
    Dim Found               As Boolean
    Dim PreviewRTF          As String
    Dim FuncDescription     As String
    Dim LastModified        As Date
    Dim ImplType            As Byte
    Dim X                   As Long
    Dim Y                   As Long
    
    'select the function in question
    gEditingArea.ProcessMouseClick
    
    'Get function information and pass through event
    gEditingArea.GetFunction FunctionName, FunctionID, Found, _
        PreviewRTF, FuncDescription, LastModified, ImplType, _
        X, Y, m.lFunctionBeg, m.lFunctionLength
    
    RaiseEvent EditFunction(FunctionID, FunctionName, Found)
            
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.Editor.optEdit.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Private Sub optRules_Click()
On Error GoTo ErrSection:

    ShowFunctionInfo C_Rules
    m.bFunctionPropertiesShown = True
            
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.Editor.optRules.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub


Private Sub ShowFunctionInfo(pInfoToShow As Byte)
On Error GoTo ErrSection:

    Dim FunctionName        As String
    Dim FunctionID          As Long
    Dim Found               As Boolean
    Dim Preview             As String
    Dim FuncDescription     As String
    Dim LastModified        As Date
    Dim ImplementationType  As Byte
    Dim retval              As Variant
    Dim X                   As Long
    Dim Y                   As Long

    RemoveTradeSense
    gEditingArea.ShowSelectedFunctionInfo

    Exit Sub

ErrSection:
    retval = LockWindowUpdate(0)
    RaiseError "TSOCX.Editor.ShowFunctionInfo", eGDRaiseError_Raise, g.strAppPath

End Sub
'

Private Sub RuleEditor_GotFocus()
On Error GoTo ErrSection:
    
    Dim svStart     As Long
    
    'Turn off highlighted function when frmFunctionInfo is unloaded
    If m.lFunctionLength > 0 Then
        With RuleEditor
        
            'Save focus.  If changing between multiple Rule Editors, this
            'ensures the spot selected by the user to users AFTER unbolding
            'the previously highlighted function (from right clicking)
            svStart = .SelStart
            .SelStart = m.lFunctionBeg
            .SelLength = m.lFunctionLength
            .SelBold = False
            .SelLength = 0
            .SelStart = svStart     '<<Restore spot user clicking on
        End With
    End If
            
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.Editor.RuleEditor.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Private Sub UserControl_EnterFocus()
    
    m.bHasFocus = True

End Sub

Private Sub UserControl_ExitFocus()
    
    m.bHasFocus = False

End Sub

Private Sub UserControl_InitProperties()

    m.bShowNewFunctionButton = True

End Sub

Private Sub UserControl_Resize()
On Error Resume Next
        
    RuleEditor.Move 0, 0, ScaleWidth, ScaleHeight

End Sub

'Turns off Rich Text Formatting
Public Sub TurnOffFormatting()
On Error GoTo ErrSection:
    
    Dim svLength    As Integer
    Dim svSelStart  As Integer
    Dim retval      As Variant
    
    retval = LockWindowUpdate(frmFunctionInfo.hWnd)
    
    'If a block of text is currently highlighted, save the length
    With RuleEditor
        svSelStart = .SelStart
        svLength = .SelLength
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelProtected = False
        .SelColor = vbBlack
        .SelBold = False
        .SelItalic = False
        .SelStart = svSelStart
        .SelLength = svLength
        MoveFocus RuleEditor
    End With
    
    
    m.bExprIsFormatted = False
    m.bFunctionPropertiesShown = False
    RaiseEvent Change
            
ErrExit:
    retval = LockWindowUpdate(0)
    Exit Sub

ErrSection:
    retval = LockWindowUpdate(0)
    RaiseError "TSOCX.Editor.TurnOffFormatting", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Public Sub NewFunction(ByVal lCategoryID As Long)
On Error GoTo ErrSection:

    RaiseEvent NewFunction(lCategoryID)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.Editor.NewFunction", eGDRaiseError_Raise, g.strAppPath
    
End Sub
