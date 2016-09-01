VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEditingArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private Type mPrivate
    'Properties used to reference controls
    RTF As Object
    Editor As Editor
    InputsLine As cInputsLine
    FunctionList As cFunctionList
    StrH As cStrh
    bEditorOn As Boolean
    bEditorOnTmp As Boolean
    
    'Properties for keystroke parsing
    bDisableKey As Boolean
    iCaretOffset As Integer
    iKeyCode As Integer
    
    'String manipulation variables
    iCaretPos As Integer
    nCurInput As Byte
    strFunctionName As String
    iFuncBegPos As Integer
    iFuncLeftParen As Integer
    iInputLineBeg As Integer
    iInputLineEnd As Integer
    iInputLineLeft As Integer
    lRet As Long
    iSelectionListBeg As Integer
    strTemp As String
    Functions As cFunctions
    Func As cFunction
    lXCor As Long
    lYCor As Long
    Lists As cLists
    bEnterDoesntSelect As Boolean
    
    'var for default parameter value
    iPrevCaretPos As Integer
    nTotalInputs As Byte
    bShowInputsLine As Boolean
    'these vars are used for storing temporary selected functions off the right mouse selection in editor
    sEdFunction As String
    lEdFID As Long
    
    bRefresh As Boolean  ' true during a Refresh call
End Type
Private m As mPrivate

Property Let Usage(pData As Byte)
    m.FunctionList.Usage = pData
End Property
Property Get DisableKey() As Boolean
    DisableKey = m.bDisableKey
End Property
Property Get GetFunctionList() As cFunctionList
    Set GetFunctionList = m.FunctionList
End Property
Property Let EnterDoesntSelect(pData As Boolean)
    m.bEnterDoesntSelect = pData
End Property

'Process non navigation key
Public Sub ProcessChg()
On Error GoTo ErrSection:
    
    Dim nTotInputs As Integer, iHParam As Integer
    Dim sRFinal As String
    Dim sDefValue As String, sFinal As String
    Dim bValidFunc As Boolean, iParams As Integer
    Dim SelFunction As String
    Dim knownOpsStr As String
    Dim bForwardDirection As Boolean
    knownOpsStr = "(,)+-*/"
    
    'Set the mode of the editor.  Turn off tradesense for comment lines...
    SetEditorMode
    
    m.bDisableKey = False
    
    'Get current text and hairline focus position
    m.iCaretPos = m.RTF.SelStart
    'persist the direction we are traveling
    If m.iCaretPos > m.iPrevCaretPos Then bForwardDirection = True
    m.iPrevCaretPos = m.iCaretPos
    
    m.strTemp = m.RTF.Text
    If m.strTemp <> "" Then
        If m.RTF.SelStart = 0 Then
            m.iKeyCode = 0
        Else
            'Backspace?
            If m.iKeyCode = 8 Then
                m.iKeyCode = 0
            Else
                m.iKeyCode = Asc(Mid(m.strTemp, m.RTF.SelStart, 1))
            End If
        End If
    Else
        m.iKeyCode = 0
        Exit Sub
    End If
    
    'Numbers 0-9 dont show selection list
    If m.iKeyCode > 47 And m.iKeyCode < 58 Then
        RemoveFunctionList
    End If
    
    'Show selection list if the user starts typing a function name in an empty edit window
    If m.RTF.SelStart > 0 Then
        Dim sChr As String
        sChr = Mid(m.strTemp, m.RTF.SelStart, 1)
        If Trim(Mid(m.strTemp, 1, m.RTF.SelStart - 1)) = "" And sChr >= "a" And sChr <= "z" _
            And Not frmFunctionList.Visible And m.bEditorOn And m.bEditorOnTmp Then
            m.lRet = GetCaretPos(PT)
            ClientToScreen m.RTF.hWnd, PT
            m.lXCor = PT.X * Screen.TwipsPerPixelX
            m.lYCor = (PT.Y * Screen.TwipsPerPixelY) + 300
            m.FunctionList.ShowNewFunction = m.Editor.ShowNewFunction
            m.FunctionList.Show m.lXCor, m.lYCor
        End If
    End If
    
    'This fixes the problem having the correct tradesense list showing
    'after operators are typed.  When an existing TradeSense list is
    'displayed after typing " ", and then a character operator is typed ("OF,
    '"And", "or") a new tradeSense list should be displayed after the
    'operator.  This code searches for patterns of " " followed by an
    'operator with the tradesense list already displayed.
    If frmFunctionList.Visible Then
        '68="D"     'anD
        '100="d"    'and
        '70="F"     'oF
        '102="f"    'of
        '82="R"     'oR
        '114="r"    'or
        Select Case m.iKeyCode
            Case 68, 70, 82, 100, 102, 114
                If m.iCaretPos - 2 >= 1 Then
                    If Mid(m.strTemp, m.iCaretPos - 2, 3) = " of" Or _
                       Mid(m.strTemp, m.iCaretPos - 2, 3) = " or" Then
                        RemoveFunctionList
                    End If
                End If
                If m.iCaretPos - 3 >= 1 Then
                    If Mid(m.strTemp, m.iCaretPos - 3, 4) = " and" Then
                        RemoveFunctionList
                    End If
                End If
        End Select
    End If
    
    'Key pressed driver
    Select Case m.iKeyCode
        Case 32: SpaceBar
        Case 34: DoubleQuote
        Case 42, 43, 45, 47: MathSymbols
        Case 44: Comma
        Case 40
            LeftParen m.iCaretPos
            'biggie here is that there ARE inputs
            'set our total inputs member, used for determining when done
            FunctionHasParams m.iCaretPos, bValidFunc, iParams
            If bValidFunc And iParams < 1 Then
                'no inputs, but they placed a LParen;
                'remove the LParen and give them some notification
                RemoveChar m.RTF.Text, m.iCaretPos, sRFinal
                m.RTF.Text = sRFinal
                'reset last pos
                m.iCaretPos = m.iCaretPos - 1
                m.RTF.SelStart = m.iCaretPos
                Beep
                Exit Sub
            ElseIf Not bValidFunc Then Exit Sub
            End If
            
        Case 41: RightParen m.iCaretPos
        Case 60, 61, 62: CompareSymbols
        Case 46
            'its a '.' and if the Functionlist is up, set the category to C_BarsAgo
            If frmFunctionList.Visible Then
                frmFunctionList!cbCategory = C_BARSAGO
                Exit Sub
            End If
    End Select
    
    If m.iCaretPos > m.iSelectionListBeg Then
        If frmFunctionList.Visible Then
            'we want to search, but lets remove any operators that might be in first position
'            If InStr(1, knownOpsStr, Left(m.strTemp, 1), vbTextCompare) > 0 Then
'                m.FunctionList.Search Right(m.strTemp, Len(m.strTemp) - 1), m.iCaretPos
'            Else
                m.FunctionList.Search m.strTemp, m.iCaretPos
'            End If
        End If
        ' Comma or left parenthesis
        If bForwardDirection And (m.iKeyCode = 44 Or m.iKeyCode = 40) Then
            
            'get the default param value
            GetDefaultValue m.nCurInput, sDefValue
            
            'only add default value for a comma if there isn't an argument already there
            Dim nPos As Integer
            For nPos = m.iCaretPos + 1 To Len(m.RTF.Text)
                Dim sCurChar As String
                sCurChar = Mid(m.RTF.Text, nPos, 1)
                If sCurChar = "," Or sCurChar = ")" Then Exit For
                If sCurChar <> " " Then
                    sDefValue = ""
                    Exit For
                End If
            Next
            
            'insert def value into text stream
            If Len(sDefValue) > 0 Then
                InsertString m.RTF.Text, " " & sDefValue & " ", m.iCaretPos, sFinal
                m.RTF.Text = sFinal 'm.RTF.Text & " " & sDefValue
                Dim nSelLen As Integer
                nSelLen = Len(sDefValue)
            
                ' Left paren?
                If m.iKeyCode = 40 Then
                    ' Fill in the rest of the default values
                    nPos = m.iCaretPos + Len(sDefValue) + 1
                    Dim nInput As Byte
                    nInput = m.nCurInput + 1
                    GetDefaultValue nInput, sDefValue
                    Do While Len(sDefValue) > 0
                        'string together parms separated by commas.
                        InsertString m.RTF.Text, ", " & sDefValue, nPos, sFinal
                        m.RTF.Text = sFinal
                        nPos = nPos + 2 + Len(sDefValue)
        
                        'next argument
                        nInput = nInput + 1
                        GetDefaultValue nInput, sDefValue
                    Loop
        
                    'add closing paren
                    InsertString m.RTF.Text, ") ", nPos, sFinal
                    nPos = nPos + 1
                    m.RTF.Text = sFinal
        
                    m.iInputLineEnd = nPos
                End If
                
                'select the first parm value so user has easy way to overwrite
                m.RTF.SelStart = m.iCaretPos + 1
                m.RTF.SelLength = nSelLen
            End If
            
        End If
        Exit Sub
    Else
        RemoveFunctionList
    End If
    m.iInputLineEnd = m.iInputLineEnd + 1
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.ProcessChg", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub RemoveChar(strBegin As String, charPos As Integer, strFinal As String)
On Error GoTo ErrEnd
    Dim leftStr As String, rightStr As String
    If Len(strBegin) < charPos Then charPos = Len(strBegin)
    leftStr = Left(strBegin, charPos - 1)
    rightStr = Right(strBegin, Len(strBegin) - (charPos))
    strFinal = leftStr & rightStr
Exit Sub
ErrEnd:
    'at least, give 'em back what they sent in
    strFinal = strBegin
End Sub
Private Sub InsertString(strTo As String, strIns As String, insPos As Integer, strFinal)
On Error GoTo ErrEnd
    Dim leftStr As String, rightStr As String
    If Len(strTo) < insPos Then insPos = Len(strTo)
    leftStr = Left(strTo, insPos)
    rightStr = Right(strTo, Len(strTo) - (insPos))
    strFinal = leftStr & strIns & rightStr
Exit Sub
ErrEnd:
    'at least, give 'em back what they sent in
    strFinal = strTo
End Sub

'Search back to the most recent Carriage return, then look for the first non
'blank character going forward.  If it is a quote (') then enter comment mode.
Private Sub SetEditorMode()
On Error GoTo ErrSection:
    
    Dim EnterPos    As Integer
    Dim NonBlankPos As Integer
    Dim SearchPos   As Integer
    
    m.bEditorOnTmp = True
    If m.iCaretPos = 0 Then Exit Sub
    If Len(m.strTemp) = 0 Then Exit Sub
    
    'Search for most recent carriage return
    If m.iCaretPos > Len(m.strTemp) Then SearchPos = Len(m.strTemp) Else SearchPos = m.iCaretPos
    EnterPos = InStrRev(m.strTemp, Chr(10), SearchPos)
    
    'Search forward for nonblank character
    NonBlankPos = m.StrH.FindNonBlank(m.strTemp, EnterPos + 1)
    If NonBlankPos > 0 Then
        If Mid(m.strTemp, NonBlankPos, 1) = "'" Then
            m.bEditorOnTmp = False
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.SetEditorMode", eGDRaiseError_Raise, g.strAppPath

End Sub

'Process non navigation key
Public Sub ProcessKeyDown(pKeyCode As Integer)
On Error GoTo ErrSection:
    
    Dim SelFunction     As String
    
    m.bDisableKey = False
    m.iKeyCode = pKeyCode
    
    'Get current text and hairline focus position
    m.iCaretPos = m.RTF.SelStart
    'persist the position in the text
    m.iPrevCaretPos = m.iCaretPos
    
    m.strTemp = m.RTF.Text

    'Key pressed driver
    Select Case pKeyCode
        Case 8: BackSpace
        Case 9
            If Not frmFunctionList.Visible Then Exit Sub
            TabKey
        Case 13
            If Not frmFunctionList.Visible Then Exit Sub
            EnterKey
        Case 27: Escape
        Case 33: PageUp
        Case 34: PageDown
        Case 35, 36: ProcessMouseClick
        Case 38: UpArrow
        Case 40: DownArrow
        
        'Remove helper lists if user moves outside range of list
        Case Else
            If m.iCaretPos > m.iSelectionListBeg Then
                If frmFunctionList.Visible Then
                    m.FunctionList.Search m.strTemp, m.iCaretPos
                End If
                Exit Sub
            Else
                RemoveFunctionList
            End If
            m.iInputLineEnd = m.iInputLineEnd + 1
    
    End Select
    
    If (m.iKeyCode = 9 Or m.iKeyCode = 13) Then
        'Turn off editing while new text is pasted into RTF box.  This prevents
        'other events from triggering
        m.Editor.TurnOffEditing
        
        'get the default param value
        Dim sDefValue As String, sFinal As String, svCaretPos As Integer
        svCaretPos = m.iCaretPos - 1
        GetDefaultValue m.nCurInput, sDefValue
        Dim nPos As Integer, nSelLen As Integer
        nSelLen = Len(sDefValue)
        'insert the def value into the text stream
        If Len(sDefValue) > 0 Then
            With m.RTF
                InsertString .Text, " " & sDefValue, svCaretPos, sFinal
                .Text = sFinal
            End With
            nPos = svCaretPos + Len(sDefValue) + 1
            
            ' Fill in the rest of the default values
            Dim nInput As Byte
            nInput = m.nCurInput + 1
            GetDefaultValue nInput, sDefValue
            Do While Len(sDefValue) > 0
                'string together parms separated by commas.
                InsertString m.RTF.Text, ", " & sDefValue, nPos, sFinal
                m.RTF.Text = sFinal
                nPos = nPos + 2 + Len(sDefValue)

                'next argument
                nInput = nInput + 1
                GetDefaultValue nInput, sDefValue
            Loop

            'add closing paren
            InsertString m.RTF.Text, ") ", nPos, sFinal
            nPos = nPos + 1
            m.RTF.Text = sFinal

            m.iInputLineEnd = nPos
        End If
        
        'Reset the tabs
        m.Editor.SetTabs m.Editor.TabCnt, m.Editor.TabWidth     'Fixes tab problem in SysNav
        'select the parm value so user has easy way of overwriting
        m.Editor.SelStart = svCaretPos + 1 'm.iCaretPos
        m.Editor.SelLength = nSelLen
        'Turn editing (events) back on
        m.Editor.TurnOnEditing
        
        'there is some form of black magic going on here
        'the input line, which was shown in both TabKey and EnterKey
        'disappears during this previous code...this is the fastest
        'way to force it back on top
        If m.bShowInputsLine Then FormOnTop frmInputsLine.hWnd, True, m.lXCor, m.lYCor
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.ProcessKeyDown", eGDRaiseError_Raise, g.strAppPath

End Sub

'Process navigation key
Public Sub ProcessKeyUp(pKeyCode As Integer)
On Error GoTo ErrSection:

    Dim SelFunction     As String
    
    m.bDisableKey = False
    m.iKeyCode = pKeyCode
    
    'Get current text and hairline focus position
    m.iCaretPos = m.RTF.SelStart
    m.strTemp = m.RTF.Text

    'Key pressed driver
    Select Case pKeyCode
    
        'Action key stokes
        Case 35, 36: ProcessMouseClick
        Case 37: LeftArrow
        Case 39: RightArrow
        
        'Remove helper lists if user moves outside range of list
        Case Else
            If m.iCaretPos > m.iSelectionListBeg Then
                If frmFunctionList.Visible Then
                    m.FunctionList.Search m.strTemp, m.iCaretPos
                End If
                Exit Sub
            Else
                'RemoveFunctionList
            End If
            m.iInputLineEnd = m.iInputLineEnd + 1
    
    End Select
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.cEditingArea.ProcessKeyUp", eGDRaiseError_Raise, g.strAppPath

End Sub

'check and see if the function is found
'if not sent in, try to find closest function to cursor local
Public Function IsValidFunction(ByRef pFunction As String, ByRef bSecurity As Byte) As Boolean
On Error GoTo ErrEnd
    Dim FID As Long, bValid As Boolean
    'this is the ONLY method that sets and clears these 2 vars
    m.sEdFunction = ""
    m.lEdFID = -1
    
    bSecurity = 9 'something really large
    
    If Len(pFunction) = 0 Then GetClosestFunction pFunction
    
    'we tried and found nothing
    If Len(pFunction) = 0 Then Exit Function
    
    bValid = m.Functions.NameFound(CStr(pFunction), FID) And Len(pFunction) > 0
    If bValid Then
        m.sEdFunction = pFunction
        m.lEdFID = FID
        'if its valid - whats the security level
        bSecurity = m.Functions.Item(CStr(FID)).SecurityLevel
    End If
    IsValidFunction = bValid
    
Exit Function
ErrEnd:
End Function

Public Sub ShowSelectedFunctionInfo()
On Error GoTo ErrEnd
    Dim retval As Variant
    Dim cFunctionUsage As New cInputsLine
    Dim sFUsage As String
    Dim sDescription As String
    Dim sExpression As String
    
    'will be working with the frmFunctionInfo
    retval = LockWindowUpdate(frmFunctionInfo.hWnd)
    
    
    'the selected function should be sitting in our module vars
    'with selected function, get the info we need and populate the frmFunctionInfo
    'Assign and build InputsLine
    cFunctionUsage.FunctionRef = m.Functions.Item(CStr(m.lEdFID))
    sFUsage = cFunctionUsage.EditText       '<<Build parmline here...
    sDescription = m.Functions.Item(CStr(m.lEdFID)).Description
    'only if security allows
    If m.Functions.Item(CStr(m.lEdFID)).SecurityLevel < 2 Then
        sExpression = m.Functions.Item(CStr(m.lEdFID)).GetRTF(m.Functions.Item(CStr(m.lEdFID)).Formatted)
    End If
    'populate the form
    Dim RT As New cRichText
 
    With RT
        .RTBox = frmFunctionInfo.txtDesc
        .BuildRTF
        'build our display
        'USAGE
        .AddText "Usage: ", rtfPlain
        .AddText sFUsage, rtfBold
        .AddText vbNewLine
        'DESCRIPTION
        .AddText "Description: ", rtfPlain
        .AddText sDescription, rtfPlain
        .AddText vbNewLine
        'EXPRESSION
        If Len(sExpression) > 0 Then
            .AddText "Expression: ", rtfPlain
        End If
        .BuildRTF
    End With
    If Len(sExpression) > 0 Then
        frmFunctionInfo.txtDesc.SelStart = Len(frmFunctionInfo.txtDesc.Text)
        frmFunctionInfo.txtDesc.SelRTF = sExpression
    End If
    'display it vbModal
    frmFunctionInfo.Show vbModal
    retval = LockWindowUpdate(0)
Exit Sub
ErrEnd:
    retval = LockWindowUpdate(0)
End Sub

Private Sub GetClosestFunction(pFunction As String)
On Error GoTo ErrEnd
    Dim sValidChars As String
    Dim sRight      As String
    Dim sLeft       As String
    Dim sText       As String
    Dim cursorPos   As Integer
    Dim StartPos    As Integer
    Dim endpos      As Integer
    Dim rtfLength   As Integer
    Dim bEndFound   As Boolean
        
        
    'valid chars are U/Lcase letters and spaces
    sValidChars = "abcdefghijklmnopqrstuvwxyz"
    sValidChars = sValidChars & UCase(sValidChars) & " "
    
    sText = m.RTF.Text
    rtfLength = Len(sText)
    cursorPos = m.RTF.SelStart
    endpos = cursorPos
    If endpos = 0 Then endpos = 1
    'build the right side
    While Not endpos > rtfLength And Not bEndFound
        'is the current char valid
        If InStr(1, sValidChars, Mid(sText, endpos, 1)) > 0 Then
            'yep
            sRight = sRight & Mid(sText, endpos, 1)
            endpos = endpos + 1
        ElseIf endpos = 1 Then
            'we are at begin of line and it isnt a valid char
            bEndFound = True
        ElseIf IsNumeric(Mid(sText, endpos, 1)) And Not Mid(sText, endpos - 1, 1) = " " Then
            'based on its positioning (context) - its valid
            sRight = sRight & Mid(sText, endpos, 1)
            endpos = endpos + 1
        Else
            'nope - its not valid
            bEndFound = True
        End If
    Wend
    
    StartPos = cursorPos - 1
    bEndFound = StartPos < 1
    'build the left side
    While Not StartPos = 0 And Not bEndFound
        If InStr(1, sValidChars, Mid(sText, StartPos, 1)) > 0 Then
            'yep
            sLeft = Mid(sText, StartPos, 1) & sLeft
            StartPos = StartPos - 1
        ElseIf StartPos = 1 Then
            'we are at begin of line and it isnt a valid char
            bEndFound = True
        ElseIf StartPos = rtfLength And _
               (IsNumeric(Mid(sText, StartPos, 1)) And Not Mid(sText, StartPos - 1, 1) = " ") Then
            'we are at end of line and but numeric char's positioning (context) is good
            sLeft = Mid(sText, StartPos, 1) & sLeft
            StartPos = StartPos - 1
        ElseIf IsNumeric(Mid(sText, StartPos, 1)) And _
            (Not Mid(sText, StartPos - 1, 1) = " " And Not Mid(sText, StartPos + 1, 1) = " ") Then
            'based on its positioning (context) - its valid
            sLeft = Mid(sText, StartPos, 1) & sLeft
            StartPos = StartPos - 1
        Else
            bEndFound = True
        End If
    Wend
    
    'there is a possibility of operators {and,back,or,if,of}
    'reuse sText; vbTextCompare is not case sensitive
    sText = Trim(sLeft & sRight)
    StripOperators sText
    
    pFunction = Trim(sText)
Exit Sub
ErrEnd:
End Sub

Private Sub StripOperators(sText As String)
On Error GoTo ErrEnd
    Dim bOpNeedCheck As Boolean
    sText = Trim(sText) 'make sure this is done so we don't wack valid names
    bOpNeedCheck = True
    'vbTextCompare is not case sensitive
    While bOpNeedCheck
        If InStr(1, sText, "and ", vbTextCompare) Then
            sText = Right(sText, Len(sText) - 4)
        ElseIf InStr(1, sText, " and", vbTextCompare) Then
            sText = Left(sText, Len(sText) - 4)
        ElseIf InStr(1, sText, "back ", vbTextCompare) Then
            sText = Right(sText, Len(sText) - 5)
        ElseIf InStr(1, sText, " back", vbTextCompare) Then
            sText = Left(sText, Len(sText) - 5)
        ElseIf InStr(1, sText, "or ", vbTextCompare) Then
            sText = Right(sText, Len(sText) - 3)
        ElseIf InStr(1, sText, " or", vbTextCompare) Then
            sText = Left(sText, Len(sText) - 3)
        ElseIf InStr(1, sText, "if ", vbTextCompare) Then
            sText = Right(sText, Len(sText) - 3)
        ElseIf InStr(1, sText, " if", vbTextCompare) Then
            sText = Left(sText, Len(sText) - 3)
        ElseIf InStr(1, sText, "of ", vbTextCompare) Then
            sText = Right(sText, Len(sText) - 3)
        ElseIf InStr(1, sText, " of", vbTextCompare) Then
            sText = Left(sText, Len(sText) - 3)
        ElseIf InStr(1, sText, "not ", vbTextCompare) Then
            sText = Right(sText, Len(sText) - 4)
        ElseIf InStr(1, sText, " not", vbTextCompare) Then
            sText = Left(sText, Len(sText) - 3)
        Else
            bOpNeedCheck = False
        End If
    Wend
Exit Sub
ErrEnd:
End Sub
'Retrieve the function sitting under the focus
Public Sub GetFunction(pFunction As String, pFunctionID As Long, _
        pFound As Boolean, pPreview As String, _
        pFunctionDescription As String, pLastModified As Date, _
        pImplementationTypeID As Byte, pX As Long, pY As Long, _
        pFuncBeg As Long, pFuncLength As Long)
On Error GoTo ErrSection:

    Dim Symbol      As String
    Dim BeginPos    As Integer
    Dim endpos      As Integer
    Dim FunctionID  As Long
    Dim svCaretPos  As Long
    Dim Coloring    As cColoring
    
    m.iCaretPos = m.RTF.SelStart
    
    BeginPos = m.StrH.PrevSepSymbolPos(m.strTemp, m.iCaretPos, Symbol)
    If m.iCaretPos = 0 Then
        endpos = m.StrH.NextSepSymbolPos(m.strTemp, 1, Len(m.strTemp), Symbol)
    Else
        endpos = m.StrH.NextSepSymbolPos(m.strTemp, BeginPos + 1, Len(m.strTemp), Symbol)
    End If
    
    If Len(m.strTemp) = endpos And Symbol = "" Then
        pFunction = m.StrH.curFunction(m.strTemp, endpos)
    Else
        pFunction = m.StrH.curFunction(m.strTemp, endpos - 1)
    End If
            
    'Position focus at beginning of function and get X,Y coordinates
    'for positioning frmFunctionInfo window
    svCaretPos = m.iCaretPos
    m.RTF.SelStart = BeginPos
    'm.RTF.SelStart = 0       '<<Sets at beginning of RTF Box always
    m.lRet = GetCaretPos(PT)
    ClientToScreen m.RTF.hWnd, PT
    pX = PT.X * Screen.TwipsPerPixelX
    pY = (PT.Y * Screen.TwipsPerPixelY) + 250

    'Restore caret...
    m.iCaretPos = svCaretPos
    m.RTF.SetFocus
    m.RTF.SelStart = m.iCaretPos

    pFound = m.Functions.NameFound(CStr(pFunction), pFunctionID) And Len(pFunction) > 0
    If pFound Then
        With m.Functions.Item(CStr(pFunctionID))
            Set Coloring = New cColoring
            pPreview = Coloring.Color(.Preview)
            pFunctionDescription = .Description
            pLastModified = .LastModified
            pImplementationTypeID = .ImplementationTypeID
        End With
    Else
        pPreview = ""
        pFunctionDescription = ""
        pLastModified = 0
        pImplementationTypeID = 0
    End If
    
    'Highlight current function
    If endpos - BeginPos > 0 Then
        m.RTF.SelStart = BeginPos
        m.RTF.SelLength = Len(pFunction) + 1
        m.RTF.SelBold = True
        m.RTF.SelLength = 0
        m.RTF.SelBold = False
        pFuncBeg = BeginPos
        pFuncLength = Len(pFunction)
    End If
    
ErrExit:
    Set Coloring = Nothing
    Exit Sub

ErrSection:
    RaiseError "TSOCX.cEditingArea.GetFunction", eGDRaiseError_Raise, g.strAppPath

End Sub

'Used to offset mCaretPosition AFTER the mouse press event triggers
Property Get CaretOffset()
    CaretOffset = m.iCaretOffset
End Property
Property Let RtfTextBox(pData As Object)
    Set m.RTF = pData
End Property
Property Let EditorCtl(pData As Editor)
    Set m.Editor = pData
End Property
Property Let FunctionsRef(pData As cFunctions)
    Set m.Functions = pData
End Property
Property Let Lists(pData As cLists)
    Set m.Lists = pData
End Property

Public Sub RtfTextBoxSetFocus()
    If GetFocus <> m.RTF.hWnd And Not m.bRefresh Then
        SetFocus (m.RTF.hWnd)
    End If
End Sub

'This must be called before using...
Public Sub Refresh()
On Error GoTo ErrSection:

    m.bRefresh = True
    
    Dim cEditorOptions  As cEditorOptions

    With m.FunctionList
        .FunctionsRef = m.Functions
        .Lists = m.Lists
        .Load
    End With
    
    Set cEditorOptions = New cEditorOptions
    With cEditorOptions
        .Load
        m.bEditorOn = .EditorOn
    End With
    Set cEditorOptions = Nothing
    
    m.bRefresh = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.Refresh", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub MakeListInvisible()
    frmFunctionList.Visible = False
End Sub
Public Sub MakeParmLineInvisible()
    frmInputsLine.Visible = False
End Sub

'Process mouse click in Editing area...
Public Sub ProcessMouseClick()
On Error GoTo ErrSection:
    
    Dim StartPos                As Integer
    Static FunctionHighlighted  As Boolean

    m.iKeyCode = 0
    m.iCaretPos = m.RTF.SelStart
    m.strTemp = m.RTF.Text
    
    'Set the mode of the editor.  Turn off tradesense for comment lines...
    SetEditorMode
    
    If frmInputsLine.Visible Then
    
        'focus moves outside parmline text box boundaries. Remove.
        If m.iCaretPos < m.iInputLineBeg Or _
           m.iCaretPos > m.iInputLineEnd Then
            RemoveInputsLine
            RemoveFunctionList
            Exit Sub
        End If
    
        'focus moves left of Left paren of function
        If m.iCaretPos < m.iInputLineLeft Then
            m.nCurInput = 0
            m.InputsLine.HighlightFunction True
            FunctionHighlighted = True
        Else
            If FunctionHighlighted Then
                m.InputsLine.HighlightFunction False
                FunctionHighlighted = False
            End If
            
            'Go back and find left paren position of InputsLine to display
            StartPos = m.StrH.OpenLeftParenPos(m.strTemp, m.iCaretPos)
            If StartPos > 0 Then
                LeftParen StartPos
            Else
                LeftParen m.iCaretPos
            End If
            SetInput
            
            'm.RTF.SelStart = m.iCaretPos
        End If
    Else
        'if function list if displayed and the caret moves outside the
        'range of the selection then remove i
        If m.iCaretPos < m.iFuncLeftParen Or m.iCaretPos > m.iFuncLeftParen + 25 Then
            RemoveFunctionList
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.ProcessMouseClick", eGDRaiseError_Raise, g.strAppPath

End Sub

'Things to do when the Selection List is made invisible
Private Sub RemoveFunctionList()
    frmFunctionList.Visible = False
    m.iSelectionListBeg = 0
End Sub

Private Sub EnterKey()
On Error GoTo ErrSection:
    
    'Set the mode of the editor.  Turn off tradesense for comment lines...
    SetEditorMode
    
    If frmFunctionList.Visible Then
        
        '3/15/2001 - If at the end of a line, the enter key is most likely
        'used because the user needs to get to the next line to type the
        'THEN keyword.  This was always selecting an item from the list.
        'Now, only tab selects from the list...
        If Not m.bEnterDoesntSelect Then
            ReturnFunction m.strTemp, m.iCaretPos
            ShowInputsLine m.iCaretPos
            m.bDisableKey = True
        End If
        RemoveFunctionList
    Else
        RemoveInputsLine
    End If
    
    'Adjust focus past left parenthesis
    If frmInputsLine.Visible Then
        m.iCaretPos = m.iCaretPos + 1
        m.RTF.SelStart = m.iCaretPos
        m.RTF.SetFocus
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.EnterKey", eGDRaiseError_Raise, g.strAppPath

End Sub

'Space bar:  If a list is displayed already, search it based on characters
'typed.  Otherwise, determine if a list should be displayed.
Private Sub SpaceBar()
On Error GoTo ErrSection:

    Dim ShowFunctionList    As Boolean
    Dim xWord               As String
    Dim Fnd                 As Boolean
    
    If frmFunctionList.Visible Then
        m.FunctionList.Search m.strTemp, m.iCaretPos
        Exit Sub
    End If
    
    'Set flag to enable display of function list...
    ShowFunctionList = True
    m.FunctionList.ListID = "Functions"
    
    'Verify that a separator word preceeded the space bar...
    Fnd = False
    xWord = m.StrH.curWord(m.strTemp, m.iCaretPos - 1)
    
    'tjr 1/03 - added check for no entry to access tradesense on entry
    'tjr 2/03 - removed check as it was fouling up when to display tradesense
    '           ref: issue #201
'    If Len(xWord) > 0 Then
        If m.StrH.SeparatorWordFound(xWord) Or m.iCaretPos <= 1 Or xWord = "" Then
            'If the operator is one of the following then don't show tradesense
            If (xWord <> ")" And xWord <> """" And _
               xWord <> "Endif" And xWord <> "then" And xWord <> "Else") Then
                Fnd = True
            End If
        End If
        
        If Fnd Then
            
            If xWord = "Of" Then
                m.FunctionList.ListID = "Of"     'Symbols list
            End If
        
            'Parmline is visible, if comma was the separator check the parm
            If frmInputsLine.Visible And (xWord = "," Or xWord = "(") Then
            
                'Determine if parmlist should be displayed...
                If m.nCurInput = 0 Then
                    m.nCurInput = m.Func.HiddenInputs + 1
                End If
                
            Else
                RemoveInputsLine
            End If
            m.iSelectionListBeg = m.iCaretPos - 1
        Else
            ShowFunctionList = False
        End If
'    End If
    
    If Not m.bEditorOn Then Exit Sub
    If Not m.bEditorOnTmp Then Exit Sub
    
    'Convert the caret position from screen coordinates to twips
    If ShowFunctionList Then
        'RichTextBox1.SetFocus
        m.lRet = GetCaretPos(PT)
        ClientToScreen m.RTF.hWnd, PT
        m.lXCor = PT.X * Screen.TwipsPerPixelX
        m.lYCor = (PT.Y * Screen.TwipsPerPixelY) + 300
        'tjr 1/03 - added code to best guess category list user would
        '           want to be viewing at this time
        If xWord = "." Then
            'default to C_ALL
            frmFunctionList!cbCategory = C_BARSAGO
        Else
            'default to C_ALL
            frmFunctionList!cbCategory = C_ALL
        End If
        m.FunctionList.ShowNewFunction = m.Editor.ShowNewFunction
        m.FunctionList.Show m.lXCor, m.lYCor
    End If
  
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.SpaceBar", eGDRaiseError_Raise, g.strAppPath

End Sub

'Double Quote:  Start a text string.  If a found followed by a comparison
'symbol precedes the double quote then show a list.
Private Sub DoubleQuote()
On Error GoTo ErrSection:

    Dim Operator            As String
    Dim FunctionName        As String
    Dim FunctionID          As Long
    Dim SearchPos           As Integer
    
    'Operator preceding double quote is a comparison operator...
    Operator = m.StrH.curWord(m.strTemp, m.iCaretPos - 1)
    If Operator = "=" Or Operator = ">" Or Operator = ">=" Or Operator = "<" Or _
       Operator = "<=" Or Operator = "<>" Then
        
        'Search for Function preceding operator...
        SearchPos = m.StrH.FindPrevChar(m.strTemp, m.iCaretPos - 1, Operator)
        FunctionName = m.StrH.curFunction(m.strTemp, SearchPos - 1)
        If m.Functions.NameFound(CStr(FunctionName), FunctionID) And Len(FunctionName) > 0 Then
            
            'If a listID is found for the function then show the list
            If m.Functions.Item(CStr(FunctionID)).ListID > 0 Then
                m.FunctionList.ListID = m.Functions.Item(CStr(FunctionID)).ListID
                m.iSelectionListBeg = m.iCaretPos - 1
                
                'Show TradeSense list
                m.lRet = GetCaretPos(PT)
                ClientToScreen m.RTF.hWnd, PT
                m.lXCor = PT.X * Screen.TwipsPerPixelX
                m.lYCor = (PT.Y * Screen.TwipsPerPixelY) + 300
                m.FunctionList.Show m.lXCor, m.lYCor
            End If
        End If
    Else
        'Double quote entered after a comma.  If the inputs line is shown,
        'then this is an input.  If it is a text type and a list ID is defined
        'for the input show the list.
        If Operator = "," Or Operator = "(" Then
            If frmInputsLine.Visible Then
                If m.nCurInput = 0 Then
                    m.nCurInput = m.Func.HiddenInputs + 1
                End If
                If m.Func.Inputs.Item(m.nCurInput).ListID > 0 Then
                    m.FunctionList.ListID = m.Func.Inputs.Item(m.nCurInput).ListID
                
                    'Show TradeSense list
                    m.lRet = GetCaretPos(PT)
                    ClientToScreen m.RTF.hWnd, PT
                    m.lXCor = PT.X * Screen.TwipsPerPixelX
                    m.lYCor = (PT.Y * Screen.TwipsPerPixelY) + 300
                    m.FunctionList.Show m.lXCor, m.lYCor
                End If
            Else
                RemoveInputsLine
            End If
            m.iSelectionListBeg = m.iCaretPos - 1
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.cEditingArea.DoubleQuote", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub MathSymbols()
On Error GoTo ErrSection:

    RemoveFunctionList
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.MathSymbols", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub CompareSymbols()
On Error GoTo ErrSection:

    RemoveFunctionList
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.CompareSymbols", eGDRaiseError_Raise, g.strAppPath

End Sub

'When a comma is typed, attempt to show the parameter list if it is not already displayed,
'Otherwise, move between parameters in the currently displayed parmline...
Private Sub Comma()
On Error GoTo ErrSection:

    Dim Continue    As Boolean
    Dim StartPos    As Integer
    
    If Not m.bEditorOn Then Exit Sub
    If Not m.bEditorOnTmp Then Exit Sub
    
    'Steve: commented this out so we always determine current function
    'If parm line is already displayed move to next parm
'    If frmInputsLine.Visible Then
'        SetInput
'    Else
        'Parmline is NOT display.  Determine what parm line to display using the LeftParm
        'function (first find the left paren following the function)
        StartPos = m.StrH.OpenLeftParenPos(m.strTemp, m.iCaretPos)
        Set m.Func = Nothing
        If StartPos > 0 Then
            LeftParen StartPos
'            SetInput
        End If
'    End If
    
    
    ' Do we have a function?
    If Not m.Func Is Nothing Then
        ' Get the parms in the function
        Dim nInputs As Byte
        nInputs = m.Func.HiddenInputs + m.StrH.ParmsInFxn(m.strTemp, m.iCaretPos)
        
        ' Have all the arguments been specified?
        If nInputs > m.Func.Inputs.Count Then
            ' Remove the comma
            RemoveChar m.strTemp, m.iCaretPos, m.strTemp
            m.RTF.Text = m.strTemp
            m.iCaretPos = m.iCaretPos - 1
            If m.iCaretPos < 0 Then m.iCaretPos = 0
            m.RTF.SelStart = m.iCaretPos

            ' Ignore the keystroke
            m.iKeyCode = 0
            
            ' Highlight the next input
            Dim nPos As Integer, nParens As Integer, nCommas As Integer
            nParens = 1
            nCommas = 0
            For nPos = m.iCaretPos + 1 To Len(m.RTF.Text)
                Dim sChar As String
                sChar = Mid(m.RTF.Text, nPos, 1)
                If sChar = "(" Then
                    nParens = nParens + 1
                ElseIf sChar = ")" Then
                    nParens = nParens - 1
                    If nParens = 0 Then
                        If nCommas = 1 Then
                            m.RTF.SelStart = m.iCaretPos
                            m.RTF.SelLength = Len(Trim(Mid(m.RTF.Text, m.RTF.SelStart, nPos - m.iCaretPos)))
                            Exit For
                        Else
                            m.RTF.SelStart = nPos + 1
                        End If
                        Exit For
                    End If
                ElseIf sChar = "," And nParens = 1 Then
                    nCommas = nCommas + 1
                    If nCommas = 1 Then
                        m.iCaretPos = nPos + 1
                    Else
                        m.RTF.SelStart = m.iCaretPos
                        m.RTF.SelLength = Len(Trim(Mid(m.RTF.Text, m.RTF.SelStart, nPos - m.iCaretPos)))
                        Exit For
                    End If
                End If
            Next
            
            ' Reset the input
            SetInput
            
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.Comma", eGDRaiseError_Raise, g.strAppPath

End Sub

'Search back starting at m.iCaretPos for Open Left Paren, if found
'execute left paren logic to show a New cInputsLine
Private Sub RightParen(pCaretPos As Integer)
On Error GoTo ErrSection:

    Dim StartPos    As Integer
    
    RemoveFunctionList
    RemoveInputsLine
    
    'Go back and find left paren position of m.InputsLine just removed.
    'Then, determine if this was nested by searching back again for
    'another function left paren.  If found, show that parmline.
    StartPos = m.StrH.OpenLeftParenPos(m.strTemp, pCaretPos - 1)
    If StartPos > 0 Then
        StartPos = m.StrH.OpenLeftParenPos(m.strTemp, StartPos - 1)
        If StartPos > 0 Then
            LeftParen StartPos
        End If
    End If
    m.RTF.SelStart = pCaretPos
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.RightParen", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub Escape()
On Error GoTo ErrSection:

    If frmFunctionList.Visible Then
        RemoveFunctionList
        Exit Sub
    Else
        RemoveInputsLine
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.Escape", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub BackSpace()
On Error GoTo ErrSection:

    If frmFunctionList.Visible Then
        If m.iCaretPos > m.iSelectionListBeg Then
            m.FunctionList.Search m.strTemp, m.iCaretPos
            Exit Sub
        Else
            RemoveFunctionList
        End If
    End If
    
    If frmInputsLine.Visible Then
        If m.iCaretPos - 1 < m.iInputLineLeft Then
            RemoveInputsLine
        Else
            'Set previous parm if necessary
            SetInput
        End If
    End If
    
    'If backspace over a double quote then toggle the mode
    'If Asc(Mid(m.strTemp, m.RTF.SelStart, 1)) = 34 Then
    '    DoubleQuote
    'End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.BackSpace", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub LeftParen(pStart As Integer)
On Error GoTo ErrSection:

    With frmFunctionList
        If .Visible Then .Visible = False
    End With
    ShowInputsLine pStart
    SetInput
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.LeftParen", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub FunctionHasParams(pStart As Integer, bValidFunction As Boolean, iParams As Integer, Optional iHiddenParams As Integer)
'returns wether this is a valid function and the amount of parameters it will have
On Error GoTo ErrEnd
    Dim curFunctionName As String
    Dim FunctionID      As Long
    Dim svCaretPos      As Integer

    curFunctionName = m.StrH.curFunction(m.strTemp, pStart - 1)
    bValidFunction = m.Functions.NameFound(CStr(curFunctionName), FunctionID) And Len(curFunctionName) > 0
    If bValidFunction Then
        With m.Functions.Item(CStr(FunctionID))
            If .Inputs.Count > .HiddenInputs Then
                iParams = .Inputs.Count - .HiddenInputs
                iHiddenParams = .HiddenInputs
            Else
                iParams = 0
                iHiddenParams = 0
            End If
        End With
    End If
Exit Sub
ErrEnd:
End Sub
Private Sub ShowInputsLine(pStart As Integer)
On Error GoTo ErrSection:

    Dim curFunctionName As String
    Dim FunctionID      As Long
    Dim svCaretPos      As Integer
    
    'tjr 2/03 - this is not the best place for this, however, all code
    'currently works to check the function for input line - here.  Sometimes
    'the caller sends a '(' sometimes they do not.  Lets just figure it out and
    'send appropriatly
    Dim cur As Integer
    Const validInputLineChars = "abcdefghijklmnopqrstuvwxyz"
    
    cur = pStart
    While cur > 1 And Not InStr(1, validInputLineChars, Mid(m.strTemp, cur, 1), vbTextCompare) > 0
        cur = cur - 1
    Wend
    curFunctionName = m.StrH.curFunction(m.strTemp, cur)
    m.bShowInputsLine = True 'default it shown
    If m.Functions.NameFound(CStr(curFunctionName), FunctionID) And Len(curFunctionName) > 0 Then
    
        'Dont Set the parmline variables unless inputs exists to show.  Other
        '-wise, m.Func can be pointed to a function with no inputs.  For
        'example "Average (Account Balance,"  When "," is pressed, mFunciton
        'should still point to Average.
'        With m.Functions.Item(CStr(FunctionID))
'            If .Inputs.Count <= .HiddenInputs Then Exit Sub
'        End With
    
        'Set a reference to the function just typed.  Finds its starting and
        'ending position (in the Rtf box)
        Set m.Func = m.Functions.Item(CStr(FunctionID))
        
            
        m.iFuncBegPos = m.StrH.FindFunctionBegPos(m.strTemp, pStart - 1)
        m.iFuncLeftParen = pStart
        
        'Assign and build m.InputsLine
        With m.InputsLine
            .FunctionRef = m.Func
            m.strFunctionName = .EditText       '<<Build parmline here...
            .FunctionBegin = m.iFuncBegPos
            .FunctionLeftParen = m.iFuncLeftParen
            If m.Func.Inputs.Count > 0 Then
                m.nCurInput = 0
                'If all inputs are hidden inputs then quit...
'                If m.Func.Inputs.Count = m.Func.HiddenInputs Then Exit Sub
                m.nCurInput = m.Func.HiddenInputs + 1
            Else
                m.nCurInput = 0
'                Exit Sub
            End If
        End With
        
        'Save the beginning of the first function of the first m.InputsLine.  This
        'is used to control highlighting of nested functions
        If m.iInputLineBeg = 0 Then
            m.iInputLineLeft = m.StrH.OpenLeftParenPos(m.strTemp, m.iFuncLeftParen - 1)
            If m.iInputLineLeft = 0 Then
                m.iInputLineLeft = m.iFuncLeftParen
                m.iInputLineBeg = m.iFuncBegPos
            Else
                'Nest function...
                m.iInputLineBeg = m.StrH.FindFunctionBegPos(m.strTemp, m.iInputLineLeft - 1)
            End If
            
            'pStart points to the "(" character.  Point one past this
            'to start the search...
            m.iInputLineEnd = m.StrH.CalInputsLineEnd(m.strTemp, pStart + 1)
        End If
        
        'Position focus at beginning of function and get X,Y
        svCaretPos = m.iCaretPos
        m.RTF.SelStart = m.iFuncBegPos
        m.lRet = GetCaretPos(PT)
        If m.bEditorOn And m.bEditorOnTmp Then
            ClientToScreen m.RTF.hWnd, PT
            m.lXCor = PT.X * Screen.TwipsPerPixelX
            m.lYCor = (PT.Y * Screen.TwipsPerPixelY) + 250     'NOTE: 300 should really be the height of the text in case the font is changed
            
            'Show m.InputsLine...
            m.InputsLine.Show m.lXCor, m.lYCor
        End If
    
        'Restore caret...
        m.iCaretPos = svCaretPos
        m.RTF.SetFocus
        m.RTF.SelStart = m.iCaretPos
    Else
        'nothing to show
        m.bShowInputsLine = False
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.ShowInputsLine", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub GetDefaultValue(nParam As Byte, sDefValue As String)
On Error GoTo ErrEnd
    Dim Cnt As Integer, X As Integer
    sDefValue = ""
    
    With m.Func.Inputs
    
        'If no parms exist for function strip out left paren
        Cnt = .Count
        If Cnt = 0 Then Exit Sub
            
        sDefValue = CStr(.Item(nParam).DefaultValue)
        If .Item(nParam).ParmTypeID = 6 Then
            If .Item(nParam).DefaultValue = 0 Then
                sDefValue = "False"
            Else
                sDefValue = "True"
            End If
        End If
    End With
    
Exit Sub
ErrEnd:
    'no biggie, just don't get a default value
End Sub

'This routine's job is to control the parameter position within the parmline
'and display the list box for the current parm if necessary...
Private Sub SetInput()
On Error GoTo ErrSection:
    
    Dim NextParm        As Integer
    Dim ParmCnt         As Integer
    Dim X               As Byte
    
    RemoveFunctionList
    
    If m.Func Is Nothing Then Exit Sub
    
    'Find position of parameter in parmline based on m.iCaretPos...
    NextParm = m.StrH.ParmToHighlight(m.strTemp, m.iFuncBegPos, m.iCaretPos)
    
    'If m.Func.HiddenInputs = 1 Then
    '    NextParm = NextParm + 1
    'End If
    NextParm = NextParm + m.Func.HiddenInputs
    
    'If nextparm exceeds the parms for the function then leave
    ParmCnt = m.Func.Inputs.Count
    If NextParm <= ParmCnt Then
        
        'Removed 12/27/2000
        'Unhighlight current parm...
        'When a comma is typed in the parmline area of an existing
        'function with parms, the parmline is displayed, but the
        'Prev parm is zero
        'For X = 1 To ParmCnt
        '    m.InputsLine.HighlightParm X, False
        'Next X
        With frmFunctionList
            If .Visible Then .Visible = False
        End With
        
        'Assign new parm position...
        m.nCurInput = NextParm
        m.FunctionList.ListID = m.Func.Inputs.Item(m.nCurInput).ListID
        
        'Determine if List should be displayed...(if not backspace key and mouseclick)
        If m.FunctionList.ListID <> 0 And m.iKeyCode <> 8 And m.iKeyCode <> 0 Then
                
            'Show list here...
            m.iSelectionListBeg = m.iCaretPos
            m.lRet = GetCaretPos(PT)
            If m.bEditorOn And m.bEditorOnTmp Then
                ClientToScreen m.RTF.hWnd, PT
                m.lXCor = PT.X * Screen.TwipsPerPixelX
                m.lYCor = (PT.Y * Screen.TwipsPerPixelY) + 550
                m.FunctionList.Show m.lXCor, m.lYCor
                m.RTF.SelStart = m.iCaretPos
            End If
        End If
        
        'Highlight current parm...
        m.InputsLine.HighlightParm m.nCurInput, True
        
        '???? 12/27/2000 Why?
        'm.iInputLineEnd = m.iInputLineEnd + 1
    Else
        m.nCurInput = NextParm
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.SetInput", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub TabKey()
On Error GoTo ErrSection:

    Dim FLen        As Integer
    Dim LParenPos   As Integer
    
    If frmFunctionList.Visible Then
        ReturnFunction m.strTemp, m.iCaretPos
        ShowInputsLine m.iCaretPos
        m.bDisableKey = True
    End If
    
    If frmInputsLine.Visible Then
        m.iInputLineEnd = m.StrH.CalInputsLineEnd(m.strTemp, m.iInputLineLeft)
        
        ' Adjust focus past left parenthesis
        LParenPos = InStr(m.iCaretPos, m.strTemp, "(")
        If LParenPos > 0 Then
            m.iCaretPos = m.iCaretPos + 1
        End If
        m.RTF.SelStart = m.iCaretPos
        m.RTF.SetFocus
        
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.TabKey", eGDRaiseError_Raise, g.strAppPath

End Sub

'Returns the currently selected row from the functions list box and
'displays it in the Rtf box at the current caret position.
Private Sub ReturnFunction(pStr As String, pCaretPos As Integer)
On Error GoTo ErrSection:
'note: pCaretPos is after inserted text

    Dim FunctionEnd             As Long
    Dim OperSymbol              As String
    Dim RightChars              As String
    Dim SelFunction             As String
    Dim FunctionID              As Long
    Dim ParmsFound              As Boolean
    Dim OpPrev                  As Long
    Dim OpNext                  As Long
    Dim TempStr                 As String
    Dim Parens                  As Long
    Dim CurChar                 As String
    Dim X                       As Long
    Dim svCaretPos              As Integer
    Dim rSide                   As String
    
    rSide = Right(pStr, Len(pStr) - pCaretPos)
    'Get selection
    SelFunction = m.FunctionList.Selected
    RemoveFunctionList
    
    'Determine number of enterable parms
    ParmsFound = False
    If m.Functions.NameFound(CStr(SelFunction), FunctionID) And Len(SelFunction) > 0 Then
        If m.Functions.Item(CStr(FunctionID)).Inputs.Count - _
           m.Functions.Item(CStr(FunctionID)).HiddenInputs > 0 Then
            ParmsFound = True
        End If
    End If
    
    'Determine spot to return the function.  If the operator left of
    'the spacebar is "(" or double quote, return it right next to it,
    'otherwise insert one space (after IF, AND, etc.)
    OpPrev = m.StrH.PrevSepSymbolPos2(pStr, pCaretPos, OperSymbol, SelFunction)
    If OperSymbol = "" Then
        TempStr = SelFunction
    Else
        If OperSymbol = "(" Or OperSymbol = """" Then
            TempStr = Left(pStr, OpPrev) & SelFunction
        Else
            'make sure there is a space
            TempStr = Left(pStr, OpPrev) & " " & SelFunction  'include the space
        End If
    End If
    svCaretPos = Len(TempStr)
    'based on need for parameters, add the paren
    If ParmsFound Then
        TempStr = TempStr & " (" & rSide
        svCaretPos = svCaretPos + 2
    Else
        TempStr = TempStr & rSide
    End If

   
    'Turn off editing while new text is pasted into RTF box.  This prevents
    'other events from triggering
    m.Editor.TurnOffEditing
    
    With m.RTF
        .Text = TempStr
        .SetFocus
        .SelStart = svCaretPos
        pCaretPos = svCaretPos
        pStr = TempStr
    End With
    m.Editor.SetTabs m.Editor.TabCnt, m.Editor.TabWidth     'Fixes tab problem in SysNav
    
    'Reset the tabs
    
    'Turn editing (events) back on
    m.Editor.TurnOnEditing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.ReturnFunction", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub PageUp()
On Error GoTo ErrSection:

    If frmFunctionList.Visible Then
        m.FunctionList.PageUp False
        m.bDisableKey = True
    Else
        If frmInputsLine.Visible Then
            RemoveInputsLine
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.PageUp", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub PageDown()
On Error GoTo ErrSection:

    If frmFunctionList.Visible Then
        m.FunctionList.PageDown False
        m.bDisableKey = True
    Else
        If frmInputsLine.Visible Then
            RemoveInputsLine
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.PageDown", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub UpArrow()
On Error GoTo ErrSection:

    If frmFunctionList.Visible Then
        m.FunctionList.SearchUp False
        m.bDisableKey = True
    Else
        If frmInputsLine.Visible Then
            RemoveInputsLine
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.UpArrow", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub DownArrow()
On Error GoTo ErrSection:

    If frmFunctionList.Visible Then
        m.FunctionList.SearchDown False
        m.bDisableKey = True
    Else
        If frmInputsLine.Visible Then
            RemoveInputsLine
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.DownArrow", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub LeftArrow()
On Error GoTo ErrSection:
    
    Dim CurChar     As String
        
    If frmFunctionList.Visible Then
        If m.iCaretPos < m.iSelectionListBeg Then
            RemoveFunctionList
        End If
    End If
    
    If frmInputsLine.Visible Then
        CurChar = Mid(m.strTemp, m.iCaretPos + 1, 1)
        Select Case CurChar
            Case "("
                LeftArrowLeftParen
            Case ")"
                LeftArrowRightParen
            Case ","
                LeftArrowComma
            Case Else
                LeftArrowOther
        End Select
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.cEditingArea.LeftArrow", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub LeftArrowLeftParen()
On Error GoTo ErrSection:
    
    Dim StartPos                As Integer
    Static FunctionHighlighted  As Boolean
    
    'Unhighlight the currently selected parm (first parm)
    'Highlight function name
    If m.iCaretPos - 1 < m.iInputLineLeft Then
        m.nCurInput = 0
        m.InputsLine.HighlightFunction True
        FunctionHighlighted = True
    Else
        If FunctionHighlighted Then
            m.InputsLine.HighlightFunction False
            FunctionHighlighted = False
        End If
        
        'Go back and find left paren position of m.InputsLine to display
        StartPos = m.StrH.OpenLeftParenPos(m.strTemp, m.iCaretPos - 1)
        If StartPos > 0 Then
            LeftParen StartPos
        Else
            LeftParen m.iCaretPos
        End If
        m.RTF.SelStart = m.iCaretPos
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.cEditingArea.LeftArrowLeftParen", eGDRaiseError_Raise, g.strAppPath

End Sub

'caret moved to one position left of right parenthesis
Private Sub LeftArrowRightParen()
On Error GoTo ErrSection:
    
    Dim StartPos    As Integer

    'Go back and find left paren position of m.InputsLine to display
    StartPos = m.StrH.FindPrevChar(m.strTemp, m.iCaretPos - 1, "(")
    If StartPos > 0 Then
        LeftParen StartPos
        SetInput
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.LeftArrowRightParen", eGDRaiseError_Raise, g.strAppPath

End Sub

'Determine the parm to highlight based on caret position
Private Sub LeftArrowComma()
On Error GoTo ErrSection:

    SetInput
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.LeftArrowComma", eGDRaiseError_Raise, g.strAppPath

End Sub

'Determine the parm to highlight based on caret position
Private Sub LeftArrowOther()
On Error GoTo ErrSection:

    If m.iCaretPos < m.iInputLineBeg Then
        RemoveInputsLine
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.LeftArrowOther", eGDRaiseError_Raise, g.strAppPath

End Sub

'Action to take if right arrow pressed with the m.InputsLine displayed.  Take
'action if the current character is left paren, right paren, or comma
Private Sub RightArrow()
On Error GoTo ErrSection:
    
    Dim CurChar     As String
    
    If frmInputsLine.Visible Then
        CurChar = Mid(m.strTemp, m.iCaretPos, 1)
        Select Case CurChar
            Case "("
                LeftParen m.iCaretPos
            Case ")"
                RightArrowRightParen
            Case ","
                RightArrowComma
            Case Else
                RightArrowOther
        End Select
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.RightArrow", eGDRaiseError_Raise, g.strAppPath

End Sub

'Right arrow pressed to the right of right paren
Private Sub RightArrowRightParen()
On Error GoTo ErrSection:
    
    Dim StartPos    As Integer

    If m.iCaretPos > m.iInputLineEnd Or m.iCaretPos = Len(m.strTemp) Then
        RemoveInputsLine
    Else
        'Go back and find left paren position of m.InputsLine to display
        StartPos = m.StrH.OpenLeftParenPos(m.strTemp, m.iCaretPos - 1)
        If StartPos > 0 Then
            LeftParen StartPos
            SetInput
        End If
        m.RTF.SelStart = m.iCaretPos
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.RightArrowRightParen", eGDRaiseError_Raise, g.strAppPath

End Sub

'Determine the parm to highlight based on caret position
Private Sub RightArrowComma()
On Error GoTo ErrSection:
    
    Dim StartPos    As Integer
    
    'Go back and find left paren position of InputsLine to display
    StartPos = m.StrH.OpenLeftParenPos(m.strTemp, m.iCaretPos)
    If StartPos > 0 Then
        LeftParen StartPos
    Else
        LeftParen m.iCaretPos
    End If
    SetInput
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.RightArrowComma", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub RightArrowOther()
On Error GoTo ErrSection:

    If m.iCaretPos > m.iInputLineEnd Then
        RemoveInputsLine
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.RightArrowOther", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub RemoveInputsLine()
On Error GoTo ErrSection:

    frmInputsLine.Visible = False
    Unload frmFunctionInfo
    m.nCurInput = 0
    m.iInputLineBeg = 0
    m.iInputLineEnd = 0
    m.iInputLineLeft = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.RemoveInputsLine", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub NewFunction(ByVal lCategoryID As Long)
On Error GoTo ErrSection:

    m.Editor.NewFunction lCategoryID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditingArea.NewFunction", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Sub Class_Initialize()
    
    Set m.StrH = New cStrh
    Set m.FunctionList = New cFunctionList
    Set m.InputsLine = New cInputsLine
    m.bRefresh = False

End Sub

Private Sub Class_Terminate()
    
    Set m.FunctionList = Nothing
    Set m.InputsLine = Nothing
    Set m.StrH = Nothing
    Set m.Functions = Nothing
    Set m.Func = Nothing
    Set m.Lists = Nothing
    Set m.RTF = Nothing
    Set m.Editor = Nothing

End Sub
