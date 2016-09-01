Attribute VB_Name = "mCommon"
Option Explicit
Option Compare Text

Type gGlobal
    strAppPath As String                ' Calling application path
    dbNav As Database                   ' Navigator database
    EditorOptions As cEditorOptions     ' Editor options
    
    lLCD As Long                        ' Last good customer ID
End Type
Global g As gGlobal

'Implementation types
Global Const C_Builtin = 1
Global Const C_Custom = 2

'Source types
Global Const C_MM = 1
Global Const C_System = 2
Global Const C_Both = 3

'Parm Return types
Global Const C_RetNumericConstant = 1
Global Const C_RetText = 2
Global Const C_RetTrueFalse = 3
Global Const C_RetNumeric = 4
Global Const C_RetBars = 5
Global Const C_RetTrueFalseConstant = 6
Global Const C_RetTrades = 7

Global Const gTokenLen = 6

'Data types identifying valid phrases in a rule
Global Const C_FUNC_NUMERIC = 1
Global Const C_FUNC_BOOLEAN = 2
Global Const C_FUNC_BOOLEAN_CONSTANT = 3
Global Const C_FUNC_NUMERIC_CONSTANT = 4
Global Const C_PARM_NUMERIC = 5         'Always a constant numeric
Global Const C_PARM_BOOLEAN = 6         'Always a constant boolean
Global Const C_PARM_BARS = 7
Global Const C_ADD = 8
Global Const C_OR = 9
Global Const C_MULTI = 10
Global Const C_COMPARE = 11
Global Const C_AND = 12
Global Const C_NUMERIC = 13
Global Const C_LEFTPAREN = 14
Global Const C_RIGHTPAREN = 15
Global Const C_FLEFTPAREN = 16
Global Const C_FRIGHTPAREN = 17
Global Const C_ERROR = 18
Global Const C_OFFSET = 19
Global Const C_TEXT = 20
Global Const C_OF = 21
Global Const C_COMMA = 22
Global Const C_ENTER = 23
Global Const C_IF = 24
Global Const C_PARM_TRADES = 25
Global Const C_NOT = 26
Global Const C_PARM_NUMERIC_ARRAY = 27
Global Const C_PARM_BOOLEAN_ARRAY = 28
Global Const C_PARAGRAPH = 80
'These are used for controling rule display in frmMM
'Global Const C_TITLE = 50
'Global Const C_RULENAME = 51
'Global Const C_ACTIONLABEL = 52

'New Tokens required to be compatible with new TradeSense
Global Const C_DoubleQuote = 32
Global Const C_Then = 35
Global Const C_COMMENT = 36
Global Const C_Else = 41
Global Const C_ElseIf = 42
Global Const C_EndIf = 43
Global Const C_Tab = 44
'tjr 2/03 - sorry if this hoses mark, didn't know we
'were using 45 and had bill view it as a comment block in engine
'...upon further review, comments should never be found
'...in coded text, so we should be fine with 45 as a 'non-comment'
Global Const C_EnterFormatting = 45
Global Const C_BRACKETCOMMENT = 81

'Ini file access
Const V_EMPTY = 0    ' Empty
Const V_NULL = 1     ' Null
Const V_INTEGER = 2  ' Integer
Const V_LONG = 3     ' Long
Const V_SINGLE = 4   ' Single
Const V_DOUBLE = 5   ' Double
Const V_CURRENCY = 6 ' Currency
Const V_DATE = 7     ' Date
Const V_STRING = 8   ' String

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FirstCharValid
'' Description: First character must be alphabetic or an underscore
'' Inputs:      String to check, Error Message
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FirstCharValid(ByVal strData As String, strErrMsg As String) As Boolean
On Error GoTo ErrSection:

    FirstCharValid = False
    If (Left(strData, 1) >= "a" And Left(strData, 1) <= "z") Or Left(strData, 1) = "_" Then
        FirstCharValid = True
    Else
        strErrMsg = "First character of : " & UCase(strData) & " must be alphabetic"
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mCommon.FirstCharValid", eGDRaiseError_Raise, g.strAppPath
    
End Function
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemainCharsValid
'' Description: Make sure the remaining characters are valid
'' Inputs:      String to check, Error Message
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemainCharsValid(ByVal strData As String, strErrMsg As String) As Boolean
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim strCurChar As String            ' Current character in the string
    
    RemainCharsValid = True
       
    ' Check the beginning characters
    strCurChar = Mid(strData, 1, 3)
    Select Case strCurChar
        Case "If ", "Or "
            RemainCharsValid = False
            strErrMsg = UCase(strData) & " cannot start with " & UCase(strCurChar) & "."
            Exit Function
    End Select
        
    strCurChar = Mid(strData, 1, 4)
    Select Case strCurChar
        Case "and ", "not "
            RemainCharsValid = False
            strErrMsg = UCase(strData) & " cannot start with " & UCase(strCurChar) & "."
            Exit Function
    End Select
        
    strCurChar = Mid(strData, 1, 5)
    Select Case strCurChar
        Case "back "
            RemainCharsValid = False
            strErrMsg = UCase(strData) & " cannot start with " & UCase(strCurChar) & "."
            Exit Function
    End Select
    
    For lIndex = 2 To Len(strData)
        strCurChar = Mid(strData, lIndex, 1)
        Select Case strCurChar
            Case "+", "-", "/", "*", ".", "%", Chr(34), _
                Chr(39), ">", "<", "=", "(", ")"
                RemainCharsValid = False
                strErrMsg = "Name: " & UCase(strData) & _
                    " cannot contain " & UCase(strCurChar)
                Exit Function
        End Select
        
        strCurChar = Mid(strData, lIndex, 2)
        Select Case strCurChar
            Case ">=", "<=", "<>"
                RemainCharsValid = False
                strErrMsg = "Name: " & UCase(strData) & _
                    " cannot contain " & UCase(strCurChar)
                Exit Function
        End Select
        
        strCurChar = Mid(strData, lIndex, 4)
        Select Case strCurChar
            Case " or ", " of "
                RemainCharsValid = False
                strErrMsg = UCase(strData) & " cannot contain " & UCase(strCurChar)
                Exit Function
        End Select
        
        strCurChar = Mid(strData, lIndex, 5)
        Select Case strCurChar
            Case " and ", " not "
                RemainCharsValid = False
                strErrMsg = UCase(strData) & " cannot contain " & UCase(strCurChar)
                Exit Function
        End Select
        
        strCurChar = Mid(strData, lIndex, 6)
        Select Case strCurChar
            Case " back "
                RemainCharsValid = False
                strErrMsg = UCase(strData) & " cannot contain " & UCase(strCurChar)
                Exit Function
        End Select
    Next lIndex
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mCommon.RemainCharsValid", eGDRaiseError_Raise, g.strAppPath
    
End Function
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OperatorsFound
'' Description: Checks a string for operators
'' Inputs:      String to check, Error Message
'' Returns:     True if operators found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OperatorsFound(ByVal strData As String, strErrMsg As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bErrorFound As Boolean          ' Was an error found?
    Dim lLenChar As Long                ' Length of a character string
    
    OperatorsFound = False
    bErrorFound = False
    For lIndex = 1 To Len(strData)
        If Mid(strData, lIndex, 4) = " And" Or _
           Mid(strData, lIndex, 5) = " back" Or _
           Mid(strData, lIndex, 3) = " of" Or _
           Mid(strData, lIndex, 3) = " if" Or _
           Mid(strData, lIndex, 3) = " or" Then
           
            'Determine length of reserve words phrases...
            lLenChar = 3
            If Mid(strData, lIndex, 4) = " And" Then
                lLenChar = 4
            Else
                If Mid(strData, lIndex, 5) = " back" Then
                    lLenChar = 5
                End If
            End If
           
            'If at the end of the function name or the character following the reserve
            'words phrases is a blank then mark as an error...
            If lIndex + lLenChar > Len(strData) Then
                bErrorFound = True
            Else
                If Mid(strData, lIndex + lLenChar, 1) = " " Then
                    bErrorFound = True
                End If
            End If
            If bErrorFound Then
                strErrMsg = "The name: " & UCase(strData) & " cannot contain operators (IF, OR, OF, AND)."
                OperatorsFound = True
                Exit Function
            End If
        End If
    Next lIndex
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mCommon.OperatorsFound", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Color
'' Description: Color some text appropriately
'' Inputs:      Text to color
'' Returns:     Colored text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Color(strText As String) As String
On Error GoTo ErrSection:
    
    Const C_DEFAULT_CF = 1
    Const C_FUNCTION_CF = 2
    Const C_PARAM_CF = 3
    Const C_OPERATOR_CF = 4
    Const C_COMMENT_CF = 5
    Const C_ERROR_CF = 6
    Const C_PARENSTART_CF = 7
    
    Dim lBegOfStr As Long
    Dim lPos As Long
    Dim lLength As Long
    Dim alColor As cGdArray
    Dim alStart As cGdArray
    Dim alLength As cGdArray
    Dim lTokens As Long
    Dim strRichText As String
    Dim lIndex As Long
    Dim i As Integer
    
    Dim lParenNest As Long
    Dim lParenColor As Long
    Dim RT As New cRichText
    Dim strWorkText As String, strNextWorkText As String
    Dim rtfFormat As Byte

    'holds colors used for parenthesis
    Dim alParenColors As cGdArray
    Set alParenColors = New cGdArray
    alParenColors(0) = 33023
    alParenColors(1) = 16512
    alParenColors(2) = 8421504
    alParenColors(3) = 12582912
    alParenColors(4) = 12632064

    Dim hdr2 As String
    'initialize our richtext class
    With RT
        'assign control
        .RTBox = frmRichTextBox.RTB
        .CreateCustomFormat C_DEFAULT_CF
        With .RTBox
            .SelItalic = False
            .SelBold = False
        End With
        .CreateCustomFormat C_FUNCTION_CF
        With .RTBox
            .SelItalic = g.EditorOptions.FunctionsItalics
            .SelBold = g.EditorOptions.FunctionsBoldFace
            .SelColor = g.EditorOptions.FunctionsColor
        End With
        .CreateCustomFormat C_PARAM_CF
        With .RTBox
            .SelItalic = g.EditorOptions.ParmItalics
            .SelBold = g.EditorOptions.ParmBoldFace
            .SelColor = g.EditorOptions.ParmColor
        End With
        .CreateCustomFormat C_OPERATOR_CF
        With .RTBox
            .SelItalic = g.EditorOptions.OperatorsItalics
            .SelBold = g.EditorOptions.OperatorsBoldFace
            .SelColor = g.EditorOptions.OperatorsColor
        End With
        .CreateCustomFormat C_COMMENT_CF
        With .RTBox
            .SelItalic = False
            .SelBold = False
            .SelColor = RGB(0, 192, 0) 'vbGreen
        End With
        .CreateCustomFormat C_ERROR_CF
        With .RTBox
            .SelItalic = g.EditorOptions.ErrorItalics
            .SelBold = g.EditorOptions.ErrorBoldFace
            .SelColor = g.EditorOptions.ErrorColor
        End With
        'create formats used for parenthesis
        For i = 0 To alParenColors.Size - 1
            .CreateCustomFormat C_PARENSTART_CF + i
            With .RTBox
                .SelItalic = g.EditorOptions.OperatorsItalics
                .SelBold = g.EditorOptions.OperatorsBoldFace
                .SelColor = alParenColors(i)
            End With
        Next
    End With
 
    ' Used to store string information based on Tokens...
    Set alColor = New cGdArray
    Set alStart = New cGdArray
    Set alLength = New cGdArray
    
    ' Build color information arrays...
    lPos = 1: lTokens = 0
    Do Until lPos = 0
        lPos = InStr(lPos, strText, "~")
        If lPos > 0 Then
            lTokens = lTokens + 1
            alColor.Add Mid(strText, lPos + 1, 2)
            lLength = Val(Mid(strText, lPos + 3, 3))
            alLength.Add lLength
            
            'Calculate the beginning of the tag...
            'Calculate the beginning of the string (offset from the tag begin)
            lBegOfStr = (lPos + gTokenLen) - (lTokens * gTokenLen) - 1 'minus1 to adj for zero based
            'lBegOfStr = lBegOfStr + Val(Mid(.Text, lPos + 3, 3)) - 1
            alStart.Add lBegOfStr
            
            lPos = lPos + gTokenLen
        End If
    Loop
    
    ' Strip out lTokens
    lPos = 1
    Do Until lPos = 0
        lPos = InStr(1, strText, "~")
        If lPos > 0 Then
            If lPos > 1 Then
                strText = Left(strText, lPos - 1) & _
                        Right(strText, Len(strText) - _
                        (lPos + gTokenLen - 1))
            Else
                strText = Right(strText, Len(strText) - gTokenLen)
            End If
        End If
    Loop
    
    'Color strings based on array values
    For lIndex = 0 To alColor.Size - 1
        'what string are we working on
        strWorkText = Mid(strText, alStart(lIndex) + 1, alLength(lIndex))
        
        'sadly, we need to do a look-ahead for purpose of cosmetics
        'for speed - should use the lookup in the alColor array and check type, but this will work for now
        If lIndex + 1 <= alStart.Size - 1 Then
            strNextWorkText = Mid(strText, alStart(lIndex + 1) + 1, alLength(lIndex + 1))
        Else
            strNextWorkText = ""
        End If
    
        If Not (strWorkText = "(" Or strNextWorkText = ")" Or _
                strWorkText = "." Or strNextWorkText = ".") Then
            strWorkText = strWorkText & " " 'add a space
        End If
    
        'for purpose of nested paren coloring...
        If alColor(lIndex) = C_LEFTPAREN Or alColor(lIndex) = C_FLEFTPAREN Then
           lParenNest = lParenNest + 1
           lParenColor = (lParenNest Mod alParenColors.Size)
        End If
        If (alColor(lIndex) = C_RIGHTPAREN Or alColor(lIndex) = C_FRIGHTPAREN) And _
           lParenNest > 0 Then
           lParenColor = (lParenNest Mod alParenColors.Size)
           lParenNest = lParenNest - 1
        End If
        
        'safe use of format mask
        rtfFormat = 0
        'if anywhere in this string the user has placed the '\'
        'we must force it to be a literal or the rtf box will
        'think it is a formatting character
        strWorkText = Replace(strWorkText, "\", "\\")
        
        With g.EditorOptions
            Select Case alColor(lIndex)
                
                Case C_Tab
                    RT.AddText "\tab "
'                    strRichText = strRichText & "\tab "
                
                'Show first error only...
                Case C_ERROR
                    RT.AddText strWorkText, rtfUseCustomFormat + C_ERROR_CF
                
                Case C_COMMENT
                    RT.AddText strWorkText, rtfUseCustomFormat + C_COMMENT_CF
                
                Case C_BRACKETCOMMENT
                    strWorkText = Replace(strWorkText, "{", "\{")
                    strWorkText = Replace(strWorkText, "}", "\}")
                    RT.AddText strWorkText, rtfUseCustomFormat + C_COMMENT_CF
                
                Case C_TEXT
                    RT.AddText strWorkText, rtfUseCustomFormat + C_DEFAULT_CF
                
                Case C_FUNC_NUMERIC, C_FUNC_BOOLEAN, _
                     C_FUNC_NUMERIC_CONSTANT, C_FUNC_BOOLEAN_CONSTANT
                    RT.AddText strWorkText, rtfUseCustomFormat + C_FUNCTION_CF
                
                Case C_PARM_NUMERIC, C_PARM_BOOLEAN, C_PARM_BARS
                    RT.AddText strWorkText, rtfUseCustomFormat + C_PARAM_CF
                
                Case C_LEFTPAREN, C_RIGHTPAREN, _
                     C_FLEFTPAREN, C_FRIGHTPAREN
                    RT.AddText strWorkText, rtfUseCustomFormat + (C_PARENSTART_CF + lParenColor)
                
                Case C_ADD, C_MULTI, C_COMPARE, C_AND, C_NOT, C_OR, _
                     C_OFFSET, C_OF, C_IF, C_Then, C_Else, C_ElseIf, C_EndIf
                    RT.AddText strWorkText, rtfUseCustomFormat + C_OPERATOR_CF
                
                Case C_PARAGRAPH, C_EnterFormatting
                    RT.AddText "\par "
'                    strRichText = strRichText & "\par "
                    
                Case Else
                    RT.AddText strWorkText, rtfUseCustomFormat + C_DEFAULT_CF
            End Select
        End With
    Next lIndex
    
    'Enclose with trailing brace and stip out last blank format
'    strRichText = strRichText & "}"
    If Len(Trim(RT.Text)) > 0 Then
        RT.BuildRTF
        Color = RT.RTBox.TextRTF
        'sad state here: must force a replace of the fonttbl for compatibiltiy with daves system
        'Color = Replace(Color, "{\fonttbl{\f0\fnil\fcharset0 MS Sans Serif;}}", "{\fonttbl{\f0\fnil\fcharset0 Arial;}}")
    Else
        Color = ""
    End If

ErrExit:
'    Color = Replace(Color, "\f0\fnil\fcharset0 MS Sans Serif;", "\f0\fswiss Arial;")
    Exit Function
    
ErrSection:
    RaiseError "mCommon.Color", eGDRaiseError_Raise, g.strAppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FormatRTFHeader
'' Description: Get the RTF Header into the given string
'' Inputs:      String to put the header into
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FormatRTFHeader(strRichText As String)

    strRichText = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Arial;}{\f3\fswiss Arial;}}" & _
        "{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red255\green0\blue0;\red0\green150\blue0;}" & _
        "\deflang1033\pard"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HighlightSymbols
'' Description: Highlight symbols in the string
'' Inputs:      Properties to work with
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HighlightSymbols(strText As String, _
        strRTF As String, lCurPos As Long, alStart As cGdArray, _
        alLength As cGdArray, lColor As Long, bBold As Boolean, _
        bItalic As Boolean)
On Error GoTo ErrSection:

    Dim strWorkRTF As String
    Dim strWorkPlainText As String
    Dim lStartPos As Long
    Dim lLength As Long
    Dim lSlashPos As Long
    
    ' Use for getting the next phrase.  (Searching for right paren).
    Dim lNextStartPos As Long
    Dim lNextLength As Long
    Dim lNextPlainText As String
    
    ' Get current text information
    strWorkPlainText = Mid(strText, alStart(lCurPos) + 1, alLength(lCurPos))
    lStartPos = alStart(lCurPos)
    lLength = alLength(lCurPos)
    If lCurPos + 1 <= alStart.Size - 1 Then
        lNextStartPos = alStart(lCurPos + 1)
        lNextLength = alLength(lCurPos + 1)
        lNextPlainText = Mid(strText, lNextStartPos + 1, lNextLength)
    Else
        lNextStartPos = 0
    End If
    
    'font size 9> fs18  10>fs20)
    'font name f2> points to font number 2 in header
    strWorkRTF = "\plain\f2\fs20"
    
    'Insert color code...
    Select Case lColor
        Case vbBlue: strWorkRTF = strWorkRTF & "\cf1"
        Case vbRed: strWorkRTF = strWorkRTF & "\cf2"
        Case vbGreen: strWorkRTF = strWorkRTF & "\cf3"
    End Select
    
    'Insert bold
    If bBold Then
        strWorkRTF = strWorkRTF & "\b"
    End If
    
    'Insert Italic
    If bItalic Then
        strWorkRTF = strWorkRTF & "\i"
    End If

    'if text has a "\" then make it "\\"...
    lSlashPos = InStr(1, strWorkPlainText, "\")
    Do Until lSlashPos = 0
        strWorkPlainText = Left(strWorkPlainText, lSlashPos) & "\" & _
            Right(strWorkPlainText, Len(strWorkPlainText) - lSlashPos)
        lSlashPos = InStr(lSlashPos + 2, strWorkPlainText, "\")
    Loop
    
    
    'Insert text after RTF codes
    strRTF = strRTF & strWorkRTF & " " & strWorkPlainText
    
    'Insert blank after text (unless text is a parenthesis)
    If strWorkPlainText <> "(" And lNextPlainText <> ")" And _
        strWorkPlainText <> "." And lNextPlainText <> "." Then
        strRTF = strRTF & "\plain\f2\fs20" & "  "
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mCommon.HighlightSymbols", eGDRaiseError_Raise, g.strAppPath
    
End Sub
