VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStrh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cStrh.cls
'' Description: Handle certain things for TradeSense expressions
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/17/2009   DAJ         Handle comment before offset (.) operator
'' 05/15/2009   DAJ         Allow vbCrLf to be a word out of curWord
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Private Type mPrivate
    astrOperName() As String
    alOperPos() As Long
    astrOperAlpha() As String
    lOperTotal As Long
    lOperAlphaTotal As Long
End Type
Private m As mPrivate

'Insert 1 space bar character between all non blank consecutive characters
Public Sub InsertOneBlank(pText As String)
On Error GoTo ErrSection:

    Dim tmpstr          As String
    Dim TextLen         As Integer
    Dim curPos          As Integer
    Dim CurChar         As String
    
    tmpstr = ""
    TextLen = Len(pText)
    
    curPos = 1: CurChar = Mid(pText, curPos, 1)
    Do Until curPos > TextLen
    
        Do Until CurChar <> " "
            curPos = curPos + 1
            If curPos > TextLen Then
                GoTo EndIt:
            End If
            CurChar = Mid(pText, curPos, 1)
        Loop
        
        Do Until CurChar = " "
            tmpstr = tmpstr & CurChar
            curPos = curPos + 1
            If curPos > TextLen Then
                GoTo EndIt:
            End If
            CurChar = Mid(pText, curPos, 1)
        Loop
        
        tmpstr = tmpstr & CurChar   'Insert 1 blank
    Loop
    
EndIt:
    pText = tmpstr
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cStrH.InsertOneBlank", eGDRaiseError_Raise, g.strAppPath

End Sub

'This routine searches forward in pStr starting at pStart and searches for
'a nonblank character.  Its position is returned
Public Function FindNonBlank(pStr As String, pStart As Integer) As Integer
On Error GoTo ErrSection:
    
    Dim CurChar As String
    Dim X       As Integer
    For X = pStart To Len(pStr)
        CurChar = Mid(pStr, X, 1)
        If CurChar <> Chr(32) Then
            FindNonBlank = X
            Exit Function
        End If
    Next
    FindNonBlank = 0
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.FindNonBlank", eGDRaiseError_Raise, g.strAppPath

End Function

'Searches pSourceStr for pSearchFor, Replaces all occurrences of it with
'pReplaceWith string
Public Sub ReplaceStr(pSourceStr As String, pSearchFor As String, pReplaceWith As String)
On Error GoTo ErrSection:
    
    Dim pos         As Long
    Dim RightStart  As Long
    Dim X           As Long
    
    For X = 1 To Len(pSourceStr)
        pos = InStr(1, pSourceStr, pSearchFor)
        If pos > 0 Then
            RightStart = Len(pSourceStr) - pos + (Len(pReplaceWith) - 1)
            If RightStart > 0 Then
                pSourceStr = Left(pSourceStr, pos - 1) & pReplaceWith & _
                    Right(pSourceStr, Len(pSourceStr) - pos - _
                    (Len(pSearchFor) - 1))
            Else
                pSourceStr = Left(pSourceStr, pos - 1) & pReplaceWith
            End If
        Else
            Exit For
        End If
    Next X
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cStrH.ReplaceStr", eGDRaiseError_Raise, g.strAppPath

End Sub

'This routine searches backward in pStr starting at pStart and searches for
'pChar.  Its position is returned
Public Function FindPrevChar(pStr As String, pStart As Integer, _
                             pChar As String) As Integer
On Error GoTo ErrSection:

    Dim X           As Integer
    
    For X = pStart To 1 Step -1
        If Mid(pStr, X, 1) = pChar Then
            FindPrevChar = X
            Exit Function
        End If
    Next
    FindPrevChar = 0
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.FindPrevChar", eGDRaiseError_Raise, g.strAppPath

End Function

Public Function SeparatorWordFound(pword As String) As Boolean
On Error GoTo ErrSection:
    
    Dim X       As Integer
    
    'Note: quote and decimal point were missing from original case statement
    SeparatorWordFound = False
    For X = 1 To m.lOperTotal
        If pword = m.astrOperName(X) Then
            SeparatorWordFound = True
            Exit For
        End If
    Next X
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.SeparatorWordFound", eGDRaiseError_Raise, g.strAppPath

End Function

'Searches back for the current function starting at pStart and going
'backwards.  A separator signals the beginning of the function.
Public Function curFunction(pStr As String, pCaretPos As Integer) As String
On Error GoTo ErrSection:
    
    Dim FunctionBegPos  As Integer
    Dim FLen  As Integer
    
    curFunction = ""
    FunctionBegPos = FindFunctionBegPos(pStr, pCaretPos)
    FLen = pCaretPos - FunctionBegPos + 1
    If FLen > 0 Then
        curFunction = Trim(Mid(pStr, FunctionBegPos, FLen))
    Else
        curFunction = Mid(pStr, FunctionBegPos, 1)
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.curFunction", eGDRaiseError_Raise, g.strAppPath

End Function

'Returns the beginning position of the function
Public Function FindFunctionBegPos(pStr As String, pCaretPos As Integer)
On Error GoTo ErrSection:

    Dim SymbolEndPos    As Integer
    Dim FunctionBegPos  As Integer
    Dim Symbol          As String
    Dim CurChar         As String
    Dim X               As Integer
    
    SymbolEndPos = PrevSepSymbolPos(pStr, pCaretPos, Symbol)
    If SymbolEndPos = 1 And Symbol = "" Then
        FindFunctionBegPos = 1
        Exit Function
    End If
    
    'Find beginning position of function by searching forward for a non-
    'blank character.
    FunctionBegPos = 0
    For X = SymbolEndPos + 1 To Len(pStr)
        CurChar = Mid(pStr, X, 1)
        If CurChar <> Chr(32) Then
            FunctionBegPos = X
            Exit For
        End If
    Next
    
    If FunctionBegPos = 0 Then
        FunctionBegPos = SymbolEndPos
    End If
    
    'Function begin will always be at least 2 spaces to right of symbol
    'ending position.  Adjust if necessary
    If FunctionBegPos >= SymbolEndPos + 1 Then
        FindFunctionBegPos = FunctionBegPos
    Else
        'If Symbol = "and" Then
        '    FindFunctionBegPos = SymbolEndPos + 2
        'Else
            'handles math,compare symbols immediately preceding function name
            FindFunctionBegPos = FunctionBegPos + 1
        'End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.FindFunctionBegPos", eGDRaiseError_Raise, g.strAppPath

End Function

'Find matching function right parenthesis, return it's position
Public Function FunctionRightParen(pStr As String, pLeftParenStart As Long, _
    pRightParenPos As Long) As Boolean
On Error GoTo ErrSection:

    Dim X       As Long
    Dim pos     As Long
    Dim Parens  As Long
    Dim CurChar As String
    
    FunctionRightParen = False
    pRightParenPos = 0
    Parens = 0
    For X = pLeftParenStart To Len(pStr)
        CurChar = Mid(pStr, X, 1)
        Select Case CurChar
            Case "(": Parens = Parens + 1
            Case ")": Parens = Parens - 1
        End Select
        pos = X
        If Parens = 0 Then Exit For
    Next X
    If Parens = 0 Then
        FunctionRightParen = True
        pRightParenPos = pos
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.FunctionRightParen", eGDRaiseError_Raise, g.strAppPath

End Function

'Search back for a separator symbol (starting from .SelStart).
'The last position of the separator symbol is returned plus 1
'(this function kept for binary compatibility)
Public Function PrevSepSymbolPos(pStr As String, pStart As Integer, pSymbol As String) As Integer
    PrevSepSymbolPos = PrevSepSymbolPos2(pStr, pStart, pSymbol)
End Function

'Search back for a separator symbol (starting from .SelStart).
'The last position of the separator symbol is returned plus 1
'pFunc is used in cEditing.ReturnFunction and is the function the user selected. This
'allows us to check to make sure the first operator found isn't embedded in the function.
'For example: a user types in "if" which highlights "IFF" in the function list and
'presses enter. We don't want to view the "if" at the current cursor position as the
'operator "if". It is the start of the function "IFF".
Public Function PrevSepSymbolPos2(pStr As String, pStart As Integer, pSymbol As String, Optional pFunc As String) As Integer
On Error GoTo ErrSection:
    
    Dim X               As Integer
    Dim SymbolLen       As Integer
    Dim Fnd             As Boolean
    Dim Highest         As Long
    Dim HighestPos      As Long
    Dim OperStart       As Long
    Dim CurChar         As String
    
    SymbolLen = 0
    pSymbol = ""
    
    'InStrRev won't work until a valid position is found (sometimes pStart is > length)
    If pStart > Len(pStr) Then OperStart = Len(pStr) Else OperStart = pStart
    
    If OperStart = 0 Then
        PrevSepSymbolPos2 = 0
        pSymbol = ""
        Exit Function
    End If
    
'    Dim opStr As String
'    Dim opAStr() As String
'    Dim idx As Integer
'    opStr = pStr
'    opAStr = Split(pStr, " ")
'    If Not UBound(opAStr()) = LBound(opAStr()) Then
'        opStr = ""
'        For idx = 1 To UBound(opAStr())
'            opStr = opStr + opAStr(idx)
'        Next
'        pStr = opStr
'    End If
    'Search backwards through the string for the FIRST occurrence of an operator
    For X = 1 To m.lOperTotal
    
        ' Find the first occurence of operator in string that is not a part of a phrase
        m.alOperPos(X) = InStrRev(pStr, m.astrOperName(X), OperStart, vbTextCompare)
        Do While m.alOperPos(X) > 0
            'If operator is alphabetic make sure it is not part of a phrase
            If Not AndOrImbeddedInWord(pStr, m.alOperPos(X), m.astrOperName(X)) Then
                'Make sure this operator isn't embedded in the function input (if any)
                If InStr(1, pFunc, Mid(pStr, m.alOperPos(X), pStart - m.alOperPos(X) + 1)) <= 0 Then
                    Exit Do
                End If
            End If
            If m.alOperPos(X) <= 1 Then
                m.alOperPos(X) = 0
            Else
                m.alOperPos(X) = InStrRev(pStr, m.astrOperName(X), m.alOperPos(X) - 1, vbTextCompare)
            End If
        Loop
        
        'If minus is part of a number (negative number)
        'it is not a separator
        If m.astrOperName(X) = "-" And m.alOperPos(X) > 0 Then
            If IsNumeric(Mid(pStr, m.alOperPos(X) + 1, 1)) Or _
               Mid(pStr, m.alOperPos(X) + 1, 1) = "." Then
                m.alOperPos(X) = 0
            End If
        End If
        
        'Decimal point vs Offset logic...
        If m.astrOperName(X) = "." Then
            If m.alOperPos(X) > 1 Then
                Dim Z               As Integer
                Dim Phrase          As String
                Dim PrevPhraseEnd   As Integer
                Dim PrevPhraseBeg   As Integer
                
                'Search left of decimal point looking for first
                'non-blank character
                For Z = m.alOperPos(X) - 1 To 1 Step -1
                    If Mid(pStr, Z, 1) <> " " Then
                        PrevPhraseEnd = Z
                        Exit For
                    End If
                Next Z
                
                'Continue searching left for first blank character.  This
                'marks the beginning of the phrase...
                If PrevPhraseEnd > 0 Then
                    For Z = PrevPhraseEnd - 1 To 1 Step -1
                        If Mid(pStr, Z, 1) = " " Then
                            PrevPhraseBeg = Z
                            Exit For
                        End If
                    Next Z
                    
                    'Phrase left of decimal point...
                    If PrevPhraseBeg > 0 Then
                        Phrase = Trim(Mid(pStr, PrevPhraseBeg, PrevPhraseEnd - PrevPhraseBeg + 1))
                    Else
                        Phrase = Trim(Mid(pStr, 1, PrevPhraseEnd - PrevPhraseBeg))
                    End If
                    
                    'If phrase left of decimal point is not numeric (function or input)
                    'and not a minus sign and not a separate symbol (this means
                    'that the number is a fraction (ie .75)...
                    If Not IsNumeric(Phrase) And _
                       Phrase <> "-" And _
                       Not SeparatorWordFound(Phrase) Then
                    Else
                        If Phrase <> ")" Then
                            m.alOperPos(X) = 0
                        End If
                    End If
                End If
            End If
        End If
    
    Next X

    'Take the position that is highest means found first.
    Fnd = False
    Highest = -999999
    For X = 1 To m.lOperTotal
        If m.alOperPos(X) > Highest And m.alOperPos(X) > 0 Then
            Fnd = True
            Highest = m.alOperPos(X)
            HighestPos = X
        End If
    Next X
    
    'Return Symbol and it's position
    If Fnd Then
        pSymbol = m.astrOperName(HighestPos)
        SymbolLen = Len(m.astrOperName(HighestPos))
        PrevSepSymbolPos2 = Highest + SymbolLen - 1
    Else
        PrevSepSymbolPos2 = 1
        pSymbol = ""
    End If
    
    'If Symbol position being returned is a carriage return then move forward until it
    'isn't
    For X = PrevSepSymbolPos2 To Len(pStr)
        CurChar = Mid(pStr, X, 1)
        If CurChar <> Chr(13) And CurChar <> Chr(10) Then
            Exit For
        End If
    Next X
    PrevSepSymbolPos2 = X
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.PrevSepSymbolPos2", eGDRaiseError_Raise, g.strAppPath

End Function

'Searches back for the current word starting (search back until blank fnd).
Public Function curWord(pStr As String, pStart As Integer) As String
On Error GoTo ErrSection:
    
    Dim X               As Integer
    Dim CurChar         As String
    Dim curPos          As Integer
    Dim Symbol          As String
    Dim endpos          As Integer
    
    curWord = ""
    Symbol = ""
    
    For X = pStart To 1 Step -1
        CurChar = Mid(pStr, X, 1)
        
        ' DAJ 05/15/2009: Aardvark Issue #4931 states that the TradeSense window is not
        ' coming up with a space after a carriage return/line feed.  I tracked it down
        ' to the following line -- we don't want to back up past the carriage return/
        ' line feed anymore...
        If CurChar <> " " Then ' And CurChar <> Chr(13) And CurChar <> Chr(10) Then
            endpos = X
            Exit For
        End If
    Next X
    
    If endpos <= 0 Then Exit Function
    
    For X = endpos To 1 Step -1
        CurChar = Mid(pStr, X, 1)
        If CurChar = " " Then
            Exit Function
        End If
        If CurChar = Chr(13) Then
            curWord = Chr(13) & Chr(10)
            Exit Function
        Else
            'If current position has a symbol, then quit and return as the current
            'word
            curPos = X
            curPos = NextSepSymbolPos(pStr, curPos, curPos, Symbol)
            If Symbol <> "" Then
                Exit For
            End If
        End If
    Next
    
    'Separator symbol found, count forward the length of the symbol to get to the
    'current word
    If Symbol <> "" Then
        curWord = Trim(Mid(pStr, X + Len(Symbol), pStart - (X + Len(Symbol) - 1)))
        If curWord = "" Or curWord = Chr(13) & Chr(10) Then curWord = Symbol
    Else
        'Blank found, start counting one after blank and strip out word
        curWord = Trim(Mid(pStr, X + 1, pStart - X))
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.curWord", eGDRaiseError_Raise, g.strAppPath

End Function

'Returns the beginning position of the next separator symbol
Public Function NextSepSymbolPos(pStr As String, pStart As Integer, _
                    pEndPos As Integer, pSymbol As String) As Integer
On Error GoTo ErrSection:

    Dim X               As Integer
    Dim SymbolLen       As Integer
    Dim Fnd             As Boolean
    Dim Lowest          As Long
    Dim LowestPos       As Long
    Dim lInComment      As Long
    
    NextSepSymbolPos = 0
    pSymbol = ""
    SymbolLen = 0
    Fnd = False
    Lowest = 999999
    
    'Search for FIRST occurrences of operator's search backwards
    For X = 1 To m.lOperTotal
        m.alOperPos(X) = InStr(pStart, pStr, m.astrOperName(X), vbTextCompare)
                'no sense in continueing if nothing was there
        If m.alOperPos(X) = 0 Then GoTo NextX
        
        'If operator is alphabetic make sure it is not part of a phrase
        If m.alOperPos(X) > 0 Then
            If AndOrImbeddedInWord(pStr, m.alOperPos(X), m.astrOperName(X)) Then
                m.alOperPos(X) = 0
            End If
        End If
        
        'If minus is part of a number (negative number)
        'it is not a separator
        If m.astrOperName(X) = "-" Then
            If IsNumeric(Mid(pStr, m.alOperPos(X) + 1, 1)) Or _
               Mid(pStr, m.alOperPos(X) + 1, 1) = "." Then
                m.alOperPos(X) = 0
            End If
        End If
        
        'Decimal point vs Offset logic...
        If m.astrOperName(X) = "." Then
            If X > 1 Then
                Dim Z               As Integer
                Dim Phrase          As String
                Dim PrevPhraseEnd   As Integer
                Dim PrevPhraseBeg   As Integer
                
                lInComment = 0&
                
                'Search left of decimal point looking for first
                'non-blank character
                For Z = m.alOperPos(X) - 1 To 1 Step -1
                    If Mid(pStr, Z, 1) <> " " Then
                        If Mid(pStr, Z, 1) = "{" Then
                            lInComment = lInComment - 1&
                        ElseIf Mid(pStr, Z, 1) = "}" Then
                            lInComment = lInComment + 1&
                        ElseIf lInComment = 0 Then
                            PrevPhraseEnd = Z
                            Exit For
                        End If
                    End If
                Next Z
                
                'Continue searching left for first blank character.  This
                'marks the beginning of the phrase...
                If PrevPhraseEnd > 0 Then
                    For Z = PrevPhraseEnd - 1 To 1 Step -1
                        If Mid(pStr, Z, 1) = " " Then
                            If Mid(pStr, Z, 1) = "{" Then
                                lInComment = lInComment - 1&
                            ElseIf Mid(pStr, Z, 1) = "}" Then
                                lInComment = lInComment + 1&
                            ElseIf lInComment = 0 Then
                                PrevPhraseBeg = Z
                                Exit For
                            End If
                        End If
                    Next Z
                    
                    'Phrase left of decimal point...
                    If PrevPhraseBeg > 0 Then
                        Phrase = Trim(Mid(pStr, PrevPhraseBeg, PrevPhraseEnd - PrevPhraseBeg + 1))
                    Else
                        Phrase = Trim(Mid(pStr, 1, PrevPhraseEnd - PrevPhraseBeg))
                    End If
                    
                    'If phrase left of decimal point is not numeric (function or input)
                    'and not a minus sign and not a separate symbol (this means
                    'that the number is a fraction (ie .75)...
                    If Not IsNumeric(Phrase) And _
                       Phrase <> "-" And _
                       Not SeparatorWordFound(Phrase) Then
                    Else
                        'If phrase left is a right paren then it is a function right paren
                        If Phrase <> ")" Then
                            m.alOperPos(X) = 0
                        End If
                    End If
                End If
            End If
        End If

        If m.alOperPos(X) < Lowest And m.alOperPos(X) > 0 Then
            Fnd = True
            Lowest = m.alOperPos(X)
            LowestPos = X
        End If

NextX:
    Next X

'    'Take the position that is highest means found first.
'    Fnd = False
'    Lowest = 999999
'    For X = 1 To m.lOperTotal
'        If m.alOperPos(X) < Lowest And m.alOperPos(X) > 0 Then
'            Fnd = True
'            Lowest = m.alOperPos(X)
'            LowestPos = X
'        End If
'    Next
    
    If Not Fnd Then
        NextSepSymbolPos = pStart
    Else
        If Lowest > pEndPos Then
            NextSepSymbolPos = pEndPos
            pSymbol = ""
        Else
            NextSepSymbolPos = Lowest
            pSymbol = m.astrOperName(LowestPos)
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.NextSepSymbolPos", eGDRaiseError_Raise, g.strAppPath

End Function
  
'Check for And/or imbedded in a word
Public Function AndOrImbeddedInWord(pText As String, pCurPos As Long, _
    Symbol As String) As Boolean
On Error GoTo ErrSection:

    Dim X       As Integer
    
    AndOrImbeddedInWord = False
    If pCurPos - 1 > 0 Then
        If Mid(pText, pCurPos - 1, 1) <> " " And _
           Asc(Mid(pText, pCurPos - 1, 1)) <> 10 And _
           Mid(pText, pCurPos - 1, 1) <> vbTab Then
            For X = 1 To m.lOperAlphaTotal
                If Symbol = m.astrOperAlpha(X) Then
                    AndOrImbeddedInWord = True
                    Exit Function
                End If
            Next X
        End If
    End If
    
    If pCurPos + Len(Symbol) < Len(pText) Then
        If Mid(pText, pCurPos + Len(Symbol), 1) <> " " And _
           Asc(Mid(pText, pCurPos + Len(Symbol), 1)) <> 13 Then
            For X = 1 To m.lOperAlphaTotal
                If Symbol = m.astrOperAlpha(X) Then
                    AndOrImbeddedInWord = True
                    Exit Function
                End If
            Next X
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.AddOrImbeddedInWord", eGDRaiseError_Raise, g.strAppPath

End Function

'This routine determines the current parm that should be highlighted based
'on the position of the caret while the cInputsLine is displayed.
Public Function ParmToHighlight(pStr As String, pLeftParenPos As Integer, pCaretPos As Integer) As Byte
On Error GoTo ErrSection:
    
    Dim CurChar     As String
    Dim Commas      As Byte
    Dim Parens      As Integer
    Dim X           As Integer
    
    'Default to first parm
    ParmToHighlight = 0
    
    'Count the commas (which are used to separate parameters)
    'Skip past any imbedded functions (matching parenthesis of parm lists skip)
    Commas = 0
    For X = pCaretPos To pLeftParenPos Step -1
        If X = 0 Then Exit For
        CurChar = Mid(pStr, X, 1)
        If CurChar = "," Then
            Commas = Commas + 1
        Else
            If CurChar = ")" Then
                Parens = 1
                Do Until Parens = 0
                    X = X - 1
                    If X = 1 Then Exit Do
                    CurChar = Mid(pStr, X, 1)
                    If CurChar = "(" Then Parens = Parens - 1
                    If CurChar = ")" Then Parens = Parens + 1
                Loop
            End If
        End If
    Next
    ParmToHighlight = Commas + 1
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.ParmToHighlight", eGDRaiseError_Raise, g.strAppPath

End Function

'This routine determines the number of parameters currently in the function
'at the position of the caret.
Public Function ParmsInFxn(pStr As String, pCaretPos As Integer) As Byte
On Error GoTo ErrSection:

    ' Get the parameters to the right
    Dim CurChar     As String
    Dim nParms      As Integer
    Dim Parens      As Integer
    Dim X           As Integer
    Dim nLeftParen  As Integer
    nParms = 0
    
    ' Find the open left parenthesis for the current function
    nLeftParen = OpenLeftParenPos(pStr, pCaretPos)
    If nLeftParen <= 0 Then Exit Function
    nParms = 1
    
    'Count the commas (which are used to separate parameters)
    For X = nLeftParen + 1 To Len(pStr)
        If X > Len(pStr) Then Exit For
        CurChar = Mid(pStr, X, 1)
        If CurChar = "," Then
            nParms = nParms + 1
        ElseIf CurChar = "(" Then
            'Skip past any imbedded functions (matching parenthesis)
            Parens = 1
            Do Until Parens = 0
                X = X + 1
                If X > Len(pStr) Then Exit For
                CurChar = Mid(pStr, X, 1)
                If CurChar = ")" Then Parens = Parens - 1
                If CurChar = "(" Then Parens = Parens + 1
            Loop
        ElseIf CurChar = ")" Then
            ' Closing parenthesis
            Exit For
        End If
    Next
    
    ParmsInFxn = nParms
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.ParmsInFxn", eGDRaiseError_Raise, g.strAppPath

End Function

'This routine search backward in pStr starting at pStart and searches for
'an open left parenthesis.  A running total of left and right parens is
'used.  Right parens=-1 and left parens=1.  When the count equals 1 a left
'parenthesis has been found.  End if beginning of a line is encountered.
Public Function OpenLeftParenPos(pStr As String, pStart As Integer) As Integer
On Error GoTo ErrSection:

    Dim LParenPos   As Integer
    Dim Parens      As Integer
    Dim RParenFnd   As Boolean
    Dim CurChar     As String
    Dim X           As Integer
    
    'Count the commas (which are used to separate parameters)
    Parens = 0
    RParenFnd = False
    For X = pStart To 1 Step -1
        CurChar = Mid(pStr, X, 1)
        'If curChar = Chr(10) Then
            'Parens = 0
        '    Exit For      'Quit at beginning of line
        'End If
        
        If CurChar = ")" Then
            Parens = Parens - 1
            RParenFnd = True
        Else
            If CurChar = "(" Then
                Parens = Parens + 1
                LParenPos = X
            End If
        End If
        If Parens = 1 Then Exit For
        'If Parens = 1 And RParenFnd Then Exit For
    Next
    
    'If NO right parens were found, the starting point is to far to the
    'left side of the cInputsLine.  Continue
    If Parens > 0 Then
        OpenLeftParenPos = LParenPos
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.OpenLeftParenPos", eGDRaiseError_Raise, g.strAppPath

End Function

'This function returns the position of the last character in the cInputsLine
'1.  Search back from function Left paren beg and find matching parms
'2.  If no matching, go back to the last parm (last comma) and then go
'       forward for the nearest separator.
Public Function CalInputsLineEnd(pStr As String, pStart As Integer) As Integer
On Error GoTo ErrSection:

    Dim CommaPos    As Integer
    Dim NextSepPos  As Integer
    Dim NextSymbol  As String
    Dim ParenCnt    As Integer
    Dim X           As Integer
    Dim CurChar     As String
    
    'Find first left paren
    ParenCnt = 1
    For X = pStart To Len(pStr)
        CurChar = Mid(pStr, X, 1)
        If CurChar = "(" Then
            ParenCnt = ParenCnt + 1
        Else
            If CurChar = ")" Then
                ParenCnt = ParenCnt - 1
            Else
                If CurChar = "," Then
                    CommaPos = X
                End If
            End If
        End If
        
        If ParenCnt = 0 Then Exit For
    Next X
    
    'Matching parens found, use last paren as ending position
    If ParenCnt = 0 Then
        CalInputsLineEnd = X
    Else
        'No parens were entered, use the length of the text as the end
        If CommaPos = 0 Then
            CalInputsLineEnd = Len(pStr)
        Else
            'Search forward for separator symbol from last comma position
            NextSepPos = NextSepSymbolPos(pStr, CommaPos, Len(pStr), NextSymbol)
            If NextSepPos > 0 Then
                CalInputsLineEnd = NextSepPos
            Else
                CalInputsLineEnd = CommaPos
            End If
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.CallInputsLineEnd", eGDRaiseError_Raise, g.strAppPath

End Function

'Returns true if string is a valid numeric
Public Function IsANumeric(pText As String) As Boolean
On Error GoTo ErrSection:

    Dim curPos          As Integer
    Dim CurChar         As String
    Dim FirstNbrFnd     As Boolean
    Dim X               As Integer
    
    'If minus is all by itself is it NOT a numeric
    If pText = "-" Then
        IsANumeric = False
        Exit Function
    End If
    
    IsANumeric = True
    
    ' Search for Exit percentage parameter
    For curPos = 1 To Len(pText)
        CurChar = Mid(pText, curPos, 1)
        If (CurChar >= "0" And CurChar <= "9") Or CurChar = "-" Or _
            CurChar = "." Then
            FirstNbrFnd = True
        Else
            If Not FirstNbrFnd Then
                'First character must be 0 to 9, minus, or dec point
                IsANumeric = False
                Exit Function
            Else
                'Trailing percent sign is allowed
                If Not CurChar = "%" Then
                    IsANumeric = False
                End If
            End If
        End If
    Next
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cStrH.IsANumeric", eGDRaiseError_Raise, g.strAppPath

End Function

Private Sub Class_Initialize()
        
    ReDim m.astrOperName(100)
    ReDim m.alOperPos(100)
    ReDim m.astrOperAlpha(100)
    m.lOperTotal = 31
    
    m.astrOperName(1) = "elseif"
    m.astrOperName(2) = "Endif"
    m.astrOperName(3) = "then"
    m.astrOperName(4) = "Else"
    m.astrOperName(5) = "and"
    m.astrOperName(6) = "not"
    m.astrOperName(7) = "or"
    m.astrOperName(8) = "of"
    m.astrOperName(9) = "if"
    m.astrOperName(10) = Chr(13) & Chr(10)
    m.astrOperName(11) = Chr(9)
    m.astrOperName(12) = ">="
    m.astrOperName(13) = "<="
    m.astrOperName(14) = "<>"
    m.astrOperName(15) = ">"
    m.astrOperName(16) = "<"
    m.astrOperName(17) = "="
    m.astrOperName(18) = "+"
    m.astrOperName(19) = "-"
    m.astrOperName(20) = "*"
    m.astrOperName(21) = "/"
    m.astrOperName(22) = "("
    m.astrOperName(23) = ")"
    m.astrOperName(24) = ","
    m.astrOperName(25) = "."
    m.astrOperName(26) = "'"
    m.astrOperName(27) = """"
    m.astrOperName(28) = "DoUntil"
    m.astrOperName(29) = "EndDo"
    m.astrOperName(30) = "{"
    m.astrOperName(31) = "}"
    
    'Points to array indexes in m.astrOperName that are alphabetic
    m.astrOperAlpha(1) = "elseif"
    m.astrOperAlpha(2) = "Endif"
    m.astrOperAlpha(3) = "then"
    m.astrOperAlpha(4) = "Else"
    m.astrOperAlpha(5) = "and"
    m.astrOperAlpha(6) = "not"
    m.astrOperAlpha(7) = "or"
    m.astrOperAlpha(8) = "of"
    m.astrOperAlpha(9) = "if"
    m.astrOperAlpha(10) = "DoUntil"
    m.astrOperAlpha(11) = "EndDo"
    m.lOperAlphaTotal = 11
    
End Sub

Private Sub Class_Terminate()
    
    Erase m.astrOperName
    Erase m.alOperPos

End Sub
