VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDbTrees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
'This class is responsible for building a tree of phrases.
'Rules: - Internal functions are inserted into tree
'       - Options inputs of functions are defaulted into tree (if not specified)

'Work variables
Private mFunctions          As cFunctions
Private mLast               As Long
Private mTradesAddr         As Long
Private mPortfolioAddr      As Long
Private mSystemsAddr        As Long
Private mErrorCode          As Byte
Private mCalledFromEditor   As Boolean
Private mRetVal             As Variant

'for testing gdArray destroying
Private mAddr               As cGdArray

'Trees database...
Private mPhraseType         As cGdArray
Private mReturnType         As cGdArray
Private mPhrase             As cGdArray
Private mFormatting         As cGdArray  'Array of user formatting in phrase
Private mPhrasesArray       As cGdArray  'Array Sub phrases
Private mPhraseTypesArray   As cGdArray  'Array of sub phrase types
Private mPhraseFormattingArray As cGdArray
Private mClassName          As cGdArray  'Class name of function
Private mClassPtr           As cGdArray  'Ptr to array holding class instance
Private mLate               As cGdArray
Private mSize               As cGdArray  'Array size set and allocated (debugging)
Private mNodeFrom           As cGdArray
Private mNodeTo             As cGdArray
Private mParent             As cGdArray
Private mRuleID             As cGdArray
Private mExeTrue            As cGdArray
Private mExeFalse           As cGdArray
Private mPoolAddr           As cGdArray  'Holds the address of the array
                                         'for function types and expr's. It
                                         'holds a single value for nbr's.
Private mParmList           As cGdArray  'Address to Parmlist
Private mRebuildArray       As cGdArray  '1=Rebuild array.  Set to 1 when
                                         'New parms are inserted during opt.
                                         'This forces the array to be rebuilt
Private mFID                As cGdArray  'Function ID (if function)
Private mInputID            As cGdArray  'Input ID (if parent is function)
Private mReq                As cGdArray  '1=Function Input is required
Private mEntered            As cGdArray  '1=Phrase was user entered, not defaulted
Private mError              As cGdArray  '1=Phrase is in error

'Variabls to store gdarray addresses...
Private aPhraseType         As Long
Private aReturnType         As Long
Private aPhrase             As Long
Private aFormatting         As Long
Private aPhrasesArray       As Long
Private aPhraseTypesArray   As Long
Private aPhraseFormattingArray  As Long
Private aClassName          As Long
Private aClassPtr           As Long
Private aLate               As Long
Private aSize               As Long
Private aNodeFrom           As Long
Private aNodeTo             As Long
Private aParent             As Long
Private aRuleID             As Long
Private aExeTrue            As Long
Private aExeFalse
Private aPoolAddr           As Long
Private aParmList           As Long
Private aRebuildArray       As Long
Private aFID                As Long
Private aInputID            As Long
Private aReq                As Long
Private aEntered            As Long
Private aError              As Long

'Vector's used to store bottom node for each statement (line) in rule
Private mStatementBottom    As cGdArray
Private mMaxNode            As Long

Private Const eClass = "Class:cDbTrees "

'The first error found is flagged is saved to mError while the tree is built.
Property Get ErrMsg() As String
    Select Case mErrorCode
        Case 1: ErrMsg = "The expressions in this operation are not compatible (ie text cannot be used with numbers)."
        Case 2: ErrMsg = "TradeSense could not interpret one or more characters to the right of this function."
        Case 3: ErrMsg = "There are too many inputs entered for this function"
        Case 4: ErrMsg = "Invalid use of this operation (+,-,>,And,etc.)"
        Case 5: ErrMsg = "Text expressions cannot be used in math operations"
        Case 6: ErrMsg = "Expression could not be evaluated, please check the syntax"
        Case 7: ErrMsg = "One or more inputs are required for this function but were not entered"
        Case 8: ErrMsg = "An expression is required to the right of the IF, ELSEIF, DoUntil, and NOT."
        Case 9: ErrMsg = "An error was detected in one or more phrases before IF, ELSEIF, DoUntil, or NOT.  Please recheck these phrases and reverify."
        Case 10: ErrMsg = "ELSE and ENDIF must be used by themselves on their own line."
        Case 11: ErrMsg = "An error was detected in the phrases after THEN.  Please recheck these phrases and reverify."
        Case 12: ErrMsg = "Function or operator argument type mismatch."
        Case Else
            ErrMsg = ""
    End Select
End Property

'Must be set prior to calling "BuildTree"
Property Let Functions(pData As cFunctions)
    Set mFunctions = pData
End Property
Property Let TradesAddr(pData As Long)
    mTradesAddr = pData
End Property

'Keep to maintain binary compatibility
Property Let ModelAddr(pData As Long)
End Property
'This is the good one
Property Let PortfolioAddr(pData As Long)
    mPortfolioAddr = pData
End Property

Property Let SystemsAddr(pData As Long)
    mSystemsAddr = pData
End Property

'Returns pointers to bottom of tree of each statement in all rules
Property Get StatementBottom() As Long
    StatementBottom = mStatementBottom.ArrayHandle
End Property

'Returns the address of each gdarray for each node in the tree
Property Get AddrPhraseType() As Long
    AddrPhraseType = aPhraseType
End Property
Property Get AddrReturnType() As Long
    AddrReturnType = aReturnType
End Property
Property Get AddrPhrase() As Long
    AddrPhrase = aPhrase
End Property
Property Get AddrFormatting() As Long
    AddrFormatting = aFormatting
End Property
Property Get AddrClassName() As Long
    AddrClassName = aClassName
End Property
Property Get AddrClassPtr() As Long
    AddrClassPtr = aClassPtr
End Property
Property Get AddrLate() As Long
    AddrLate = aLate
End Property
Property Get AddrSize() As Long
    AddrSize = aSize
End Property
Property Get AddrNodeFrom() As Long
    AddrNodeFrom = aNodeFrom
End Property
Property Get AddrNodeTo() As Long
    AddrNodeTo = aNodeTo
End Property
Property Get AddrParent() As Long
    AddrParent = aParent
End Property
Property Get AddrRuleID() As Long
    AddrRuleID = aRuleID
End Property
Property Get AddrExeTrue() As Long
    AddrExeTrue = aExeTrue
End Property
Property Get AddrExeFalse() As Long
    AddrExeFalse = aExeFalse
End Property
Property Get AddrPoolAddr() As Long
    AddrPoolAddr = aPoolAddr
End Property
Property Get AddrParmList() As Long
    AddrParmList = aParmList
End Property
Property Get AddrReBuildArray() As Long
    AddrReBuildArray = aRebuildArray
End Property
Property Get AddrFID() As Long
    AddrFID = aFID
End Property
Property Get AddrInputID() As Long
    AddrInputID = aInputID
End Property
Property Get AddrReq() As Long
    AddrReq = aReq
End Property
Property Get AddrEntered() As Long
    AddrEntered = aEntered
End Property
Property Get AddrError() As Long
    AddrError = aError
End Property

'Properties that return 1 element of a phrase node from the Tree
Property Get PhraseType(ByVal i As Long) As Byte
    PhraseType = gdGetNum(aPhraseType, i)
End Property
Property Get ReturnType(ByVal i As Long) As Byte
    ReturnType = gdGetNum(aReturnType, i)
End Property
Property Get Phrase(ByVal i As Long) As String
    Phrase = gdGetStr(aPhrase, i)
End Property
Property Get Formatting(ByVal i As Long) As String
    Formatting = gdGetStr(aFormatting, i)
End Property
Property Get PhrasesArray(ByVal i As Long) As Long
    PhrasesArray = gdGetNum(aPhrasesArray, i)
End Property
Property Get PhraseTypesArray(ByVal i As Long) As Long
    PhraseTypesArray = gdGetNum(aPhraseTypesArray, i)
End Property
Property Get ClassName(ByVal i As Long) As String
    ClassName = gdGetStr(aClassName, i)
End Property
Property Get ClassPtr(ByVal i As Long) As Long
    ClassPtr = gdGetNum(aClassPtr, i)
End Property
Property Get Late(ByVal i As Long) As Byte
    Late = gdGetNum(aLate, i)
End Property
Property Get Size(ByVal i As Long) As Long
    Size = gdGetNum(aSize, i)
End Property
Property Get NodeFrom(ByVal i As Long) As Long
    NodeFrom = gdGetNum(aNodeFrom, i)
End Property
Property Get NodeTo(ByVal i As Long) As Long
    NodeTo = gdGetNum(aNodeTo, i)
End Property
Property Get Parent(ByVal i As Long) As Long
    Parent = gdGetNum(aParent, i)
End Property
Property Get RuleID(ByVal i As Long) As Long
    RuleID = gdGetNum(aRuleID, i)
End Property
Property Get ExeTrue(ByVal i As Long) As Long
    ExeTrue = gdGetNum(aExeTrue, i)
End Property
Property Get ExeFalse(ByVal i As Long) As Long
    ExeFalse = gdGetNum(aExeFalse, i)
End Property
Property Get PoolAddr(ByVal i As Long) As Long
    PoolAddr = gdGetNum(aPoolAddr, i)
End Property
Property Get ParmList(ByVal i As Long) As Long
    ParmList = gdGetNum(aParmList, i)
End Property
Property Get ReBuildArray(ByVal i As Long) As Byte
    ReBuildArray = gdGetNum(aRebuildArray, i)
End Property
Property Get FID(ByVal i As Long) As Long
    FID = gdGetNum(aFID, i)
End Property
Property Get InputID(ByVal i As Long) As Long
    InputID = gdGetNum(aInputID, i)
End Property
Property Get Req(ByVal i As Long) As Byte
    Req = gdGetNum(aReq, i)
End Property
Property Get Entered(ByVal i As Long) As Byte
    Entered = gdGetNum(aEntered, i)
End Property
Property Get PhraseError(ByVal i As Long) As Byte
    PhraseError = gdGetNum(aError, i)
End Property

Property Get Last() As Long
    Last = mLast
End Property


Private Sub SetPhraseType(pPos As Long, pValue As Byte)
    gdSetNum aPhraseType, pPos, pValue
End Sub
Private Sub SetReturnType(pPos As Long, pValue As Byte)
    gdSetNum aReturnType, pPos, pValue
End Sub
Private Sub SetPhrase(pPos As Long, pValue As String)
    gdSetStr aPhrase, pPos, pValue
End Sub
Private Sub SetFormatting(pPos As Long, pValue As String)
    gdSetStr aFormatting, pPos, pValue
End Sub
Private Sub SetPhrasesArray(pPos As Long, pValue As Long)
    gdSetNum aPhrasesArray, pPos, pValue
End Sub
Private Sub SetPhraseTypesArray(pPos As Long, pValue As Long)
    gdSetNum aPhraseTypesArray, pPos, pValue
End Sub
Private Sub SetClassName(pPos As Long, pValue As String)
    gdSetStr aClassName, pPos, pValue
End Sub
Private Sub SetClassPtr(pPos As Long, pValue As Long)
    gdSetNum aClassPtr, pPos, pValue
End Sub
Private Sub SetLate(pPos As Long, pValue As Byte)
    gdSetNum aLate, pPos, pValue
End Sub
Private Sub SetSize(pPos As Long, pValue As Long)
    gdSetNum aSize, pPos, pValue
End Sub
Private Sub SetNodeFrom(pPos As Long, pValue As Long)
    gdSetNum aNodeFrom, pPos, pValue
End Sub
Private Sub SetNodeTo(pPos As Long, pValue As Long)
    gdSetNum aNodeTo, pPos, pValue
End Sub
Private Sub SetParent(pPos As Long, pValue As Long)
    gdSetNum aParent, pPos, pValue
End Sub
Private Sub SetRuleID(pPos As Long, pValue As Long)
    gdSetNum aRuleID, pPos, pValue
End Sub
Private Sub SetExeTrue(pPos As Long, pValue As Long)
    gdSetNum aExeTrue, pPos, pValue
End Sub
Private Sub SetExeFalse(pPos As Long, pValue As Long)
    gdSetNum aExeFalse, pPos, pValue
End Sub
Private Sub SetPoolAddr(pPos As Long, pValue As Long)
    gdSetNum aPoolAddr, pPos, pValue
End Sub
Private Sub SetParmList(pPos As Long, pValue As Long)
    gdSetNum aParmList, pPos, pValue
End Sub
Private Sub SetRebuildArray(pPos As Long, pValue As Byte)
    gdSetNum aRebuildArray, pPos, pValue
End Sub
Private Sub SetFID(pPos As Long, pValue As Long)
    gdSetNum aFID, pPos, pValue
End Sub
Private Sub SetInputID(pPos As Long, pValue As Long)
    gdSetNum aInputID, pPos, pValue
End Sub
Private Sub SetReq(pPos As Long, pValue As Byte)
    gdSetNum aReq, pPos, pValue
End Sub
Private Sub SetEntered(pPos As Long, pValue As Byte)
    gdSetNum aEntered, pPos, pValue
End Sub
Private Sub SetError(pPos As Long, pValue As Byte)
    gdSetNum aError, pPos, pValue
End Sub

'==========================================================================
'THIS IS THE ENTRY POINT INTO THE TREE.  USE THIS ROUTINE TO ADD
'CODEDTEXT INTO THE TREE.
Public Sub AddCodedText(pCodedText As String, pName As String)
On Error GoTo ErrSection:
    
    Dim X               As Long
    Dim Y               As Long
    Dim Statement       As String
    Dim tPhrases        As Long
    Dim tPhraseTypes    As Long
    Dim tPhraseFormatting As Long
    Dim aPhrs           As Long
    Dim aPhrTypes       As Long
    Dim aPhrFormatting  As Long
    Dim PhrCnt          As Long
    Dim EnterFnd        As Boolean
    Dim BegPos          As Long
    Dim endpos          As Long
    
    mCalledFromEditor = False
    
    ValidateCodedText pCodedText, pName
    ConvertCT pCodedText, tPhrases, tPhraseTypes, tPhraseFormatting
    
    For X = 1 To gdGetSize(tPhrases) - 1
    
        'Search forward for an <Enter> key phrase type...
        EnterFnd = False
        BegPos = X
        For Y = X To gdGetSize(tPhrases) - 1
            If gdGetNum(tPhraseTypes, Y) = gPT_Enter Then
                EnterFnd = True
                endpos = Y
                Exit For
            End If
        Next Y
        
        If EnterFnd Then
        
            'Create arrays for storing phrases and phrase types of
            'stuff to left of <Enter>
            aPhrs = gdCreateArray(eGDARRAY_Strings, endpos - BegPos + 1, USE_DEFAULT_NULL)
            aPhrTypes = gdCreateArray(eGDARRAY_TinyInts, endpos - BegPos + 1, USE_DEFAULT_NULL)
            aPhrFormatting = gdCreateArray(eGDARRAY_Strings, endpos - BegPos + 1, USE_DEFAULT_NULL)
            mAddr.Add aPhrs
            mAddr.Add aPhrTypes
            mAddr.Add aPhrFormatting
   
            'Get rule statement up to the <enter> key
            Statement = ""
            PhrCnt = 0
            For Y = BegPos To endpos
                Statement = Statement & gdGetStr(tPhrases, Y) & " "
                PhrCnt = PhrCnt + 1
                gdSetStr aPhrs, PhrCnt, gdGetStr(tPhrases, Y)
                gdSetNum aPhrTypes, PhrCnt, gdGetNum(tPhraseTypes, Y)
                gdSetStr aPhrFormatting, PhrCnt, gdGetStr(tPhraseFormatting, Y)
            Next Y
            Statement = Trim(Statement)
        
            'Add rule statement to next level in tree
            Add 0, 0, Statement, aPhrs, aPhrTypes, aPhrFormatting, "", _
                0, 0, 0, 0, 0, 0, 1
                
            'Set X to the EndPos which is the enter key.  When returning
            'to the top of the loop, X is incremented by 1 to start the
            'next statement
            X = endpos
        
            'for debugging
            'Dim Z   As Integer
            'Debug.Print Chr(13) & Chr(10)
            'Debug.Print Statement & Chr(13) & Chr(10)
            'For Z = 1 To gdGetSize(aPhrs) - 1
            '    Debug.Print gdGetStr(aPhrs, Z) & ", " & gdGetNum(aPhrFormatting, Z)
            'Next Z
            
        End If
    Next X

    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:AddCodedText ", Err.Description
    End If
End Sub
        
'This is called from cExpression.  It adds each rule statement to the top
'level of the tree.  Each statement (comment, If, or action) is a separate
'node.
Public Sub AddEnglishText(pPhrases As Long, pPhraseTypes As Long, pPhraseFormatting As Long)
On Error GoTo ErrSection:
    Dim X           As Long
    Dim Y           As Long
    Dim Statement   As String
    Dim aPhrs       As Long
    Dim aPhrTypes   As Long
    Dim aPhrFormatting As Long
    Dim PhrCnt      As Long
    Dim EnterFnd    As Boolean
    Dim BegPos      As Long
    Dim endpos      As Long
    mCalledFromEditor = True
    
    For X = 1 To gdGetSize(pPhrases) - 1
    
        'Search forward for an <Enter> key phrase type...
        EnterFnd = False
        BegPos = X
        For Y = X To gdGetSize(pPhrases) - 1
            If gdGetNum(pPhraseTypes, Y) = gPT_Enter Then
                EnterFnd = True
                endpos = Y
                Exit For
            End If
        Next Y
        
        If EnterFnd Then
        
            'Create arrays for storing phrases and phrase types of
            'stuff to left of <Enter>
            aPhrs = gdCreateArray(eGDARRAY_Strings, endpos - BegPos + 1, USE_DEFAULT_NULL)
            aPhrTypes = gdCreateArray(eGDARRAY_TinyInts, endpos - BegPos + 1, USE_DEFAULT_NULL)
            aPhrFormatting = gdCreateArray(eGDARRAY_Strings, endpos - BegPos + 1, USE_DEFAULT_NULL)
            mAddr.Add aPhrs
            mAddr.Add aPhrTypes
            mAddr.Add aPhrFormatting
   
            'Get rule statement up to the <enter> key
            Statement = ""
            PhrCnt = 0
            For Y = BegPos To endpos
                Statement = Statement & gdGetStr(pPhrases, Y) & " "
                PhrCnt = PhrCnt + 1
                gdSetStr aPhrs, PhrCnt, gdGetStr(pPhrases, Y)
                gdSetNum aPhrTypes, PhrCnt, gdGetNum(pPhraseTypes, Y)
                gdSetStr aPhrFormatting, PhrCnt, gdGetStr(pPhraseFormatting, Y)
            Next Y
            Statement = Trim(Statement)
        
            'Add rule statement to next level in tree
            Add 0, 0, Statement, aPhrs, aPhrTypes, aPhrFormatting, "", _
                0, 0, 0, 0, 0, 0, 1
                
            'Set X to the EndPos which is the enter key.  When returning
            'to the top of the loop, X is incremented by 1 to start the
            'next statement
            X = endpos
        
            'for debugging
            'Dim Z   As Integer
            'Debug.Print Chr(13) & Chr(10)
            'Debug.Print Statement & Chr(13) & Chr(10)
            'For Z = 1 To gdGetSize(aPhrs) - 1
            '    Debug.Print gdGetStr(aPhrs, Z) & ", " & gdGetNum(aPhrFormatting, Z)
            'Next Z
            
        End If
    Next X
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:AddEnglishText ", Err.Description
    End If
End Sub
'==========================================================================


'Add's a phrase to the Tree...
Private Sub Add(pPhraseType As Byte, pReturnType As Byte, _
    pPhrase As String, pPhrasesArrayAddr As Long, pPhraseTypesArrayAddr As Long, _
    pPhraseFormatting As Long, pClassName As String, pLate As Byte, _
    pNodeFrom As Long, pNodeTo As Long, pParent As Long, _
    pPoolAddr As Long, pReq As Byte, pEntered As Byte)
On Error GoTo ErrSection:
    
    mLast = mLast + 1
    gdSetNum aPhraseType, mLast, pPhraseType
    gdSetNum aReturnType, mLast, pReturnType
    gdSetStr aPhrase, mLast, pPhrase
    gdSetStr aFormatting, mLast, ""
    gdSetNum aPhrasesArray, mLast, pPhrasesArrayAddr
    gdSetNum aPhraseTypesArray, mLast, pPhraseTypesArrayAddr
    gdSetNum aPhraseFormattingArray, mLast, pPhraseFormatting
    gdSetStr aClassName, mLast, pClassName
    gdSetNum aClassPtr, mLast, 0
    gdSetNum aLate, mLast, pLate
    gdSetNum aSize, mLast, 0
    gdSetNum aNodeFrom, mLast, pNodeFrom
    gdSetNum aNodeTo, mLast, pNodeTo
    gdSetNum aParent, mLast, pParent
    gdSetNum aRuleID, mLast, 0
    gdSetNum aExeTrue, mLast, 0
    gdSetNum aExeFalse, mLast, 0
    gdSetNum aPoolAddr, mLast, pPoolAddr
    gdSetNum aParmList, mLast, 0
    gdSetNum aRebuildArray, mLast, 1
    gdSetNum aFID, mLast, 0
    gdSetNum aInputID, mLast, 0
    gdSetNum aReq, mLast, pReq
    gdSetNum aEntered, mLast, pEntered
    gdSetNum aError, mLast, 0
    
    Exit Sub
ErrSection:
    Err.Raise Err.Number, eClass & "Sub:Add ", Err.Description
End Sub

'Search through the array of codedtext phrases and return the operator of
'the highest presidence along with it's array position.
Private Sub DeterminePhraseType(pPhraseTypes As Long, _
    pPhraseFormattingArray As Long, _
    pHighestPrecidencePhraseType As Byte, pPhraseTypeArrayPos As Long, _
    pUserFormatting As String)
On Error GoTo ErrSection:
    Dim X               As Long
    Dim Cnt             As Long
    Dim curType         As Long
    Dim P               As Long
    Dim Prec            As Byte
    Dim MaxPrec         As Byte
    Dim MaxPrecPos      As Long
    Dim MaxPrecType     As Byte
    Dim LeftParenFnd    As Boolean
    Dim RightParenFnd   As Boolean
    
    pHighestPrecidencePhraseType = 0
    pPhraseTypeArrayPos = 0
    Cnt = gdGetSize(pPhraseTypes) - 1
    
    'If the first phrase is a left paren then potentially, this is an
    'expression surrounded by outerparens
    If Cnt >= 1 Then
        If gdGetNum(pPhraseTypes, 1) = gPT_LeftPar Then
            LeftParenFnd = True
        End If
    End If
    RightParenFnd = True
    
    For X = 1 To Cnt
        curType = gdGetNum(pPhraseTypes, X)
        
        'Skip past parens and double quotes...
        Select Case curType
            Case gPT_FLParen, gPT_LeftPar
                P = P + 1
            Case gPT_FRParen, gPT_RightPar
                P = P - 1
        End Select
        
        If P = 0 Then
            
            If X < Cnt And LeftParenFnd Then
                RightParenFnd = False
            End If
            
            Select Case curType
                Case gPT_Enter
                    Prec = 11
                Case gPT_If, gPT_DoUntil
                    Prec = 10
                Case gPT_ElseIf, gPT_Else
                    Prec = 9
                Case gPT_Then
                    Prec = 8
                Case gPT_And, gPT_Or
                    Prec = 7
                Case gPT_Not
                    Prec = 6
                Case gPT_GT, gPT_GE, gPT_LT, gPT_LE, gPT_NE, gPT_EQ
                    Prec = 5
                Case gPT_Add, gPT_Sub
                    Prec = 4
                Case gPT_Mult, gPT_Div
                    Prec = 3
                Case gPT_Of
                    Prec = 2
                Case gPT_Offset
                    Prec = 1
            End Select
            If Prec > MaxPrec Then
                MaxPrecType = curType
                MaxPrec = Prec
                MaxPrecPos = X
            End If
        End If
    Next X
    
    'If operator presidence not found, then take the first phrase of the
    'array as the phrase type (Function, number, etc.)
    If MaxPrecType = 0 Then
    
        'Mark as outerparens
        If LeftParenFnd And RightParenFnd Then
            pHighestPrecidencePhraseType = gPT_OuterParens
            pPhraseTypeArrayPos = 0
            pUserFormatting = gdGetStr(pPhraseFormattingArray, 1)
            Exit Sub
        End If
        
        curType = gdGetNum(pPhraseTypes, 1)
        
        'Phrase 1 is a function...
        If curType = gPT_FCompiled Or curType = gPT_FCompiledAction Or _
           curType = gPT_FTradeSense Or curType = gPT_FInternal Then
            
            'IF the function has no parms then 1 element will be in pPhrasetypes.
            'This is OK, otherwise, ensure that the last phrase is the
            'function's right paren.  Anything right of it is junk.
            If gdGetNum(pPhraseTypes, gdGetSize(pPhraseTypes) - 1) = gPT_FRParen Or _
               Cnt = 1 Then
                pHighestPrecidencePhraseType = gdGetNum(pPhraseTypes, 1)
            Else
                'Last phrase is NOT a function right paren.  It's stuff that
                'shouldn't be there.  Flag as error.
                pHighestPrecidencePhraseType = 0
            End If
            
        Else
            'If only one phrase in array and it's a nbr, text, or parm then
            'this is valid.  Set the phrase type.
            If curType = gPT_Nbr Or curType = gPT_Text Or gPT_Comment Or _
               curType = gPT_Parm Then
            
                'If multiple phrases found then this can't be interpreted.
                'ie " xx set(units, 1)"
                If Cnt = 1 Then
                    pHighestPrecidencePhraseType = gdGetNum(pPhraseTypes, 1)
                Else
                    pHighestPrecidencePhraseType = 0
                End If
                
            Else
                'Multiple phrases exist with the first phrase being an
                'operator! This is an error!
                pHighestPrecidencePhraseType = 0
            End If
        End If
        pPhraseTypeArrayPos = 1
        pUserFormatting = gdGetStr(pPhraseFormattingArray, 1)
        
    Else
        pHighestPrecidencePhraseType = MaxPrecType
        pPhraseTypeArrayPos = MaxPrecPos
        pUserFormatting = gdGetStr(pPhraseFormattingArray, MaxPrecPos)
    End If

ErrExit:
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:DeterminePhraseType ", Err.Description
    End If
End Sub


'This routine driver the process of breaking down the tree (which is now
'filled at the top with conditions and Actions) down into sub phrases.
Public Sub BuildTree()
On Error GoTo ErrSection:
    
    Dim PType           As Byte
    Dim PhraseTypePos   As Long
    Dim aPhrs           As Long
    Dim aPhrTypes       As Long
    Dim aUserFor        As Long
    Dim Cnt             As Long
    Dim X               As Long
    Dim Y               As Long
    Dim QuitParsing     As Boolean
    Dim i               As Long
    Dim PFormatting     As String
    Dim LeftCnt         As Long
    Dim RightCnt        As Long
    
    'Break down phrases into subphrases...
    Do Until QuitParsing
        X = X + 1
        If X > mLast Then Exit Do
        
        'Determine highest presidence operator and it's array position...
        'If an operator isn't found then one phrase exists...
        DeterminePhraseType mPhraseTypesArray(X), mPhraseFormattingArray(X), _
            PType, PhraseTypePos, PFormatting
        
        'Debug.Print gdGetStr(aPhrase, X) & vbTab & PFormatting

        'Get user formatting info and move to tree
        If PFormatting <> "" Then
            SetFormatting X, PFormatting
        End If
        
        SetPhraseType X, PType
        
        'Look right and left of Operator.  Save count of phrases.
        LeftCnt = PhraseTypePos - 1
        RightCnt = gdGetSize(mPhraseTypesArray(X)) - 1 - PhraseTypePos
        
        Select Case PType
            
            'Operators expecting stuff to left of it (nothing right)
            Case gPT_Enter, gPT_Then
                If LeftCnt > 0 Then
                    SetNodeFrom X, mLast + 1
                    'SetNodeTo X, mLast + 1
                    WriteLeftSide PhraseTypePos, X
                End If
                
                'Error - stuff on right of Operator
                If RightCnt > 0 Then
                    SetErrorCode 11, X
                    'If NodeFrom(X) = 0 Then
                    '    SetNodeFrom X, mLast + 1
                    'End If
                    SetNodeTo X, mLast + 1
                    WriteRightSide PhraseTypePos, X
                End If
                
            'Operators expecting stuff to right of it (nothing left)
            Case gPT_If, gPT_Not, gPT_ElseIf, gPT_DoUntil
                
                'Error - Stuff found on Left side of Operator
                If LeftCnt > 0 Then
                    SetErrorCode 9, X
                    SetNodeFrom X, mLast + 1
                    'SetNodeTo X, mLast + 1
                    WriteLeftSide PhraseTypePos, X
                End If

                'Error - no stuff on right of Operator
                If RightCnt = 0 Then
                    SetErrorCode 8, X
                Else
                    'If NodeFrom(X) = 0 Then
                    '    SetNodeFrom X, mLast + 1
                    'End If
                    SetNodeTo X, mLast + 1
                    WriteRightSide PhraseTypePos, X
                End If
            
            'Operators which have nothing on left or right
            Case gPT_Else, gPT_EndIf, gPT_EndDo
            
                'Error - Stuff found on Left side of Operator
                If LeftCnt > 0 Then
                    SetErrorCode 10, X
                    'SetNodeFrom X, mLast + 1
                    SetNodeTo X, mLast + 1
                    WriteLeftSide PhraseTypePos, X
                Else
                    'Error - Stuff found on right of Operator
                    If RightCnt > 0 Then
                        SetErrorCode 10, X
                        'If NodeFrom(X) = 0 Then
                        '    SetNodeFrom X, mLast + 1
                        'End If
                        SetNodeTo X, mLast + 1
                        WriteRightSide PhraseTypePos, X
                    End If
                End If
            
            'Operators that have stuff to left and right of it
            Case gPT_And, gPT_Or, gPT_Of, gPT_GT, gPT_GE, gPT_LT, gPT_LE, gPT_NE, _
                 gPT_EQ, gPT_Mult, gPT_Div, gPT_Add, gPT_Sub, gPT_Offset
            
                'Error - nothing found on Left side of Operator
                If LeftCnt = 0 Then
                    SetErrorCode 4, X
                Else
                    SetNodeFrom X, mLast + 1
                    'SetNodeTo X, mLast + 1
                    WriteLeftSide PhraseTypePos, X
                End If
                
                'Error - nothing on right of operator
                If RightCnt = 0 Then
                    SetErrorCode 4, X
                Else
                    'If NodeFrom(X) = 0 Then
                    '    SetNodeFrom X, mLast + 1
                    'End If
                    SetNodeTo X, mLast + 1
                    WriteRightSide PhraseTypePos, X
                End If
                
            Case gPT_OuterParens
                
                'If one or more phrases exist between parens...
                Cnt = gdGetSize(mPhraseTypesArray(X)) - 1
                If Cnt - 2 > 0 Then
                    SetNodeFrom X, mLast + 1
                    SetNodeTo X, mLast + 1
                    
                    'Create array for storing phrases and phrase types of
                    'stuff in between outerparens
                    aPhrs = gdCreateArray(eGDARRAY_Strings, Cnt - 2, USE_DEFAULT_NULL)
                    aPhrTypes = gdCreateArray(eGDARRAY_TinyInts, Cnt - 2, USE_DEFAULT_NULL)
                    aUserFor = gdCreateArray(eGDARRAY_Longs, Cnt - 2, USE_DEFAULT_NULL)
                    mAddr.Add aPhrs
                    mAddr.Add aPhrTypes
                    mAddr.Add aUserFor
                    i = 0
                    For Y = 2 To Cnt - 1
                        i = i + 1
                        gdSetStr aPhrs, i, gdGetStr(mPhrasesArray(X), Y)
                        gdSetNum aPhrTypes, i, gdGetNum(mPhraseTypesArray(X), Y)
                        gdSetStr aUserFor, i, gdGetStr(mPhraseFormattingArray(X), Y)
                    Next Y
                    Add 0, 0, StrPhrase(aPhrs), aPhrs, aPhrTypes, aUserFor, "", _
                        0, 0, 0, X, 0, 0, mEntered(X)
                End If
                
            Case gPT_FCompiled, gPT_FCompiledAction, gPT_FInternal
                FCompiled mPhrasesArray(X), mPhraseTypesArray(X), _
                    mPhraseFormattingArray(X), X
                
            Case gPT_FTradeSense
                'When called from the editor, we are only interesed in
                'the TradeSense function and not it's codedtext.
                If mCalledFromEditor Then
                    FCompiled mPhrasesArray(X), mPhraseTypesArray(X), _
                        mPhraseFormattingArray(X), X
                Else
                    FTradeSense mPhrasesArray(X), mPhraseTypesArray(X), _
                        mPhraseFormattingArray(X), X
                End If
                      
            'Internal Data Arrays used in inputs of functions
            Case gPT_Trades, gPT_Bars, gPT_Portfolio, gPT_Systems
                SetReturnType X, GetReturnType(PType)
            
            'text, numbers, and parms handles here...
            Case Is <> 0
                SetReturnType X, GetReturnType(PType)
            
                'Special case: if parent is OF then this is a Bars label (GC, etc.)
                If mParent(X) > 0 Then
                    If mPhraseType(mParent(X)) = gPT_Of Then
                        SetPhraseType X, gPT_Bars
                        SetReturnType X, GetReturnType(gPT_Bars)
                    End If
                End If
                
            Case Else
                SetErrorCode 6, X
        End Select
        
    Loop
    
    UpdateReturnTypes
    'CheckFxnInputs
    SetNodeProcessingSeq
    BuildTreeCleanup
    
ErrExit:
    Exit Sub

ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:BuildTree ", Err.Description
    End If

End Sub

' Checks the input argument types against the required types for a function
Private Sub CheckFxnInputs()
    ' Exit if there is already an error
    If mErrorCode > 0 Then Exit Sub

    ' Loop thru the phrases
    Dim nPhrase As Long
    For nPhrase = mLast To 1 Step -1
        ' Is this a function?
        If mFID(nPhrase) > 0 Then
            ' Get the function definition
            Dim Func As cFunction
            Set Func = mFunctions.Item(CStr(mFID(nPhrase)))
            
            ' Loop thru the function nodes
            Dim nNode As Long, nNodeFrom As Long, nNodeTo As Long, nArg As Long
            nNodeFrom = mNodeFrom(nPhrase)
            nNodeTo = mNodeTo(nPhrase)
            If nNodeFrom = 0 Then nNodeFrom = nNodeTo
            If nNodeTo = 0 Then nNodeTo = nNodeFrom
            nArg = 1
            For nNode = nNodeFrom To nNodeTo
                ' Any nodes to process?
                If nNode = 0 Then Exit For
            
                ' Get the input argument type
                Dim nInputArgType As Byte
                nInputArgType = mReturnType(nNode)
                
                ' Get the argument type required by the function
                Dim nFxnArgType As Byte
                nFxnArgType = Func.Inputs.Item(nArg).ParmTypeID
                
                ' Do they match? (some coercions between data types are allowed)
                Dim bMatch As Boolean
                bMatch = False
                If nFxnArgType = nInputArgType Then
                    bMatch = True
                ElseIf nFxnArgType = gRetSeriesNbr And (nInputArgType = gRetConstantNbr Or nInputArgType = gRetSeriesNbr) Then
                    bMatch = True
                ElseIf nFxnArgType = gRetSeriesText And (nInputArgType = gRetConstantText Or nInputArgType = gRetSeriesText) Then
                    bMatch = True
                ElseIf nFxnArgType = gRetSeriesBoolean And (nInputArgType = gRetConstantBoolean Or nInputArgType = gRetSeriesBoolean Or nInputArgType = gRetConstantNbr) Then
                    bMatch = True
                ElseIf nFxnArgType = gRetConstantBoolean And (nInputArgType = gRetConstantBoolean Or nInputArgType = gRetConstantNbr) Then
                    bMatch = True
                End If
                                    
                ' Was there a mismatch?
                If Not bMatch Then
                    SetErrorCode 12, nNode
                    Exit For
                End If

                nArg = nArg + 1
            Next
        End If
        
        ' Stop checking if an error was found
        If mErrorCode > 0 Then Exit For
    Next
End Sub

'Write out all stuff to the left side of the operator to a new gdstring array
'and add a new node at the bottom of the tree pointing to these phrases.
Private Sub WriteLeftSide(pPTPos As Long, pNode As Long)
On Error GoTo ErrSection:
    Dim Y           As Long
    Dim aPhrs       As Long
    Dim aPhrTypes   As Long
    Dim aUserFor    As Long
    
    aPhrs = gdCreateArray(eGDARRAY_Strings, pPTPos - 1, USE_DEFAULT_NULL)
    aPhrTypes = gdCreateArray(eGDARRAY_TinyInts, pPTPos - 1, USE_DEFAULT_NULL)
    aUserFor = gdCreateArray(eGDARRAY_Strings, pPTPos - 1, USE_DEFAULT_NULL)
    mAddr.Add aPhrs
    mAddr.Add aPhrTypes
    mAddr.Add aUserFor
    For Y = 1 To pPTPos - 1
        gdSetStr aPhrs, Y, gdGetStr(mPhrasesArray(pNode), Y)
        gdSetNum aPhrTypes, Y, gdGetNum(mPhraseTypesArray(pNode), Y)
        gdSetStr aUserFor, Y, gdGetStr(mPhraseFormattingArray(pNode), Y)
    Next Y
    Add 0, 0, StrPhrase(aPhrs), aPhrs, aPhrTypes, aUserFor, "", _
        0, 0, 0, pNode, 0, 0, mEntered(pNode)
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:WriteLeftSide ", Err.Description
    End If
End Sub

'Write out all stuff to the right side of the operator to a new gdstring array
'and add a new node at the bottom of the tree pointing to these phrases.
Private Sub WriteRightSide(pPTPos As Long, pNode As Long)
On Error GoTo ErrSection:
    Dim Y           As Long
    Dim Cnt         As Long
    Dim i           As Long
    Dim aPhrs       As Long
    Dim aPhrTypes   As Long
    Dim aUserFor    As Long
    
    Cnt = gdGetSize(mPhraseTypesArray(pNode)) - 1
    aPhrs = gdCreateArray(eGDARRAY_Strings, Cnt - pPTPos, USE_DEFAULT_NULL)
    aPhrTypes = gdCreateArray(eGDARRAY_TinyInts, Cnt - pPTPos, USE_DEFAULT_NULL)
    aUserFor = gdCreateArray(eGDARRAY_Strings, Cnt - pPTPos, USE_DEFAULT_NULL)
    mAddr.Add aPhrs
    mAddr.Add aPhrTypes
    mAddr.Add aUserFor
    i = 0
    For Y = pPTPos + 1 To Cnt
        i = i + 1
        gdSetStr aPhrs, i, gdGetStr(mPhrasesArray(pNode), Y)
        gdSetNum aPhrTypes, i, gdGetNum(mPhraseTypesArray(pNode), Y)
        gdSetStr aUserFor, i, gdGetStr(mPhraseFormattingArray(pNode), Y)
    Next Y
    Add 0, 0, StrPhrase(aPhrs), aPhrs, aPhrTypes, aUserFor, "", _
        0, 0, 0, pNode, 0, 0, mEntered(pNode)
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:WriteRightSide ", Err.Description
    End If
End Sub

'Update the return types of all nodes.  Up to this point only functions
'numbers, parms and text values have return types assigned.  Set the
'return type for all operators (which will be zero)
Private Sub UpdateReturnTypes()
On Error GoTo ErrSection:
    Dim X           As Long
    Dim RTLeft      As Byte
    Dim RTRight     As Byte
    Dim Pass        As Byte
    
    If mErrorCode > 0 Then Exit Sub
    
    'Build 2 dimensional array for defaulting return types based on the
    'left and right sides of the operator.  NOTE: When boolean and Number
    'are used together the default is Number.  This is adjusted later in
    'this routine once the operator is known (i.e. comparison operator
    'switch to boolean).
    '
    '            Const#    ConstTxt  SeriesBol Series# SeriesTxt ConstBol
    '            -------------------------------------------------------
    ' Const#     Const#    Err       SeriesBol Series# Err       Const#
    ' ConstTxt   Err       ConstTxt  Err       Err     SeriesTxt Err
    ' SeriesBol  SeriesBol Err       SeriesBol Series# Err       SeriesBol
    ' Series#    Series#   Err       Series#   Series# Err       Series#
    ' SeriesTxt  Err       SeriesTxt Err       Err     SeriesTxt Err
    ' ConstBol   Const#    Err       SeriesBol Series# Err       ConstBol
    
    Dim RetTypes(14, 14)      As Byte
    RetTypes(gRetConstantNbr, gRetConstantNbr) = gRetConstantNbr
    RetTypes(gRetConstantNbr, gRetConstantText) = 0
    RetTypes(gRetConstantNbr, gRetSeriesBoolean) = gRetSeriesBoolean
    RetTypes(gRetConstantNbr, gRetSeriesNbr) = gRetSeriesNbr
    RetTypes(gRetConstantNbr, gRetSeriesText) = 0
    RetTypes(gRetConstantNbr, gRetConstantBoolean) = gRetConstantNbr
    
    RetTypes(gRetConstantText, gRetConstantNbr) = 0
    RetTypes(gRetConstantText, gRetConstantText) = gRetConstantText
    RetTypes(gRetConstantText, gRetSeriesBoolean) = 0
    RetTypes(gRetConstantText, gRetSeriesNbr) = 0
    RetTypes(gRetConstantText, gRetSeriesText) = gRetSeriesText
    RetTypes(gRetConstantText, gRetConstantBoolean) = 0
    
    RetTypes(gRetSeriesBoolean, gRetConstantNbr) = gRetSeriesBoolean
    RetTypes(gRetSeriesBoolean, gRetConstantText) = 0
    RetTypes(gRetSeriesBoolean, gRetSeriesBoolean) = gRetSeriesBoolean
    RetTypes(gRetSeriesBoolean, gRetSeriesNbr) = gRetSeriesNbr
    RetTypes(gRetSeriesBoolean, gRetSeriesText) = 0
    RetTypes(gRetSeriesBoolean, gRetConstantBoolean) = gRetSeriesBoolean
    
    RetTypes(gRetSeriesNbr, gRetConstantNbr) = gRetSeriesNbr
    RetTypes(gRetSeriesNbr, gRetConstantText) = 0
    RetTypes(gRetSeriesNbr, gRetSeriesBoolean) = gRetSeriesNbr
    RetTypes(gRetSeriesNbr, gRetSeriesNbr) = gRetSeriesNbr
    RetTypes(gRetSeriesNbr, gRetSeriesText) = 0
    RetTypes(gRetSeriesNbr, gRetConstantBoolean) = gRetSeriesNbr
    
    RetTypes(gRetSeriesText, gRetConstantNbr) = 0
    RetTypes(gRetSeriesText, gRetConstantText) = gRetSeriesText
    RetTypes(gRetSeriesText, gRetSeriesBoolean) = 0
    RetTypes(gRetSeriesText, gRetSeriesNbr) = 0
    RetTypes(gRetSeriesText, gRetSeriesText) = gRetSeriesText
    RetTypes(gRetSeriesText, gRetConstantBoolean) = 0
    
    RetTypes(gRetConstantBoolean, gRetConstantNbr) = gRetConstantNbr
    RetTypes(gRetConstantBoolean, gRetConstantText) = 0
    RetTypes(gRetConstantBoolean, gRetSeriesBoolean) = gRetSeriesBoolean
    RetTypes(gRetConstantBoolean, gRetSeriesNbr) = gRetSeriesNbr
    RetTypes(gRetConstantBoolean, gRetSeriesText) = 0
    RetTypes(gRetConstantBoolean, gRetConstantBoolean) = gRetConstantBoolean
    
    'Make two passes to ensure that input return types are assigned correc-
    'tly.  If the left node of an operator is not assigned yet, and the right
    'node is an input.  The input return type cannot be assigned.  Making 2
    'passes ensures the inputs are all assign correctly.
    For Pass = 1 To 2
        For X = mLast To 1 Step -1
            
            'Convert parm return types to be consistent to what they are
            'being compared to on the other side of the operator
            AdjustParmReturnTypes X
            
            If mReturnType(X) = 0 And mError(X) = 0 Then
                
                Select Case mPhraseType(X)
                    Case gPT_OuterParens, gPT_Offset
                        If mNodeFrom(X) <> 0 Then
                            SetReturnType X, mReturnType(mNodeFrom(X))
                        End If
                        
                    'Math operators
                    Case gPT_Mult, gPT_Div, gPT_Add, gPT_Sub
                        RTLeft = mReturnType(mNodeFrom(X))
                        RTRight = mReturnType(mNodeTo(X))
                        
                        If RTLeft <> 0 And RTRight <> 0 Then
                            If RetTypes(RTLeft, RTRight) = 0 Then
                                SetErrorCode 1, X
                            Else
                                SetReturnType X, RetTypes(RTLeft, RTRight)
                            End If
                        End If
                        
                    'Override series of nbr defaults to boolean with
                    'comparison operators
                    Case gPT_And, gPT_Or, gPT_GT, gPT_GE, gPT_LT, gPT_LE, gPT_NE, _
                         gPT_EQ
                        RTLeft = mReturnType(mNodeFrom(X))
                        RTRight = mReturnType(mNodeTo(X))
                        
                        If RTLeft <> 0 And RTRight <> 0 Then
                            If RetTypes(RTLeft, RTRight) = 0 Then
                                SetErrorCode 1, X
                            Else
                                SetReturnType X, RetTypes(RTLeft, RTRight)
                            End If
                        
                            'Override because of operator used.
                            If RetTypes(RTLeft, RTRight) = gRetSeriesNbr Or _
                               RetTypes(RTLeft, RTRight) = gRetSeriesText Then
                                SetReturnType X, gRetSeriesBoolean
                            End If
                            If RetTypes(RTLeft, RTRight) = gRetConstantNbr Or _
                               RetTypes(RTLeft, RTRight) = gRetConstantText Then
                                SetReturnType X, gRetConstantBoolean
                            End If
                        End If
                        
                    Case gPT_If, gPT_Not, gPT_ElseIf, gPT_DoUntil
                        SetReturnType X, mReturnType(mNodeTo(X))
                    
                    Case gPT_Of, gPT_Then
                        SetReturnType X, mReturnType(mNodeFrom(X))
                        
                End Select
                
            End If
        Next X
        
    Next Pass
         
    'STAND ALONE INPUT (NOT PART OF FUNCTION)
    'Set it to the return value on the opposite side of the operator (The
    'constant version of the return type)
    Dim Y    As Long
    Dim RT   As Byte
    For X = 1 To mLast
        If mPhraseType(X) = gPT_Parm Then
            If mParent(X) > 0 Then
                If mFID(mParent(X)) <= 0 And mNodeFrom(mParent(X)) > 0 Then
                    
                    'If the parent phrase type is an offset, then the input
                    'is always a constant numeric value.
                    If mPhraseType(mParent(X)) = gPT_Offset Then
                        RT = gRetConstantNbr
                    Else
                        'Now, set to the constant version of this return
                        'type (number, boolean, text)
                        RT = mReturnType(mNodeFrom(mParent(X)))
                        Select Case RT
                            Case gRetSeriesNbr: RT = gRetConstantNbr
                            Case gRetSeriesBoolean: RT = gRetConstantBoolean
                            Case gRetSeriesText: RT = gRetConstantText
                        End Select
                    End If
                    gdSetNum aReturnType, X, RT
                End If
            End If
        End If
    Next X
    
    'Look for duplicate parms (and not an input to a function)...
    'If dups found and the dup is an input to a function, change this
    'input outside of the function to the return type of the function input.
    For X = 1 To mLast
        If mPhraseType(X) = gPT_Parm Then
            For Y = 1 To mLast
                If X <> Y And mPhrase(X) = mPhrase(Y) And _
                   mPhraseType(Y) = gPT_Parm And mFID(mParent(Y)) > 0 Then
                    gdSetNum aReturnType, X, mReturnType(Y)
                    Exit For
                End If
            Next Y
        End If
    Next X
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:UpdateReturnTypes ", Err.Description
    End If
End Sub

'Assign execution sequence to the IF/ELSEIF/ELSE/ENDIF operators
'Each of these operators has 1 of 2 paths:  True path, false path.  They
'are stored in aExeTrue and aExeFalse gdarray's respectively.
Private Sub SetNodeProcessingSeq()
On Error GoTo ErrSection:
    Dim X           As Long
    Dim Y           As Long
    Dim O           As Long
    Dim NextSeq     As Long
    
    For X = 1 To mLast
    
        'Quit when the parent phrase type is not <Enter>
        If Parent(X) <> 0 Then
            If PhraseType(Parent(X)) <> gPT_Enter Then Exit For
        End If
        
        Select Case PhraseType(X)
            Case gPT_If, gPT_ElseIf
                SetExeTrue X, X + 1
                O = 1
                'Search for matching Else, ElseIF, or Endif (which ever is 1st)
                For Y = X + 1 To mLast
                    Select Case PhraseType(Y)
                        Case gPT_If: O = O + 1
                        Case gPT_Else, gPT_ElseIf
                            If O = 1 Then Exit For
                        Case gPT_EndIf
                            If O = 1 Then
                                Exit For
                            Else
                                O = O - 1
                            End If
                    End Select
                Next Y
                SetExeFalse X, Y
            
            Case gPT_DoUntil
                SetExeFalse X, X + 1
                O = 1
                'This section skips past imbedded DoUntil blocks...
                For Y = X + 1 To mLast
                    Select Case PhraseType(Y)
                        Case gPT_DoUntil: O = O + 1
                        Case gPT_EndDo: O = O - 1
                    End Select
                    If O = 0 Then Exit For
                Next Y
                SetExeTrue X, Y + 1
            
            'Special case: set the exefalse side of ELSE to the EndIf.  The
            'Compiled actions use the false side for their True execution
            'when the next operator after compiled actions are ELSE.
            Case gPT_Else
                SetExeTrue X, X + 1
                O = 1
                'This section skips past imbedded IF blocks...
                For Y = X + 1 To mLast
                    Select Case PhraseType(Y)
                        Case gPT_If: O = O + 1
                        Case gPT_Else, gPT_ElseIf: O = O - 1
                        Case gPT_EndIf: O = O - 1
                    End Select
                    If O = 0 Then Exit For
                Next Y
                SetExeFalse X, Y
            
            Case gPT_EndIf
                'if the next statement after is an Else. IE...
                '  IF X=1 THEN
                '     IF Y>0 THEN
                '         set (X,1)
                '     ENDIF
                '  ELSE     <---------follows Endif
                '     SET (Y,1)
                '  ENDIF    <---------should point here
                If PhraseType(X + 1) = gPT_Else Or _
                   PhraseType(X + 1) = gPT_ElseIf Then
                    O = 1
                    'Find EndIf of the Else and use it's location
                    For Y = X + 1 To mLast
                        Select Case PhraseType(Y)
                            Case gPT_If: O = O + 1
                            Case gPT_EndIf: O = O - 1
                        End Select
                        If O = 0 Then Exit For
                    Next Y
                    SetExeTrue X, Y
                Else
                    SetExeTrue X, X + 1
                End If
                SetExeFalse X, 0
                
            'Branch true back to matching DoUntil operator
            Case gPT_EndDo
                O = -1
                For Y = X - 1 To 1 Step -1
                    Select Case PhraseType(Y)
                        Case gPT_DoUntil: O = O + 1
                        Case gPT_EndDo: O = O - 1
                    End Select
                    If O = 0 Then Exit For
                Next Y
                SetExeTrue X, Y
                SetExeFalse X, 0
            
        End Select
    Next X
    
    'Assign execution sequence to the IF/ELSEIF/ELSE/ENDIF operators
    'Each of these operators has 1 of 2 paths:  True path, false path.  They
    'are stored in aExeTrue and aExeFalse gdarray's respectively.
    For X = 1 To mLast
    
        'Quit when the parent phrase type is not <Enter>
        If Parent(X) <> 0 Then
            If PhraseType(Parent(X)) <> gPT_Enter Then Exit For
        End If
        
        'If the next statement to execute is an Else, or ElseIF, then
        'search for the nearest Endif
        If PhraseType(X) = gPT_FCompiledAction Then
            SetExeFalse X, 0
            If PhraseType(X + 1) = gPT_Else Or _
               PhraseType(X + 1) = gPT_ElseIf Then
                For Y = X + 1 To mLast
                    If PhraseType(Y) = gPT_EndIf Then
                        SetExeTrue X, Y
                        Exit For
                    End If
                Next Y
                'SetExeTrue X, ExeFalse(X + 1)
            Else
                SetExeTrue X, X + 1
            End If
        End If
        
    Next X
    
    'Assign unique numbers by statement.  ie.  Statement 1 is assigned to all
    'nodes down through the tree starting at node 1.  When finished, saved the
    'node number of the bottom node in the statement.  This is used by the
    'evaluator to update nodes for immediate and late calculating.  The
    'tree is read from bottom up starting from the node values stored in
    'mStatementBottom.
    For X = 1 To mLast
        If mPhraseType(X) = gPT_Enter Then
            NextSeq = NextSeq + 1
            mMaxNode = -999999
            SetExecutionSequence X, NextSeq
            mStatementBottom.Add mMaxNode, X
        Else
            Exit For
        End If
    Next X
        
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:SetNodeProcessingSeq ", Err.Description
    End If
End Sub

'This routine assigns a unqiue value to each node by rule.  For example, all phrase
'nodes in Rule 1 get 1, all phrase nodes in rule 2 get 2, etc.  When phrases are
'evaluated in cEvaluator, nodes must be executed by sequence because users put rules
'in a specified sequence.
Private Sub SetExecutionSequence(pNode As Long, pNextSeq As Long)
On Error GoTo ErrSection:
    Dim Y           As Long
    
    SetRuleID pNode, pNextSeq
    If pNode > mMaxNode Then
        mMaxNode = pNode
    End If
    
    Select Case PhraseType(pNode)
        Case gPT_FCompiled, gPT_FCompiledAction, _
             gPT_FTradeSense, gPT_FInternal
            
            If NodeFrom(pNode) > 0 Then
                For Y = NodeFrom(pNode) To NodeTo(pNode)
                    SetExecutionSequence Y, pNextSeq
                Next Y
            End If
            
        Case Else
            If NodeFrom(pNode) > 0 Then
                SetExecutionSequence NodeFrom(pNode), pNextSeq
            End If
            If NodeTo(pNode) > 0 Then
                SetExecutionSequence NodeTo(pNode), pNextSeq
            End If
            
    End Select
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:SetExecutionSequence ", Err.Description
    End If
End Sub

Private Sub AdjustParmReturnTypes(pX As Long)
On Error GoTo ErrSection:
    Dim X               As Long
    Dim Parent          As Long
    Dim wReturnType     As Long
    
    If mPhraseType(pX) = gPT_Parm Then
        
        Parent = gdGetNum(aParent, pX)
        If Parent > 0 Then
            'If parm is to a function input then set the return type to that
            'of the function input
            If mInputID(pX) > 0 Then
                SetReturnType pX, mFunctions.Item(CStr(mFID(Parent))).Inputs.Item(CStr(mInputID(pX))).ParmTypeID
                GoTo ErrExit
            End If
                    
            Select Case mPhraseType(Parent)
                    
                Case gPT_Mult, gPT_Div, gPT_Add, gPT_Sub, _
                     gPT_And, gPT_Or, gPT_GT, gPT_GE, gPT_LT, gPT_LE, gPT_NE, _
                     gPT_EQ
                    
                    'Check the opposite of the operator.  Set to parm return
                    'type to the series version of the return type.
                    If mPhraseType(mNodeFrom(Parent)) <> gPT_Parm Then
                        wReturnType = mReturnType(mNodeFrom(Parent))
                    Else
                        wReturnType = mReturnType(mNodeTo(Parent))
                    End If
                    
                    If wReturnType = gRetSeriesNbr Or wReturnType = gRetConstantNbr Then
                        SetReturnType pX, gRetSeriesNbr
                    ElseIf wReturnType = gRetSeriesBoolean Or wReturnType = gRetConstantBoolean Then
                        SetReturnType pX, gRetSeriesBoolean
                    ElseIf wReturnType = gRetSeriesText Or wReturnType = gRetConstantText Then
                        SetReturnType pX, gRetSeriesText
                    Else
                        'Could not be determine yet...
                        SetReturnType pX, 0
                    End If
            
            End Select
        End If
        
    End If
    
ErrExit:
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:AdjustParmReturnTypes ", Err.Description
    End If
End Sub

Private Sub BuildTreeCleanup()
On Error GoTo ErrSection:
    
    'Free up any extra memory at end of arrays...
    gdFreeExtra aPhraseType
    gdFreeExtra aReturnType
    gdFreeExtra aPhrase
    gdFreeExtra aFormatting
    gdFreeExtra aClassName
    gdFreeExtra aClassPtr
    gdFreeExtra aLate
    gdFreeExtra aSize
    gdFreeExtra aNodeFrom
    gdFreeExtra aNodeTo
    gdFreeExtra aParent
    gdFreeExtra aRuleID
    gdFreeExtra aExeTrue
    gdFreeExtra aExeFalse
    gdFreeExtra aPoolAddr
    gdFreeExtra aParmList
    gdFreeExtra aRebuildArray
    gdFreeExtra aFID
    gdFreeExtra aInputID
    gdFreeExtra aReq
    gdFreeExtra aEntered
    gdFreeExtra aError
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:BuildTreeCleanup ", Err.Description
    End If
End Sub

Private Function GetReturnType(pPhraseType As Byte) As Byte
    Select Case pPhraseType
        Case gPT_And, gPT_Or, gPT_GT, gPT_GE, gPT_LT, gPT_LE, gPT_NE, gPT_EQ
            GetReturnType = gRetSeriesBoolean
        Case gPT_Mult, gPT_Div, gPT_Add, gPT_Sub, gPT_Offset
            GetReturnType = gRetSeriesNbr
        Case gPT_Nbr
            GetReturnType = gRetConstantNbr
        Case gPT_Parm
            GetReturnType = gRetSeriesNbr
        Case gPT_Text, gPT_BracketComment
            GetReturnType = gRetConstantText
        Case gPT_Trades
            GetReturnType = gRetTrades
        Case gPT_Portfolio
            GetReturnType = gRetPortfolio
        Case gPT_Systems
            GetReturnType = gRetSystems
        Case gPT_Bars
            GetReturnType = gRetBars
    End Select
End Function

Private Sub FTradeSense(ByVal pPhrases As Long, _
    ByVal pPhraseTypes As Long, pPhraseFormatting As Long, pCurrRec As Long)
On Error GoTo ErrSection:
    Dim InputPhrases            As cGdArray
    Dim InputPhraseTypes        As cGdArray
    Dim InputFormatting         As cGdArray
    Dim CTPhrases               As Long
    Dim CTPhraseTypes           As Long
    Dim CTFormatting            As Long
    Dim AdjPhrases              As Long
    Dim AdjPhraseTypes          As Long
    Dim AdjFormatting           As Long
    Dim ICnt                    As Long
    Dim FunctionID              As Long
    Dim X                       As Long
    Dim curInput                As Long
    Dim curFName                As String
    Dim Y                       As Long
    Dim P                       As Long
    Dim F                       As cFunction
    Dim ErrorCode                As Byte
    Dim FICnt                   As Long
    Dim curPhrase               As String
    
    'Function name will be stored in element 1 if called from Editor.  Other
    'wise the Function ID will be stored (codedtext functions).
    curPhrase = gdGetStr(pPhrases, 1)
    If Not IsAlpha(curPhrase) And Val(curPhrase) <> 0 Then
        FunctionID = Val(curPhrase)
        Set F = mFunctions.Item(CStr(FunctionID))
    Else
        curFName = curPhrase
        If mFunctions.NameFound(curFName, FunctionID) Then
            Set F = mFunctions.Item(CStr(FunctionID))
        Else
            Err.Raise gUserErr, , "Function: " & curPhrase & " was not found."
        End If
    End If
    
    'Retrieve Function Inputs...
    FICnt = F.Inputs.Count
    
    'Update settings of this "unbroken down" Codedtext Function...
    SetNodeFrom pCurrRec, mLast + 1
    SetNodeTo pCurrRec, mLast + 1
    SetPhrase pCurrRec, StrPhrase(pPhrases)
    SetFID pCurrRec, F.FunctionID
    SetReturnType pCurrRec, F.ReturnTypeID
    SetEntered pCurrRec, mEntered(pCurrRec)
    
    'Store phrases in between the parenthesis of the function into phrase
    'arrays (gdaRrays).  these are pointed at by the vectors: InputPhrases
    'and InputPhraseTypes.
    'Example:  "Avg Win (10)"
    ' - 10 is stored in gdArrays
    GetFunctionInputs pCurrRec, pPhrases, pPhraseTypes, pPhraseFormatting, _
            InputPhrases, InputPhraseTypes, InputFormatting, ICnt
    
    'Retrieve codedtext for TradeSense function.  Convert into phrase arrays
    ConvertCT F.CodedText, CTPhrases, CTPhraseTypes, CTFormatting
    
    'Verify inputs entered match in total with TradeSense function inputs...
    If ICnt > FICnt Then
        SetErrorCode 3, pCurrRec
    End If
    
    'Continue if inputs exist
    If ICnt > 0 Or FICnt > 0 Then
    
        'Now search through the codedtext of the Function for parms.  Replace
        'each one with the saved input arrays in the original function (
        'pointed to by InputPhrases and InputPhraseTypes addresses).
        AdjPhrases = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
        AdjPhraseTypes = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
        AdjFormatting = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
        mAddr.Add AdjPhrases
        mAddr.Add AdjPhraseTypes
        mAddr.Add AdjFormatting

        '========================== 4/3/2001
        'Everything is positional.  For example,
        '
        '   Rule: IF Win Pct (5)...
        '   TradeSense Coded Text:
        '       Occurrence (Profit>0, Trades Back,1)/Trades Back
        '
        'In the example, "5" must be inserted in place of "Trades Back", which occurs
        '2 times.  5 positionally alligns with "Trades Back", but not with the 2nd
        'occurrence of it.  Therefore, the 2nd occurrence is changed in the CT arrays
        'right after it is updated in the 1st occurrence so it's not found as an input
        'later...
        Dim CTParms     As cGdArray
        Dim ParmName    As String
        Dim Fnd         As Boolean
        Dim Z           As Long
        Set CTParms = New cGdArray
        CTParms.Create eGDARRAY_Strings
        For X = 1 To gdGetSize(CTPhraseTypes) - 1
            If gdGetNum(CTPhraseTypes, X) = gPT_Parm Then
            
                'Get position of this parm so it can be looked
                'up in InputPhrases...
                ParmName = gdGetStr(CTPhrases, X)
                
                'Lookup parm in CTParms array...
                Fnd = False
                For Y = 1 To CTParms.Size - 1
                    If CTParms.Item(Y) = ParmName Then
                        Fnd = True
                        Exit For
                    End If
                Next Y
                
                If Not Fnd Then
                    CTParms.Add ParmName
                    Y = CTParms.Size - 1
                End If
                
                'This looping accounts for expressions entered for inputs...
                For Z = 1 To gdGetSize(InputPhrases.Item(Y)) - 1
                    P = P + 1
                    gdSetStr AdjPhrases, P, gdGetStr(InputPhrases.Item(Y), Z)
                    gdSetNum AdjPhraseTypes, P, gdGetNum(InputPhraseTypes.Item(Y), Z)
                    gdSetStr AdjFormatting, P, gdGetStr(InputFormatting.Item(Y), Z)
                Next Z
            Else
                P = P + 1
                gdSetStr AdjPhrases, P, gdGetStr(CTPhrases, X)
                gdSetNum AdjPhraseTypes, P, gdGetNum(CTPhraseTypes, X)
                gdSetStr AdjFormatting, P, gdGetStr(CTFormatting, X)
            End If
        Next X
        '========================== 4/3/2001
        
        'Add new node with tradeSense coded text with parm phrases overlaid.
        Add 0, 0, StrPhrase(AdjPhrases), AdjPhrases, _
            AdjPhraseTypes, AdjFormatting, "", 0, 0, 0, pCurrRec, 0, 0, 0
    Else
        'No inputs.  Post the Codedtext from sub function
        Add 0, 0, StrPhrase(CTPhrases), CTPhrases, _
            CTPhraseTypes, CTFormatting, "", 0, 0, 0, pCurrRec, 0, 0, 0
    End If
      
    'Cleanup
    Set F = Nothing
    Set InputPhrases = Nothing
    Set InputPhraseTypes = Nothing
    Set InputFormatting = Nothing
    Set CTParms = Nothing
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:FTradeSense ", Err.Description
    End If
End Sub


'This routine updates the properties of compiled functions for the current
'node in the tree, and then creates new nodes at the end of the tree for each
'input in the function.  ALL possible inputs are created as nodes, even if
'they were not all entered in the editor.  Default values fill in inputs if
'not entered in editor.
Private Sub FCompiled(ByVal pPhrases As Long, ByVal pPhraseTypes As Long, _
        pPhraseFormatting As Long, pCurrRec As Long)
On Error GoTo ErrSection:
    
    Dim InputPhrases            As cGdArray
    Dim InputPhraseTypes        As cGdArray
    Dim InputFormatting         As cGdArray
    Dim Phr                     As Long
    Dim PhrTypes                As Long
    Dim PhrFormatting           As Long
    Dim EnteredInputs           As Long
    Dim X                       As Long
    Dim Y                       As Long
    Dim curInput                As Long
    Dim EnterableInputs         As Integer
    Dim InternalDataInputs      As Long
    Dim curInputID              As Long
    Dim curInputName            As String
    Dim curReturnTypeID         As Byte
    Dim curDefaultValue         As String
    Dim curReq                  As Byte
    Dim F                       As cFunction
    Dim FunctionID              As Long
    Dim curFName                As String
    Dim FICnt                   As Long
    Dim InputsTotal             As Long
    Dim curPhrase               As String
    
    'Function name will be stored in element 1 if called from Editor.  Other
    'wise the Function ID will be stored (codedtext functions).
    curPhrase = gdGetStr(pPhrases, 1)
    If Not IsAlpha(curPhrase) And Val(curPhrase) <> 0 Then
        FunctionID = Val(curPhrase)
        Set F = mFunctions.Item(CStr(FunctionID))
    Else
        curFName = curPhrase
        If mFunctions.NameFound(curFName, FunctionID) Then
            Set F = mFunctions.Item(CStr(FunctionID))
        Else
            Err.Raise gUserErr, , "Function: " & curPhrase & " was not found."
        End If
    End If
    
    'Retrieve Function Inputs...
    FICnt = F.Inputs.Count
    
    'Update function properties...
    SetClassName pCurrRec, F.DLLName & "." & F.ClassName
    SetLate pCurrRec, Abs(F.LateCalculating)
    SetPhrase pCurrRec, StrPhrase(pPhrases)
    SetFID pCurrRec, F.FunctionID
    SetReturnType pCurrRec, F.ReturnTypeID
    SetEntered pCurrRec, 1
    
    'Get inputs entered by user or saved in coded text (will not have inter
    'nal data inputs)...
    GetFunctionInputs pCurrRec, pPhrases, pPhraseTypes, pPhraseFormatting, _
            InputPhrases, InputPhraseTypes, InputFormatting, EnteredInputs
    
    'Flag function as an error if too many inputs are entered...
    If EnteredInputs > FICnt Then
        SetErrorCode 3, pCurrRec
    End If
    
    'Determine the maximum number of inputs.  Either the inputs from the
    'table (including internal), or the number entered (if called from
    'editor).
    If FICnt > EnteredInputs Then
        InputsTotal = FICnt
    Else
        If EnteredInputs > 0 Then
            InputsTotal = EnteredInputs
        End If
    End If
    
    'Set subphrease Beg/End values based on the number of inputs...
    If InputsTotal > 0 Then
        SetNodeFrom pCurrRec, mLast + 1
        SetNodeTo pCurrRec, mLast + InputsTotal
    Else
        GoTo ErrExit:
    End If
    
    'Determine the first input after Internal data structures.  This will
    'only applied if called from editor.  Otherwise, the first is 1.
    For Y = 1 To FICnt
        If F.Inputs.Item(Y).ParmTypeUsage <> gInternalData Then
            EnterableInputs = EnterableInputs + 1
        End If
    Next Y
    InternalDataInputs = FICnt - EnterableInputs

    'Add new nodes for each input...
    For X = 1 To InputsTotal
        
        'Get default input data from table...
        If X <= FICnt Then
            With F.Inputs.Item(X)
                curInputID = .ParmID
                curInputName = .ParmName
                curDefaultValue = .DefaultValue
                curReturnTypeID = .ParmTypeID
                curReq = .Required * -1
            End With
        End If
        
        'Internal Data Structure...
        If X <= InternalDataInputs Then
            Phr = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
            PhrTypes = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
            PhrFormatting = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
            mAddr.Add Phr
            mAddr.Add PhrTypes
            mAddr.Add PhrFormatting
            
            gdSetStr Phr, 1, curInputName
            Select Case curReturnTypeID
                Case gRetTrades: gdSetNum PhrTypes, 1, gPT_Trades
                Case gRetBars: gdSetNum PhrTypes, 1, gPT_Bars
                Case gRetPortfolio: gdSetNum PhrTypes, 1, gPT_Portfolio
                Case gRetSystems: gdSetNum PhrTypes, 1, gPT_Systems
            End Select
            Add 0, 0, curInputName, Phr, PhrTypes, PhrFormatting, "", 0, _
                0, 0, pCurrRec, 0, 0, 0
        Else
            'Nothing entered for input, look for default value from
            'input in the table.  Account for expression in default of inputs (future).
            ''If gdGetSize(InputPhraseTypes.Item(X - InternalDataInputs)) - 1 <= 0 Then
            If X > (InternalDataInputs + EnteredInputs) Then
                Phr = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
                PhrTypes = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
                PhrFormatting = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
                mAddr.Add Phr
                mAddr.Add PhrTypes
                mAddr.Add PhrFormatting
                
                'OPTIONAL INPUT: fill in the first element of the
                'arrays with the default value (alway a number or text value
                'in this vers).
                If curReq = 0 Then
                    gdSetStr Phr, 1, curDefaultValue
                    If IsNumeric(curDefaultValue) Then
                        gdSetNum PhrTypes, 1, gPT_Nbr
                    Else
                        gdSetNum PhrTypes, 1, gPT_Text
                    End If
                Else
                    SetErrorCode 7, pCurrRec
                End If
                
                Add 0, curReturnTypeID, curDefaultValue, Phr, PhrTypes, PhrFormatting, "", 0, _
                    0, 0, pCurrRec, 0, curReq, 0
            Else
                'Add the input entered as new node...
                'Don't set the return type, let BuildTree determine what it is.
                'This is to ensure that type checking is done correctly in
                'cExpression.
                If mCalledFromEditor Then
                    Add 0, 0, StrPhrase(InputPhrases.Item(X - InternalDataInputs)), _
                        InputPhrases.Item(X - InternalDataInputs), _
                        InputPhraseTypes.Item(X - InternalDataInputs), _
                        InputFormatting.Item(X - InternalDataInputs), _
                        "", 0, 0, 0, pCurrRec, 0, curReq, mEntered(pCurrRec)
                Else
                    Add 0, 0, StrPhrase(InputPhrases.Item(X)), _
                        InputPhrases.Item(X), _
                        InputPhraseTypes.Item(X), _
                        InputFormatting.Item(X), _
                        "", 0, 0, 0, pCurrRec, 0, curReq, mEntered(pCurrRec)
                End If
            End If
            
            If X <= FICnt Then
                SetInputID mLast, curInputID  'Save InputID...
                'SetFID mLast, F.FunctionID    'Save FunctionID...
            End If
            
        End If
        
    Next X
    
ErrExit:
    Set F = Nothing
    Exit Sub

ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:FCompiled ", Err.Description
    End If

End Sub

'Strip out and save each input of a function as an array of inputs.  If no
'input existing between commas, then an array of 1 element is returned with
'a blank value for the phrase and 0 for the phrase type.
Private Sub GetFunctionInputs(pCurArrayElement As Long, pPhrases As Long, _
    pPhraseTypes As Long, pPhraseFormatting As Long, _
    pInputPhrases As cGdArray, pInputPhraseTypes As cGdArray, _
    pInputPhraseFormatting As cGdArray, pEnteredInputs As Long)
On Error GoTo ErrSection:
    
    Dim Fnd             As Boolean
    Dim n               As Long
    Dim X               As Long
    Dim curType         As Byte
    Dim curPhrase       As String
    Dim curFormatting   As String
    Dim Parens          As Integer
    Dim lLastPhrase            As Long
    Dim curPhrases      As Long
    Dim curPhraseTypes  As Long
    Dim curPhraseFormatting As Long
    Dim TotPhrases      As Long
    
    ' Store address of phrase arrays
    Set pInputPhrases = New cGdArray
    Set pInputPhraseTypes = New cGdArray
    Set pInputPhraseFormatting = New cGdArray
    pInputPhrases.Create eGDARRAY_Longs
    pInputPhraseTypes.Create eGDARRAY_Longs
    pInputPhraseFormatting.Create eGDARRAY_Longs
    TotPhrases = gdGetSize(pPhrases) - 1
    
    'gdArray for first input
    pEnteredInputs = 1
    curPhrases = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
    curPhraseTypes = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
    curPhraseFormatting = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
    mAddr.Add curPhrases
    mAddr.Add curPhraseTypes
    mAddr.Add curPhraseFormatting
    
    pInputPhrases.Add curPhrases, 1
    pInputPhraseTypes.Add curPhraseTypes, 1
    pInputPhraseFormatting.Add curPhraseFormatting, 1
    
    'Get End of Function...
    For X = TotPhrases To 1 Step -1
        If gdGetNum(pPhraseTypes, X) = gPT_FRParen Then
            lLastPhrase = X
            Exit For
        End If
    Next X
    
    'Check for "Junk" to right of FRight paren.  This could happen during
    'editing.  Return an Error.
    If lLastPhrase < TotPhrases And TotPhrases > 1 Then
        If gdGetNum(pPhraseTypes, lLastPhrase + 1) <> gPT_Offset Then
            SetErrorCode 2, pCurArrayElement
        End If
    End If
    
    'This function has no inputs
    If 3 > lLastPhrase - 1 Then
        pEnteredInputs = 0
        Exit Sub
    End If
    
    '3=Pull inputs between function's left and right paren.  Don't process
    'the left and right paren of the function.
    For X = 3 To lLastPhrase - 1
        curType = gdGetNum(pPhraseTypes, X)
        curPhrase = gdGetStr(pPhrases, X)
        curFormatting = gdGetStr(pPhraseFormatting, X)
        
        Select Case curType
            Case gPT_FLParen, gPT_LeftPar
                Parens = Parens + 1
                n = n + 1
                gdSetStr curPhrases, n, curPhrase
                gdSetNum curPhraseTypes, n, curType
                gdSetStr curPhraseFormatting, n, curFormatting
                
            Case gPT_FRParen, gPT_RightPar
                Parens = Parens - 1
                n = n + 1
                gdSetStr curPhrases, n, curPhrase
                gdSetNum curPhraseTypes, n, curType
                gdSetStr curPhraseFormatting, n, curFormatting
                
            Case gPT_Comma
                If Parens = 0 Then
                    'Phrase(s) have been entered
                    If n > 0 Then
                        pEnteredInputs = pEnteredInputs + 1
                        gdFreeExtra curPhrases
                        gdFreeExtra curPhraseTypes
                        gdFreeExtra curPhraseFormatting
                        curPhrases = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
                        curPhraseTypes = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
                        curPhraseFormatting = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
                        mAddr.Add curPhrases
                        mAddr.Add curPhraseTypes
                        mAddr.Add curPhraseFormatting
                        pInputPhrases.Add curPhrases
                        pInputPhraseTypes.Add curPhraseTypes
                        pInputPhraseFormatting.Add curPhraseFormatting
                        n = 0
                    Else
                        pEnteredInputs = pEnteredInputs + 1
                        gdFreeExtra curPhrases
                        gdFreeExtra curPhraseTypes
                        gdFreeExtra curPhraseFormatting
                        curPhrases = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
                        curPhraseTypes = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
                        curPhraseFormatting = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
                        mAddr.Add curPhrases
                        mAddr.Add curPhraseTypes
                        mAddr.Add curPhraseFormatting
                        pInputPhrases.Add curPhrases
                        pInputPhraseTypes.Add curPhraseTypes
                        pInputPhraseFormatting.Add curPhraseFormatting
                    End If
                Else
                    'Comma is part of an embedded fuction input list
                    n = n + 1
                    gdSetStr curPhrases, n, curPhrase
                    gdSetNum curPhraseTypes, n, curType
                    gdSetStr curPhraseFormatting, n, curFormatting
                End If
                
            Case Else
                'Build gdArray of input phrases...
                n = n + 1
                gdSetStr curPhrases, n, curPhrase
                gdSetNum curPhraseTypes, n, curType
                gdSetStr curPhraseFormatting, n, curFormatting
                
        End Select
    Next X
    
    gdFreeExtra curPhrases
    gdFreeExtra curPhraseTypes
    gdFreeExtra curPhraseFormatting
    
ErrExit:
    Exit Sub

ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:GetFunctionInputs ", Err.Description
    End If
    
End Sub

'For debugging, this creates a string of phrases
Private Function StrPhrase(pPhrases As Long) As String
On Error GoTo ErrSection:
    Dim X       As Long
    Dim s       As String
    
    For X = 1 To gdGetSize(pPhrases) - 1
        s = s & gdGetStr(pPhrases, X) & " "
    Next X
    StrPhrase = Trim(s)
    
    Exit Function
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Function:StrPhrase ", Err.Description
    End If
End Function

Private Sub ValidateCodedText(pCodedText As String, pName As String)
On Error GoTo ErrSection:
    Dim NextTokenPos        As Long
    Dim strPhraseType       As String
    Dim strPhraseLen        As Long
    
    NextTokenPos = 1
    Do Until NextTokenPos = 0
        
        NextTokenPos = InStr(NextTokenPos, pCodedText, "~")
        If NextTokenPos > 0 Then
        
            'Verify the phrase type and phrase len are numeric
            strPhraseType = Mid(pCodedText, NextTokenPos + 1, 2)
            strPhraseLen = Mid(pCodedText, NextTokenPos + 3, 3)
            If Not IsNumeric(strPhraseType) Or _
               Not IsNumeric(strPhraseLen) Then
                Err.Raise gUserErr, , "An error was found in the codedtext of " & _
                    pName & ".  Please Reverify the rule."
            End If
            NextTokenPos = NextTokenPos + 1
        End If
    Loop
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:ValidateCodedText ", Err.Description
    End If
End Sub

'Convert the coded text string to a phrases and phrasetypes array
Private Sub ConvertCT(ByVal pCodedText As String, _
    pPhrasesArray As Long, pPhraseTypesArray As Long, pPhraseFormatting As Long)
On Error GoTo ErrSection:
    Dim curPhrase       As String
    Dim PType           As Byte
    Dim PhraseLen       As Long
    Dim NextTokenPos    As Long
    Dim X               As Long
    Dim pTypepos        As Long
    Dim OffsetValTokenPos   As Long
    Dim Continue        As Boolean
    Dim OffsetVal       As String
    Dim T               As String
    
    'Create array to storing phrases and phrase types of coded text
    pPhrasesArray = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
    pPhraseTypesArray = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
    pPhraseFormatting = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
    mAddr.Add pPhrasesArray
    mAddr.Add pPhraseTypesArray
    mAddr.Add pPhraseFormatting
    
    X = 0
    NextTokenPos = 1
    Do Until NextTokenPos = 0
        
        NextTokenPos = InStr(NextTokenPos, pCodedText, "~")
        If NextTokenPos > 0 Then
        
            'Get current phrase
            PType = Val(Mid(pCodedText, NextTokenPos + 1, 2))
            PhraseLen = Val(Mid(pCodedText, NextTokenPos + 3, 3))
            curPhrase = Mid(pCodedText, NextTokenPos + 6, PhraseLen)
            
            'If phrase is an offset and it's offset is zero, then ignore the
            'offset period and the zero value
            Continue = True
            Select Case PType
            
                Case gPT_Offset
                    
                    'Get offset information
                    OffsetValTokenPos = InStr(NextTokenPos + 1, pCodedText, "~")
                    PType = Val(Mid(pCodedText, OffsetValTokenPos + 1, 2))
                    PhraseLen = Val(Mid(pCodedText, OffsetValTokenPos + 3, 3))
                    OffsetVal = Mid(pCodedText, OffsetValTokenPos + 6, PhraseLen)
                    If OffsetVal <> "0" Then
                        X = X + 1
                        gdSetNum pPhraseTypesArray, X, gPT_Offset
                        gdSetStr pPhrasesArray, X, "."
                        gdSetStr pPhraseFormatting, X, ""
                        X = X + 1
                        gdSetNum pPhraseTypesArray, X, PType
                        gdSetStr pPhrasesArray, X, OffsetVal
                        gdSetStr pPhraseFormatting, X, T
                    End If
                    NextTokenPos = OffsetValTokenPos
                    T = ""
                    
                'Tag the next phrase with enter key formatting data
                Case gPT_EnterFormatting: T = T & "E"
                
                Case gPT_Tab: T = T & "T"
                        
                Case Else
                    X = X + 1
                    gdSetNum pPhraseTypesArray, X, PType
                    gdSetStr pPhrasesArray, X, curPhrase
                    gdSetStr pPhraseFormatting, X, T
                    T = ""
            End Select
            
            NextTokenPos = NextTokenPos + 1
        End If
    Loop
    
    gdFreeExtra pPhrasesArray
    gdFreeExtra pPhraseTypesArray
    gdFreeExtra pPhraseFormatting
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:ConvertCT ", Err.Description
    End If
End Sub

'Save the first error that occurs.  But flag all array elements in error.
Private Sub SetErrorCode(pCode As Byte, pCurArrayElement As Long)
    If mErrorCode = 0 Then
        SetError pCurArrayElement, 1
        mErrorCode = pCode
    End If
End Sub

'Converts from token value back to string.  This is used by the report to
'make it easier to interpret phrase type
Public Function GetPhraseTypeDesc(pToken As Long) As String
On Error GoTo ErrSection:
    Select Case pToken
        Case gPT_Nbr: GetPhraseTypeDesc = "Nbr"
        Case gPT_If: GetPhraseTypeDesc = "IF"
        Case gPT_DoUntil: GetPhraseTypeDesc = "DoUntil"
        Case gPT_EndDo: GetPhraseTypeDesc = "EndDo"
        Case gPT_Then: GetPhraseTypeDesc = "Then"
        Case gPT_Comma: GetPhraseTypeDesc = "Comma"
        Case gPT_Offset: GetPhraseTypeDesc = "Offset"
        Case gPT_Add: GetPhraseTypeDesc = "Add"
        Case gPT_Sub: GetPhraseTypeDesc = "Sub"
        Case gPT_Mult: GetPhraseTypeDesc = "Mult"
        Case gPT_Div: GetPhraseTypeDesc = "Div"
        Case gPT_GT: GetPhraseTypeDesc = "GT"
        Case gPT_GE: GetPhraseTypeDesc = "GE"
        Case gPT_LT: GetPhraseTypeDesc = "LT"
        Case gPT_LE: GetPhraseTypeDesc = "LE"
        Case gPT_NE: GetPhraseTypeDesc = "NE"
        Case gPT_EQ: GetPhraseTypeDesc = "EQ"
        Case gPT_And: GetPhraseTypeDesc = "And"
        Case gPT_Not: GetPhraseTypeDesc = "Not"
        Case gPT_Or: GetPhraseTypeDesc = "Or"
        Case gPT_Of: GetPhraseTypeDesc = "Of"
        Case gPT_Text: GetPhraseTypeDesc = "Text"
        Case gPT_Comment: GetPhraseTypeDesc = "Comment"
        Case gPT_Enter: GetPhraseTypeDesc = "Enter"
        Case gPT_EnterFormatting: GetPhraseTypeDesc = "EntFor"
        Case gPT_FCompiled: GetPhraseTypeDesc = "Compiled"
        Case gPT_FCompiledAction: GetPhraseTypeDesc = "Comp Act"
        Case gPT_FInternal: GetPhraseTypeDesc = "Internal"
        Case gPT_FTradeSense: GetPhraseTypeDesc = "TradeSense"
        Case gPT_Parm: GetPhraseTypeDesc = "Input"
        Case gPT_OuterParens: GetPhraseTypeDesc = "Parens"
        Case gPT_Trades: GetPhraseTypeDesc = "Trades"
        Case gPT_Portfolio: GetPhraseTypeDesc = "Portfolio"
        Case gPT_Systems: GetPhraseTypeDesc = "Systems"
        Case gPT_Bars: GetPhraseTypeDesc = "Bars"
        Case gPT_Else: GetPhraseTypeDesc = "Else"
        Case gPT_ElseIf: GetPhraseTypeDesc = "ElseIf"
        Case gPT_EndIf: GetPhraseTypeDesc = "EndIf"
        Case gPT_Tab: GetPhraseTypeDesc = "Tab"
        Case Else:
            GetPhraseTypeDesc = "None"
    End Select
    
ErrExit:
    Exit Function
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:GetPhraseTypeDesc ", Err.Description
    End If
End Function

'==================== Called by Engine ===================================
'This routine creates gdArray's for each tree node.  These keep track of the
'calculations results as of each trade.  If, ElseIf, and Then inherit their address from their
'children.
Public Sub CreateCalcArrays()
On Error GoTo ErrSection:
    Dim X               As Long
    Dim curPhraseType   As Byte
    Dim gdArrayAddr     As Long
    Dim arIndex         As cGdArray
    Dim Index           As Long
    Dim ptr             As Long
    Dim svPhrase        As String
    Dim curPhrase       As String
    Dim curReturnType   As Byte
    Dim ArraySize       As Long
    Dim rc              As Variant
    Dim ArrayToOffsetType   As Long
    Dim Continue        As Boolean
    Dim PhraseReturn    As cGdArray
    Dim aPhraseReturn   As Long
    Dim SS              As String
    
    ArraySize = gdGetSize(gdGetNum(mTradesAddr, 1))

    'create a sorted index on Phrase name
    Set arIndex = New cGdArray
    mRetVal = arIndex.Create(eGDARRAY_Longs, mLast)
    arIndex.Reserve mLast, 0

    'Build array of strings with Phrase and return type concatanated.  This
    'fixes the bug with Input names being the same as constant text names.
    Set PhraseReturn = New cGdArray
    rc = PhraseReturn.Create(eGDARRAY_Strings, mLast)
    PhraseReturn.Reserve mLast, 0
    For X = 1 To mLast
        SS = Phrase(X)
        
        'Make text constants unique so they don't have the same name as
        'inputs.
        If ReturnType(X) = gRetConstantText Or ReturnType(X) = gRetSeriesText Then
            SS = SS & "Text"
        End If
        PhraseReturn.Add SS, X
    Next X
    
    aPhraseReturn = PhraseReturn.ArrayHandle
    Index = arIndex.ArrayHandle
    rc = gdSortAsIndex(Index, aPhraseReturn, 1, eGdSort_Default, 1, 999999999)
    'rc = gdSortAsIndex(Index, aPhrase, 1, eGdSort_Default, 1, 999999999)
    
    'For X = 1 To mLast
    '    Debug.Print Phrase(gdGetNum(Index, X)) & ReturnType(gdGetNum(Index, X))
    'Next X
    'Debug.Print
    
    For X = 1 To mLast

        'Get current phrase attributes...
        ptr = gdGetNum(Index, X)
        curPhraseType = PhraseType(ptr)
        curPhrase = Phrase(ptr)
        curReturnType = ReturnType(ptr)
        
        'If curPhrase = "Units calculation" Then
        '    Debug.Print
        'End If
        
        Select Case curPhraseType
            
            Case gPT_Trades: gdArrayAddr = mTradesAddr
            Case gPT_Portfolio: gdArrayAddr = mPortfolioAddr
            Case gPT_Systems: gdArrayAddr = mSystemsAddr
            
            'Skip array allocation for these operators
            Case gPT_Enter, gPT_If, gPT_DoUntil, gPT_Then, gPT_ElseIf, gPT_Else, gPT_EndIf, gPT_EndDo
            
            Case gPT_GT, gPT_GE, gPT_LT, gPT_LE, gPT_NE, gPT_EQ, _
                 gPT_And, gPT_Or, gPT_Not
                gdArrayAddr = gdCreateArray(eGDARRAY_TinyInts, ArraySize, USE_DEFAULT_NULL)
                gdReserve gdArrayAddr, ArraySize, 0
                mAddr.Add gdArrayAddr
                
            'Offset array must be the same as the array type of it's left child.  Check it's
            'type and create it based on that type.
            Case gPT_Offset
                ArrayToOffsetType = gdGetType(PoolAddr(NodeFrom(ptr)))
                Select Case ArrayToOffsetType
                    Case 68: gdArrayAddr = gdCreateArray(eGDARRAY_Doubles, ArraySize, USE_DEFAULT_NULL)
                    Case 33: gdArrayAddr = gdCreateArray(eGDARRAY_Strings, ArraySize, USE_DEFAULT_NULL)
                    Case 84: gdArrayAddr = gdCreateArray(eGDARRAY_TinyInts, ArraySize, USE_DEFAULT_NULL)
                    Case 76: gdArrayAddr = gdCreateArray(eGDARRAY_Longs, ArraySize, USE_DEFAULT_NULL)
                End Select
                gdReserve gdArrayAddr, ArraySize, 0
                mAddr.Add gdArrayAddr

            Case gPT_Text
                gdArrayAddr = gdCreateArray(eGDARRAY_Strings, ArraySize, USE_DEFAULT_NULL)
                gdReserve gdArrayAddr, ArraySize, 0
                mAddr.Add gdArrayAddr
            
            'Get index of Array in internal array structure
            Case gPT_FInternal
                'Get address of trade function gdArrays
                If InStr(1, ClassName(ptr), "~~Trades") > 0 Then
                    gdArrayAddr = InternalItemAddr(mTradesAddr, ClassName(ptr))
                Else
                    If InStr(1, ClassName(ptr), "~~Portfolio") > 0 Then
                        gdArrayAddr = InternalItemAddr(mPortfolioAddr, ClassName(ptr))
                    Else
                        Err.Raise gUserErr, , "Could not retrieve a gdArray address for " & ClassName(ptr)
                    End If
                End If
                
            Case Else
                'Some compiled actions that directly update structures return results
                'so the array type must be valid.  For example, "Set Var Text".
                If curReturnType = gRetSeriesText Or _
                   curReturnType = gRetConstantText Then
                    gdArrayAddr = gdCreateArray(eGDARRAY_Strings, ArraySize, USE_DEFAULT_NULL)
                Else
                    gdArrayAddr = gdCreateArray(eGDARRAY_Doubles, ArraySize, USE_DEFAULT_NULL)
                End If
                gdReserve gdArrayAddr, ArraySize, 0
                mAddr.Add gdArrayAddr
               
        End Select
        
        '----------------------------------------------------------------
        'This section makes constant type arrays
        
        'IMPORTANT Note:  This logic assumes inputs are always constant
        'arrays.  If expressions are ever entered into the inputs grid, then
        'this logic will have to change.
        
        'Create constant arrays for return types of constant and inputs
        Continue = True
        If curReturnType = gRetConstantNbr Or _
            curReturnType = gRetConstantBoolean Or _
            curPhraseType = gPT_Parm Then
            
            'If constant phrase type of internal structure
            '(i.e. Begin Balance in Portfolio) then the phrase won't
            'have a valid value.  Get element 1.
            If curPhraseType = gPT_FInternal Then
                gdMakeConstantValue gdArrayAddr, _
                    gdGetNum(gdArrayAddr, 0), ArraySize
            Else
                gdMakeConstantValue gdArrayAddr, 0, ArraySize
                gdSetNum gdArrayAddr, 0, Val(curPhrase)
            End If
            SetPoolAddr ptr, gdArrayAddr
            SetRebuildArray ptr, 0
            If curPhraseType <> gPT_Parm Then
                Continue = False
            End If
        Else
            'Constant text array
            If curReturnType = gRetConstantText Then
                If curPhraseType = gPT_FInternal Then
                    gdMakeConstantValue gdArrayAddr, _
                        gdGetStr(gdArrayAddr, 0), ArraySize
                Else
                    gdMakeConstantValue gdArrayAddr, 0, ArraySize
                    gdSetStr gdArrayAddr, 0, curPhrase
                End If
                SetPoolAddr ptr, gdArrayAddr
                SetRebuildArray ptr, 0
                Continue = False
            End If
        End If
        
        '--------------------------------------------------------------------
        'Set this address to all phrases with the same name to prevent these
        'arrays from being calculated more than once.  Exception: constant
        'arrays for numbers (not parms)...
        If Continue Then
            Select Case curPhraseType
                Case gPT_If, gPT_Then, gPT_ElseIf, _
                     gPT_Else, gPT_EndIf, gPT_DoUntil, gPT_EndDo
                Case Else
                    svPhrase = curPhrase
                    Do Until svPhrase <> curPhrase
                        SetPoolAddr ptr, gdArrayAddr
                        SetRebuildArray ptr, 0
                        X = X + 1
                        If X > mLast Then Exit Do
                        ptr = gdGetNum(Index, X)
                        curPhrase = Phrase(ptr)
                    Loop
                    X = X - 1
            End Select
        End If
        
    Next X
    
    'Go back through and set the pool array address the same for the
    'parent and child of Phrases of types: parens, TradeSense, If, ElseIf,
    'and Then
    For X = mLast To 1 Step -1
        curPhraseType = PhraseType(X)
        Select Case curPhraseType
            Case gPT_FTradeSense, gPT_OuterParens, gPT_Then, gPT_Enter
                gdSetNum aPoolAddr, X, PoolAddr(NodeFrom(X))
            Case gPT_If, gPT_ElseIf, gPT_DoUntil
                gdSetNum aPoolAddr, X, PoolAddr(NodeTo(X))
        End Select
    Next X
    
    Set PhraseReturn = Nothing
    Set arIndex = Nothing
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:CreateCalcArrays ", Err.Description
    End If
End Sub

'Returns the gdArray Address of Trades Function with Trades structure.  Based
'on the Class value (in the format "~~Trades.nn where nn is the array nbr in
'the Trades structure).  If not found then function returns 0.
'Note: The pClass can be in the format nn -x where -x is the offset to the array
'This is used by the "Next" functions.
Private Function InternalItemAddr(pItemAddr As Long, pClass As String) As Long
On Error GoTo ErrSection:
    Dim ArrayPtr    As Long
    Dim ArrayAddr   As Long
    Dim ArrayNew    As Long
    Dim ShiftVal    As Long
    
    InternalItemAddr = 0
    
    If InStr(1, pClass, ".") = 0 Then Exit Function
    ArrayPtr = Val(Mid(pClass, InStr(1, pClass, ".") + 1, 3))
    ArrayAddr = gdGetNum(pItemAddr, ArrayPtr)
    InternalItemAddr = ArrayAddr
    
    '---------------------------------------------------------------------
    '"Next Function" handling
    'Look for a shift value in pClass AFTER the ArrayPtr value (ie "009-1").
    'Make a copy of the internal array and then shift it.  This array address
    'must be added to "mAddr" so it is destroyed later.
    If InStr(1, pClass, "-") > 0 Then
        ShiftVal = Val(Mid(pClass, InStr(1, pClass, "-"), 2))
        ArrayNew = gdCreateArray(gdGetType(ArrayAddr))
        gdCopy ArrayNew, ArrayAddr
        gdSetShifted ArrayNew, ShiftVal, 0
        mAddr.Add ArrayNew
        InternalItemAddr = ArrayNew
    End If
    '---------------------------------------------------------------------
    
    Exit Function
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, "Module:mComEval Sub:InternalItemAddr ", Err.Description
    End If
End Function


'===================== Called by Engine =====================================
'This routine builds parameter lists (gdArray of addresses to parms)
'for each compiled function in the tree...
Public Sub CreateParmListAddrs()
On Error GoTo ErrSection:
    Dim P           As Long
    Dim X           As Long
    Dim Y           As Long
    Dim aParms      As Long     'Address of current parms list array
    Dim TotParms    As Long
    
    'TESTING: Make sure this routine is called  when the inputs change
    'because the parmlists will now be different.  The arrays don't need
    'to be updated, but constant arrays must be updated with the new
    'input values...
    
    For X = 1 To mLast
    
        'Continue if phrase is a compiled function
        If PhraseType(X) = gPT_FCompiled Or _
           PhraseType(X) = gPT_FCompiledAction Then
            
            'Create an array used to store the addresses of each parameter
            'gdArray used by the function...
            ' - Pos 1: Results array address
            ' - Pos 2: Address to Long Array of 1 element holding the
            '             current trade nbr (used by late calc)
            ' - Pos 3: Address to Long Array of 1 holding Error Code
            ' - Pos 4: Address to String Array of 1 holding Error message
            '             set by called trading function in DLL.
            ' - Pos 5-n: Beginning of Parameters expected by trading functions
            TotParms = NodeTo(X) - NodeFrom(X) + 1 + gReservedParms
            aParms = gdCreateArray(eGDARRAY_Longs, TotParms, USE_DEFAULT_NULL)
            gdReserve aParms, TotParms, 0
            mAddr.Add aParms
            
            'Create an array for each parameter.  Add the address of the
            'parameter array to the Parameter List array (aParms)
            P = gReservedParms
            For Y = NodeFrom(X) To NodeTo(X)
                If PoolAddr(Y) = 0 Then
                    Err.Raise gUserErr, , _
                        "Error occured while building the Parameter List array.  An array address was not found for  " & Phrase(Y) & ".  Test did not complete."
                End If
                P = P + 1
                gdSetNum aParms, P, PoolAddr(Y)
            Next Y
            
            'Save address for parmlist array
            SetParmList X, aParms
            
        End If
    Next X

    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:CreateParmListAddrs ", Err.Description
    End If
End Sub

Private Sub Class_Initialize()
On Error GoTo ErrSection:
    Dim Success As Boolean
    
    Set mPhraseType = New cGdArray
    Set mReturnType = New cGdArray
    Set mPhrase = New cGdArray
    Set mFormatting = New cGdArray
    Set mPhrasesArray = New cGdArray
    Set mPhraseTypesArray = New cGdArray
    Set mPhraseFormattingArray = New cGdArray
    Set mClassName = New cGdArray
    Set mClassPtr = New cGdArray
    Set mLate = New cGdArray
    Set mSize = New cGdArray
    Set mNodeFrom = New cGdArray
    Set mNodeTo = New cGdArray
    Set mParent = New cGdArray
    Set mRuleID = New cGdArray
    Set mExeTrue = New cGdArray
    Set mExeFalse = New cGdArray
    Set mPoolAddr = New cGdArray
    Set mParmList = New cGdArray
    Set mRebuildArray = New cGdArray
    Set mFID = New cGdArray
    Set mInputID = New cGdArray
    Set mReq = New cGdArray
    Set mEntered = New cGdArray
    Set mError = New cGdArray

    Success = mPhraseType.Create(eGDARRAY_TinyInts, 2000)
    Success = mReturnType.Create(eGDARRAY_TinyInts, 2000, 0)
    Success = mPhrase.Create(eGDARRAY_Strings, 2000)
    Success = mPhrasesArray.Create(eGDARRAY_Longs, 2000)
    Success = mFormatting.Create(eGDARRAY_Strings, 2000)
    Success = mPhraseTypesArray.Create(eGDARRAY_Longs, 2000)
    Success = mPhraseFormattingArray.Create(eGDARRAY_Longs, 2000)
    Success = mClassName.Create(eGDARRAY_Strings, 2000)
    Success = mClassPtr.Create(eGDARRAY_Longs, 2000)
    Success = mLate.Create(eGDARRAY_TinyInts, 2000)
    Success = mSize.Create(eGDARRAY_Longs, 2000)
    Success = mNodeFrom.Create(eGDARRAY_Longs, 2000)
    Success = mNodeTo.Create(eGDARRAY_Longs, 2000)
    Success = mParent.Create(eGDARRAY_Longs, 2000)
    Success = mRuleID.Create(eGDARRAY_Longs, 2000)
    Success = mExeTrue.Create(eGDARRAY_Longs, 2000)
    Success = mExeFalse.Create(eGDARRAY_Longs, 2000)
    Success = mPoolAddr.Create(eGDARRAY_Longs, 2000)
    Success = mParmList.Create(eGDARRAY_Longs, 2000)
    Success = mRebuildArray.Create(eGDARRAY_TinyInts, 2000)
    Success = mFID.Create(eGDARRAY_Longs, 2000)
    Success = mInputID.Create(eGDARRAY_Longs, 2000)
    Success = mReq.Create(eGDARRAY_TinyInts, 2000)
    Success = mEntered.Create(eGDARRAY_TinyInts, 2000)
    Success = mError.Create(eGDARRAY_TinyInts, 2000)
    
    mPhraseType.Reserve 2000, 0
    mReturnType.Reserve 2000, 0
    mPhrase.Reserve 2000, 0
    mFormatting.Reserve 2000, 0
    mPhrasesArray.Reserve 2000, 0
    mPhraseTypesArray.Reserve 2000, 0
    mPhraseFormattingArray.Reserve 2000, 0
    mClassName.Reserve 2000, 0
    mClassPtr.Reserve 2000, 0
    mLate.Reserve 2000, 0
    mSize.Reserve 2000, 0
    mNodeFrom.Reserve 2000, 0
    mNodeTo.Reserve 2000, 0
    mParent.Reserve 2000, 0
    mRuleID.Reserve 2000, 0
    mExeTrue.Reserve 2000, 0
    mExeFalse.Reserve 2000, 0
    mPoolAddr.Reserve 2000, 0
    mParmList.Reserve 2000, 0
    mRebuildArray.Reserve 2000, 0
    mFID.Reserve 2000, 0
    mInputID.Reserve 2000, 0
    mReq.Reserve 2000, 0
    mEntered.Reserve 2000, 0
    mError.Reserve 2000, 0

    aPhraseType = mPhraseType.ArrayHandle
    aReturnType = mReturnType.ArrayHandle
    aPhrase = mPhrase.ArrayHandle
    aFormatting = mFormatting.ArrayHandle
    aPhrasesArray = mPhrasesArray.ArrayHandle
    aPhraseTypesArray = mPhraseTypesArray.ArrayHandle
    aPhraseFormattingArray = mPhraseFormattingArray.ArrayHandle
    aClassName = mClassName.ArrayHandle
    aClassPtr = mClassPtr.ArrayHandle
    aLate = mLate.ArrayHandle
    aSize = mSize.ArrayHandle
    aNodeFrom = mNodeFrom.ArrayHandle
    aNodeTo = mNodeTo.ArrayHandle
    aParent = mParent.ArrayHandle
    aRuleID = mRuleID.ArrayHandle
    aExeTrue = mExeTrue.ArrayHandle
    aExeFalse = mExeFalse.ArrayHandle
    aPoolAddr = mPoolAddr.ArrayHandle
    aParmList = mParmList.ArrayHandle
    aRebuildArray = mRebuildArray.ArrayHandle
    aFID = mFID.ArrayHandle
    aInputID = mInputID.ArrayHandle
    aReq = mReq.ArrayHandle
    aEntered = mEntered.ArrayHandle
    aError = mError.ArrayHandle
    
    mLast = 0
    
    Set mStatementBottom = New cGdArray
    mStatementBottom.Create eGDARRAY_Longs
    
    Set mAddr = New cGdArray
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.source, 5) = "Class" Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:Initialize ", Err.Description
    End If
End Sub

Private Sub Class_Terminate()
    Dim X   As Long
    Dim pos As Long
    
    mAddr.Sort eGdSort_Default, 1, mAddr.Size
    
    'Destroy gdArrays used in building trees (except for internal gdArrays)
    For X = mAddr.Size - 1 To 0 Step -1
        If mAddr.Item(X) > 0 Then
            gdDestroyArray mAddr.Item(X)
            mAddr.Remove X, 1
        End If
    Next X
    If mAddr.Size - 1 > 0 Then
        MsgBox "Items in mAddr still exist: " & mAddr.Size - 1
    End If
    
    Set mAddr = Nothing
    
    Set mPhraseType = Nothing
    Set mReturnType = Nothing
    Set mPhrase = Nothing
    Set mFormatting = Nothing
    Set mPhrasesArray = Nothing
    Set mPhraseTypesArray = Nothing
    Set mPhraseFormattingArray = Nothing
    Set mClassName = Nothing
    Set mClassPtr = Nothing
    Set mLate = Nothing
    Set mSize = Nothing
    Set mNodeFrom = Nothing
    Set mNodeTo = Nothing
    Set mParent = Nothing
    Set mRuleID = Nothing
    Set mExeTrue = Nothing
    Set mExeFalse = Nothing
    Set mFID = Nothing
    Set mInputID = Nothing
    Set mReq = Nothing
    Set mEntered = Nothing
    Set mError = Nothing
    Set mPoolAddr = Nothing
    Set mParmList = Nothing
    Set mRebuildArray = Nothing
    
    Set mStatementBottom = Nothing
End Sub

Private Function BS(pAddr As Long, pPos As Long) As Boolean
    Dim X   As Long
    BS = False
    For X = 1 To mAddr.Size
        If mAddr.Item(X) = pAddr Then
            BS = True
            pPos = X
            Exit Function
        End If
    Next X
End Function
