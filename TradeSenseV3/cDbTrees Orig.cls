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
'This class is responsible for building a tree of phrases.
'Rules: - Internal functions are inserted into tree
'       - Options inputs of functions are defaulted into tree (if not specified)

'Work variables
Private mFunctions          As cFunctions
Private mLast               As Long
Private mTradesAddr         As Long
Private mErrorCode          As Byte
Private mCalledFromEditor   As Boolean

'Trees database.  Addresses to phrases defining the rules.
Private aPhraseType         As Long
Private aReturnType         As Long
Private aPhrase             As Long
Private aPhrasesArray       As Long     'Address to array of sub phrases
Private aPhraseTypesArray   As Long     'Address to array of sub phrase types
Private aClassName          As Long     'Class name of function
Private aClassPtr           As Long     'Ptr to array holding class instance
Private aLate               As Long
Private aSize               As Long     'Array size set and allocated (debugging)
Private aNodeFrom           As Long
Private aNodeTo             As Long
Private aParent             As Long
Private aRuleID             As Long
Private aPoolAddr           As Long     'Holds the address of the array
                                        'for function types and expr's. It
                                        'holds a single value for nbr's.
Private aParmList           As Long     'Address to Parmlist
Private aRebuildArray       As Long     '1=Rebuild array.  Set to 1 when
                                        'New parms are inserted during opt.
                                        'This forces the array to be rebuilt
Private aFID                As Long     'Function ID (if function)
Private aInputID            As Long     'Input ID (if parent is function)
Private aReq                As Long     '1=Function Input is required
Private aEntered            As Long     '1=Phrase was user entered, not defaulted
Private aError              As Long     '1=Phrase is in error

'Array vectors used to stored the condition/action phrases
Private mCondPtr            As cVectorLng
Private mActionPtr          As cVectorLng
Private mCondTreeBottom     As cVectorLng
Private mActionTreeBottom   As cVectorLng
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
        Case 8: ErrMsg = "An expression is required to the right of the IF and THEN operator's"
        Case 9: ErrMsg = "Make sure IF and THEN operator are the first operators used, and are on separate lines"
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

'Returns pointers to the conditions and associated actions
Property Get CondPtr() As cVectorLng
    Set CondPtr = mCondPtr
End Property
Property Get CondTreeBottom() As cVectorLng
    Set CondTreeBottom = mCondTreeBottom
End Property
Property Get ActionPtr() As cVectorLng
    Set ActionPtr = mActionPtr
End Property
Property Get ActionTreeBottom() As cVectorLng
    Set ActionTreeBottom = mActionTreeBottom
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
Property Get AddrPhrasesArray() As Long
    AddrPhrasesArray = aPhrasesArray
End Property
Property Get AddrPhraseTypesArray() As Long
    AddrPhraseTypesArray = aPhraseTypesArray
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
'THIS IS THE ENTRY POINT INTO THE TREE.  USE THIS ROUTINE TO ADD CODEDTEXT INTO THE TREE.
Public Sub AddCodedText(pCodedText As String, pName As String)
On Error GoTo ErrSection:
    Dim Phrases         As Long
    Dim PhraseTypes     As Long
    
    mCalledFromEditor = False

    'Stripout IF phrase
    If Left(pCodedText, 3) = "~" & gPT_If Then
        pCodedText = Right(pCodedText, Len(pCodedText) - 9)
    End If
    
    ValidateCodedText pCodedText, pName
    ConvertCT pCodedText, Phrases, PhraseTypes
    Add 0, 0, StrPhrase(Phrases), Phrases, PhraseTypes, "", _
        0, 0, 0, 0, 0, 0, 0
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:AddCodedText ", Err.Description
    End If
End Sub
        
'This is called from cExpression.  It adds a condition or action string to
'the top level node of the tree.
Public Sub AddEnglishText(pPhrases As Long, pPhraseTypes As Long)
On Error GoTo ErrSection:
    
    mCalledFromEditor = True
    Add 0, 0, StrPhrase(pPhrases), pPhrases, pPhraseTypes, "", _
        0, 0, 0, 0, 0, 0, 1
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:AddEnglishText ", Err.Description
    End If
End Sub
'==========================================================================


'Add's a phrase to the Tree...
Private Sub Add(pPhraseType As Byte, pReturnType As Byte, _
    pPhrase As String, pPhrasesArrayAddr As Long, pPhraseTypesArrayAddr As Long, _
    pClassName As String, pLate As Byte, _
    pNodeFrom As Long, pNodeTo As Long, pParent As Long, _
    pPoolAddr As Long, pReq As Byte, pEntered As Byte)
On Error GoTo ErrSection:
    
    mLast = mLast + 1
    gdSetNum aPhraseType, mLast, pPhraseType
    gdSetNum aReturnType, mLast, pReturnType
    gdSetStr aPhrase, mLast, pPhrase
    gdSetNum aPhrasesArray, mLast, pPhrasesArrayAddr
    gdSetNum aPhraseTypesArray, mLast, pPhraseTypesArrayAddr
    gdSetStr aClassName, mLast, pClassName
    gdSetNum aClassPtr, mLast, 0
    gdSetNum aLate, mLast, pLate
    gdSetNum aSize, mLast, 0
    gdSetNum aNodeFrom, mLast, pNodeFrom
    gdSetNum aNodeTo, mLast, pNodeTo
    gdSetNum aParent, mLast, pParent
    gdSetNum aRuleID, mLast, 0
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
    pHighestPrecidencePhraseType As Byte, pPhraseTypeArrayPos As Long)
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
                    Prec = 10
                Case gPT_If
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
            Exit Sub
        End If
        
        curType = gdGetNum(pPhraseTypes, 1)
        
        If curType = gPT_FCompiled Or curType = gPT_FCompiledAction Or _
           curType = gPT_FTradeSense Or curType = gPT_FInternal Or _
           curType = gPT_Then Then
            pHighestPrecidencePhraseType = gdGetNum(pPhraseTypes, 1)
            pPhraseTypeArrayPos = 1
        Else
            'If only one phrase in array and it's a nbr, text, or parm then
            'this is valid.  Set the phrase type.
            If (curType = gPT_Nbr Or curType = gPT_Text Or gPT_Comment Or _
               curType = gPT_Parm) And Cnt = 1 Then
                pHighestPrecidencePhraseType = gdGetNum(pPhraseTypes, 1)
                pPhraseTypeArrayPos = 1
            Else
                'Multiple phrases exist with the first phrase being an
                'operator! This is an error!
                pHighestPrecidencePhraseType = 0
                pPhraseTypeArrayPos = 1
            End If
        End If
        
    Else
        pHighestPrecidencePhraseType = MaxPrecType
        pPhraseTypeArrayPos = MaxPrecPos
    End If
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
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
    Dim Cnt             As Long
    Dim X               As Long
    Dim y               As Long
    Dim QuitParsing     As Boolean
    Dim i               As Long
    
    'Break down phrases into subphrases...
    Do Until QuitParsing
        X = X + 1
        If X > mLast Then Exit Do
        
        'Determine highest presidence operator and it's array position...
        'If an operator isn't found then one phrase exists...
        DeterminePhraseType PhraseTypesArray(X), PType, PhraseTypePos
        
        Select Case PType
            
            Case gPT_If, gPT_Then, gPT_Not
            
                SetPhraseType X, PType

                'Error - no stuff on right of Operator
                If gdGetSize(PhraseTypesArray(X)) - 1 - PhraseTypePos = 0 Then
                    SetErrorCode 8, X
                Else
                    'Error - Stuff found on Left side of Operator
                    If PhraseTypePos - 1 > 0 Then
                        SetErrorCode 9, X
                    Else
                        'Stuff found on right side of Operator
                        SetNodeFrom X, mLast + 1
                        SetNodeTo X, mLast + 1
                    
                        Cnt = gdGetSize(PhraseTypesArray(X)) - 1
                        aPhrs = gdCreateArray(eGDARRAY_Strings, Cnt - PhraseTypePos, USE_DEFAULT_NULL)
                        aPhrTypes = gdCreateArray(eGDARRAY_TinyInts, Cnt - PhraseTypePos, USE_DEFAULT_NULL)
                        i = 0
                        For y = PhraseTypePos + 1 To Cnt
                            i = i + 1
                            gdSetStr aPhrs, i, gdGetStr(PhrasesArray(X), y)
                            gdSetNum aPhrTypes, i, gdGetNum(PhraseTypesArray(X), y)
                        Next y
                        Add 0, 0, StrPhrase(aPhrs), aPhrs, aPhrTypes, "", _
                            0, 0, 0, X, 0, 0, Entered(X)
                    End If
                End If
            
            'Add tree entries operations phrases
            Case gPT_And, gPT_Or, gPT_Of, gPT_GT, gPT_GE, gPT_LT, gPT_LE, gPT_NE, _
                 gPT_EQ, gPT_Mult, gPT_Div, gPT_Add, gPT_Sub, gPT_Offset, _
                 gPT_Enter

                'Make sure stuff is on both sides of operator
                If gdGetSize(PhraseTypesArray(X)) - 1 - PhraseTypePos = 0 Or _
                   PhraseTypePos - 1 = 0 Then
                    SetPhraseType X, PType
                    SetErrorCode 4, X
                Else
                    'Update properties of current node
                    SetPhraseType X, PType
                    SetNodeFrom X, mLast + 1
                    SetNodeTo X, mLast + 2
                    
                    'Create array for storing phrases and phrase types of
                    'stuff to left of operator
                    aPhrs = gdCreateArray(eGDARRAY_Strings, PhraseTypePos - 1, USE_DEFAULT_NULL)
                    aPhrTypes = gdCreateArray(eGDARRAY_TinyInts, PhraseTypePos - 1, USE_DEFAULT_NULL)
                    For y = 1 To PhraseTypePos - 1
                        gdSetStr aPhrs, y, gdGetStr(PhrasesArray(X), y)
                        gdSetNum aPhrTypes, y, gdGetNum(PhraseTypesArray(X), y)
                    Next y
                    Add 0, 0, StrPhrase(aPhrs), aPhrs, aPhrTypes, "", _
                        0, 0, 0, X, 0, 0, Entered(X)
                    
                    'Create array for storing phrases and phrase types of
                    'stuff to right of operator
                    Cnt = gdGetSize(PhraseTypesArray(X)) - 1
                    aPhrs = gdCreateArray(eGDARRAY_Strings, Cnt - PhraseTypePos, USE_DEFAULT_NULL)
                    aPhrTypes = gdCreateArray(eGDARRAY_TinyInts, Cnt - PhraseTypePos, USE_DEFAULT_NULL)
                    i = 0
                    For y = PhraseTypePos + 1 To Cnt
                        i = i + 1
                        gdSetStr aPhrs, i, gdGetStr(PhrasesArray(X), y)
                        gdSetNum aPhrTypes, i, gdGetNum(PhraseTypesArray(X), y)
                    Next y
                    Add 0, 0, StrPhrase(aPhrs), aPhrs, aPhrTypes, "", _
                        0, 0, 0, X, 0, 0, Entered(X)
                End If
                
            Case gPT_OuterParens
                SetPhraseType X, PType
                'SetReturnType X, GetReturnType(PType)
                Cnt = gdGetSize(PhraseTypesArray(X)) - 1
                
                'If one or more phrases exist between parens...
                If Cnt - 2 > 0 Then
                    SetNodeFrom X, mLast + 1
                    SetNodeTo X, mLast + 1
                    
                    'Create array for storing phrases and phrase types of
                    'stuff in between outerparens
                    aPhrs = gdCreateArray(eGDARRAY_Strings, Cnt - 2, USE_DEFAULT_NULL)
                    aPhrTypes = gdCreateArray(eGDARRAY_TinyInts, Cnt - 2, USE_DEFAULT_NULL)
                    i = 0
                    For y = 2 To Cnt - 1
                        i = i + 1
                        gdSetStr aPhrs, i, gdGetStr(PhrasesArray(X), y)
                        gdSetNum aPhrTypes, i, gdGetNum(PhraseTypesArray(X), y)
                    Next y
                    Add 0, 0, StrPhrase(aPhrs), aPhrs, aPhrTypes, "", _
                        0, 0, 0, X, 0, 0, Entered(X)
                End If
                
            Case gPT_FCompiled, gPT_FCompiledAction, gPT_FInternal
                SetPhraseType X, PType
                FCompiled PhrasesArray(X), PhraseTypesArray(X), X
                
            Case gPT_FTradeSense
                SetPhraseType X, PType
                'When called from the editor, we are only interesed in
                'the TradeSense function and not it's codedtext.
                If mCalledFromEditor Then
                    FCompiled PhrasesArray(X), PhraseTypesArray(X), X
                Else
                    FTradeSense PhrasesArray(X), PhraseTypesArray(X), X
                End If
                      
            'Internal Data Arrays used in inputs of functions
            Case gPT_Trades, gPT_Bars, gPT_Variables
                SetPhraseType X, PType
                SetReturnType X, GetReturnType(PType)
                 
            Case Is <> 0
                'text, numbers, and parms handles here...
                SetPhraseType X, PType
                SetReturnType X, GetReturnType(PType)
            
                'Special case: if parent is OF then this is a Bars label (GC, etc.)
                If Parent(X) > 0 Then
                    If PhraseType(Parent(X)) = gPT_Of Then
                        SetPhraseType X, gPT_Bars
                        SetReturnType X, GetReturnType(gPT_Bars)
                    End If
                End If
                
            Case Else
                SetErrorCode 6, X
        End Select
        
    Loop
    
    UpdateReturnTypes
    SetCondActionPtrs
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:BuildTree ", Err.Description
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
            
            If ReturnType(X) = 0 And PhraseError(X) = 0 Then
                
                Select Case PhraseType(X)
                    Case gPT_OuterParens, gPT_Offset
                        If NodeFrom(X) <> 0 Then
                            SetReturnType X, ReturnType(NodeFrom(X))
                        End If
                        
                    'Math operators
                    Case gPT_Mult, gPT_Div, gPT_Add, gPT_Sub
                        RTLeft = ReturnType(NodeFrom(X))
                        RTRight = ReturnType(NodeTo(X))
                        
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
                        RTLeft = ReturnType(NodeFrom(X))
                        RTRight = ReturnType(NodeTo(X))
                        
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
                        
                    Case gPT_If, gPT_Then, gPT_Not
                        SetReturnType X, ReturnType(NodeTo(X))
                    
                    Case gPT_Of
                        SetReturnType X, ReturnType(NodeFrom(X))
                        
                End Select
                
            End If
        Next X
        
    Next Pass
         
    'STAND ALONE INPUT (NOT PART OF FUNCTION)
    'Set it to the return value on the opposite side of the operator (The
    'constant version of the return type)
    Dim y    As Long
    Dim RT   As Byte
    For X = 1 To mLast
        If PhraseType(X) = gPT_Parm Then
            If Parent(X) > 0 Then
                If FID(Parent(X)) <= 0 And NodeFrom(Parent(X)) > 0 Then
                    'Now, set to the constant version of this return
                    'type (number, boolean, text)
                    RT = ReturnType(NodeFrom(Parent(X)))
                    Select Case RT
                        Case gRetSeriesNbr: RT = gRetConstantNbr
                        Case gRetSeriesBoolean: RT = gRetConstantBoolean
                        Case gRetSeriesText: RT = gRetConstantText
                    End Select
                    gdSetNum aReturnType, X, RT
                End If
            End If
        End If
    Next X
    
    'Look for duplicate parms (and not an input to a function)...
    'If dups found and the dup is an input to a function, change this
    'input outside of the function to the return type of the function input.
    For X = 1 To mLast
        If PhraseType(X) = gPT_Parm Then
            For y = 1 To mLast
                If X <> y And Phrase(X) = Phrase(y) And FID(Parent(y)) > 0 Then
                    gdSetNum aReturnType, X, ReturnType(y)
                    Exit For
                End If
            Next y
        End If
    Next X
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:UpdateReturnTypes ", Err.Description
    End If
End Sub

'Create 2 arrays for pointing to the condition and action phrases
Private Sub SetCondActionPtrs()
On Error GoTo ErrSection:
    Dim X           As Long
    Dim NextSeq     As Long
    
    'Check for Enter key (comments will be stripped out at this point)...
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            mCondPtr.Add NodeFrom(X)
            mActionPtr.Add NodeTo(NodeTo(X)) '<Compiled action
        Else
            Exit For
        End If
    Next X
    
    'Set the execution sequence by assign unique numbers by rules to each node.  Save
    'the very bottom node of each rule in the tree...
    For X = 1 To mCondPtr.Last
        NextSeq = NextSeq + 1
        mMaxNode = -999999
        SetExecutionSequence mCondPtr.Item(X), NextSeq
        mCondTreeBottom.Add mMaxNode
        
        '(The Parent points to the THEN...Same as mActionptr.Item(x))
        NextSeq = NextSeq + 1
        mMaxNode = -999999
        SetExecutionSequence Parent(mActionPtr.Item(X)), NextSeq
        mActionTreeBottom.Add mMaxNode
    Next X
        
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:SetCondActionPtrs ", Err.Description
    End If
End Sub

'This routine assigns a unqiue value to each node by rule.  For example, all phrase
'nodes in Rule 1 get 1, all phrase nodes in rule 2 get 2, etc.  When phrases are
'evaluated in cEvaluator, nodes must be executed by sequence because users put rules
'in a specified sequence.
Private Sub SetExecutionSequence(pStartNode As Long, pNextSeq As Long)
On Error GoTo ErrSection:
    Dim y           As Long
    
    SetRuleID pStartNode, pNextSeq
    If pStartNode > mMaxNode Then
        mMaxNode = pStartNode
    End If
    
    If NodeFrom(pStartNode) > 0 Then
        For y = NodeFrom(pStartNode) To NodeTo(pStartNode)
            SetExecutionSequence y, pNextSeq
        Next y
    End If
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:SetExecutionSequenceNumbers ", Err.Description
    End If
End Sub

Private Sub AdjustParmReturnTypes(pX As Long)
On Error GoTo ErrSection:
    Dim X               As Long
    Dim Parent          As Long
    Dim wReturnType     As Long
    
    If PhraseType(pX) = gPT_Parm Then
        
        Parent = gdGetNum(aParent, pX)
        If Parent > 0 Then
            'If parm is to a function input then set the return type to that
            'of the function input
            If InputID(pX) > 0 Then
                SetReturnType pX, mFunctions.Item(CStr(FID(Parent))).Inputs.Item(CStr(InputID(pX))).ParmTypeID
                GoTo ErrExit
            End If
                    
            Select Case PhraseType(Parent)
                    
                Case gPT_Mult, gPT_Div, gPT_Add, gPT_Sub, _
                     gPT_And, gPT_Or, gPT_GT, gPT_GE, gPT_LT, gPT_LE, gPT_NE, _
                     gPT_EQ
                    
                    'Check the opposite of the operator.  Set to parm return
                    'type to the series version of the return type.
                    If PhraseType(NodeFrom(Parent)) <> gPT_Parm Then
                        wReturnType = ReturnType(NodeFrom(Parent))
                    Else
                        wReturnType = ReturnType(NodeTo(Parent))
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
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:AdjustParmReturnTypes ", Err.Description
    End If
End Sub

'Must do this after building the Tree (routine: BuildTree)
Public Sub BuildTreeCleanup()
On Error GoTo ErrSection:
    Dim X       As Long
    
    'Free up any extra memory at end of arrays...
    gdFreeExtra aPhraseType
    gdFreeExtra aReturnType
    gdFreeExtra aPhrase
    gdFreeExtra aClassName
    gdFreeExtra aClassPtr
    gdFreeExtra aLate
    gdFreeExtra aSize
    gdFreeExtra aNodeFrom
    gdFreeExtra aNodeTo
    gdFreeExtra aParent
    gdFreeExtra aRuleID
    gdFreeExtra aPoolAddr
    gdFreeExtra aParmList
    gdFreeExtra aRebuildArray
    gdFreeExtra aFID
    gdFreeExtra aInputID
    gdFreeExtra aReq
    gdFreeExtra aEntered
    gdFreeExtra aError
    
    'Release phrases and phrase types work arrays
    For X = 1 To mLast
        gdDestroyArray PhrasesArray(X)
        gdDestroyArray PhraseTypesArray(X)
        SetPhrasesArray X, 0
        SetPhraseTypesArray X, 0
    Next X
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
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
        Case gPT_Text
            GetReturnType = gRetConstantText
        Case gPT_Trades
            GetReturnType = gRetTrades
        Case gPT_Bars
            GetReturnType = gRetBars
        Case gPT_Variables
            GetReturnType = gRetVariables
    End Select
End Function

Private Sub FTradeSense(ByVal pPhrases As Long, _
    ByVal pPhraseTypes As Long, pCurrRec As Long)
On Error GoTo ErrSection:
    Dim InputPhrases            As cVectorLng
    Dim InputPhraseTypes        As cVectorLng
    Dim CTPhrases               As Long
    Dim CTPhraseTypes           As Long
    Dim AdjPhrases              As Long
    Dim AdjPhraseTypes          As Long
    Dim ICnt                    As Long
    Dim FunctionID              As Long
    Dim X                       As Long
    Dim curInput                As Long
    Dim curFName                As String
    Dim y                       As Long
    Dim P                       As Long
    Dim F                       As cFunction
    Dim ErrorCode                As Byte
    Dim FICnt                   As Long
    
    'Function name will be stored in element 1 if called from Editor.  Other
    'wise the Function ID will be stored (codedtext functions).
    If Val(gdGetStr(pPhrases, 1)) <> 0 Then
        FunctionID = gdGetStr(pPhrases, 1)
        Set F = mFunctions.Item(CStr(FunctionID))
    Else
        curFName = gdGetStr(pPhrases, 1)
        If mFunctions.NameFound(curFName, FunctionID) Then
            Set F = mFunctions.Item(CStr(FunctionID))
        Else
            Err.Raise gUserErr, , "Function: " & gdGetStr(pPhrases, 1) & " was not found."
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
    SetEntered pCurrRec, Entered(pCurrRec)
    
    'Get any phrases entered for this function...
    'InputPhrases and InputPhraseTypes contain addresses to gdArrays which hold the
    'entered phrases and phrase types.
    GetFunctionInputs pCurrRec, pPhrases, pPhraseTypes, _
            InputPhrases, InputPhraseTypes, ICnt
    
    'Retrieve codedtext of this function
    ConvertCT F.CodedText, CTPhrases, CTPhraseTypes
    
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
        
        'Search through embedded coded text looking for inputs (gPT_Parm)...
        'For X = 1 To gdGetSize(CTPhraseTypes) - 1
        '    If gdGetNum(CTPhraseTypes, X) = gPT_Parm Then
        '        curInput = curInput + 1
        '        If curInput <= InputPhraseTypes.Last Then
        '            For Y = 1 To gdGetSize(InputPhraseTypes.Item(curInput)) - 1
        '                p = p + 1
        '                gdSetStr AdjPhrases, p, gdGetStr(InputPhrases.Item(curInput), Y)
        '                gdSetNum AdjPhraseTypes, p, _
        '                    gdGetNum(InputPhraseTypes.Item(curInput), Y)
        '            Next Y
        '        End If
        '    Else
        '        p = p + 1
        '        gdSetStr AdjPhrases, p, gdGetStr(CTPhrases, X)
        '        gdSetNum AdjPhraseTypes, p, gdGetNum(CTPhraseTypes, X)
        '    End If
        'Next X
        

        '========================== 4/3/2001
        'Everything is positional.  For example,
        '
        '   Rule: IF Win Pct (5)...
        '   TradeSense Coded Text:  Occurrence (Profit>0, Trades Back,1)/Trades Back
        '
        'In the example, "5" must be inserted in place of "Trades Back", which occurs
        '2 times.  5 positionally alligns with "Trades Back", but not with the 2nd
        'occurrence of it.  Therefore, the 2nd occurrence is changed in the CT arrays
        'right after it is updated in the 1st occurrence so it's not found as an input
        'later...
        Dim CTParms     As cVectorStr
        Dim ParmName    As String
        Dim Fnd         As Boolean
        Dim Z           As Long
        Set CTParms = New cVectorStr
        For X = 1 To gdGetSize(CTPhraseTypes) - 1
            If gdGetNum(CTPhraseTypes, X) = gPT_Parm Then
            
                'Get position of this parm so it can be looked up in InputPhrases...
                ParmName = gdGetStr(CTPhrases, X)
                
                'Lookup parm in CTParms array...
                Fnd = False
                For y = 1 To CTParms.Last
                    If CTParms.Item(y) = ParmName Then
                        Fnd = True
                        Exit For
                    End If
                Next y
                
                If Not Fnd Then
                    CTParms.Add ParmName
                    y = CTParms.Last
                End If
                
                'This looping accounts for expressions entered for inputs...
                For Z = 1 To gdGetSize(InputPhrases.Item(y)) - 1
                    P = P + 1
                    gdSetStr AdjPhrases, P, gdGetStr(InputPhrases.Item(y), Z)
                    gdSetNum AdjPhraseTypes, P, gdGetNum(InputPhraseTypes.Item(y), Z)
                Next Z
            Else
                P = P + 1
                gdSetStr AdjPhrases, P, gdGetStr(CTPhrases, X)
                gdSetNum AdjPhraseTypes, P, gdGetNum(CTPhraseTypes, X)
            End If
        Next X
        '========================== 4/3/2001
        
        'Add new node with tradeSense coded text with parm phrases overlaid.
        Add 0, 0, StrPhrase(AdjPhrases), AdjPhrases, _
            AdjPhraseTypes, "", 0, 0, 0, pCurrRec, 0, 0, 0
        
        'Cleanup
        gdDestroyArray CTPhrases
        gdDestroyArray CTPhraseTypes
    Else
        'No inputs.  Post the Codedtext from sub function
        Add 0, 0, StrPhrase(CTPhrases), CTPhrases, _
            CTPhraseTypes, "", 0, 0, 0, pCurrRec, 0, 0, 0
    End If
      
    'Cleanup
    Set F = Nothing
    For X = 1 To InputPhrases.Last
        gdDestroyArray InputPhrases.Item(X)
        gdDestroyArray InputPhraseTypes.Item(X)
    Next X
    Set InputPhrases = Nothing
    Set InputPhraseTypes = Nothing
    Set CTParms = Nothing
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
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
                            pCurrRec As Long)
On Error GoTo ErrSection:
    Dim InputPhrases            As cVectorLng
    Dim InputPhraseTypes        As cVectorLng
    Dim Phr                     As Long
    Dim PhrTypes                As Long
    Dim EnteredInputs           As Long
    Dim X                       As Long
    Dim y                       As Long
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
    
    'Function name will be stored in element 1 if called from Editor.  Other
    'wise the Function ID will be stored (codedtext functions).
    If Val(gdGetStr(pPhrases, 1)) <> 0 Then
        FunctionID = gdGetStr(pPhrases, 1)
        Set F = mFunctions.Item(CStr(FunctionID))
    Else
        curFName = gdGetStr(pPhrases, 1)
        If mFunctions.NameFound(curFName, FunctionID) Then
            Set F = mFunctions.Item(CStr(FunctionID))
        Else
            Err.Raise gUserErr, , "Function: " & gdGetStr(pPhrases, 1) & " was not found."
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
    GetFunctionInputs pCurrRec, pPhrases, pPhraseTypes, _
            InputPhrases, InputPhraseTypes, EnteredInputs
    
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
    For y = 1 To FICnt
        If F.Inputs.Item(y).ParmTypeUsage <> gInternalData Then
            EnterableInputs = EnterableInputs + 1
        End If
    Next y
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
            gdSetStr Phr, 1, curInputName
            Select Case curReturnTypeID
                Case gRetTrades: gdSetNum PhrTypes, 1, gPT_Trades
                Case gRetBars: gdSetNum PhrTypes, 1, gPT_Bars
                Case gRetVariables: gdSetNum PhrTypes, 1, gPT_Variables
            End Select
            Add 0, 0, curInputName, Phr, PhrTypes, "", 0, _
                0, 0, pCurrRec, 0, 0, 0
        Else
            'Nothing entered for input, look for default value from
            'input in the table.  Account for expression in default of inputs (future).
            If gdGetSize(InputPhraseTypes.Item(X - InternalDataInputs)) - 1 <= 0 Then
                Phr = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
                PhrTypes = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
                
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
                
                Add 0, curReturnTypeID, curDefaultValue, Phr, PhrTypes, "", 0, _
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
                        "", 0, 0, 0, pCurrRec, 0, curReq, Entered(pCurrRec)
                Else
                    Add 0, 0, StrPhrase(InputPhrases.Item(X)), _
                        InputPhrases.Item(X), _
                        InputPhraseTypes.Item(X), _
                        "", 0, 0, 0, pCurrRec, 0, curReq, Entered(pCurrRec)
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
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:FCompiled ", Err.Description
    End If
End Sub

'Strip out and save each input of a function as an array of inputs.  If no
'input existing between commas, then an array of 1 element is returned with
'a blank value for the phrase and 0 for the phrase type.
Private Sub GetFunctionInputs(pCurArrayElement As Long, pPhrases As Long, _
    pPhraseTypes As Long, _
    pInputPhrases As cVectorLng, pInputPhraseTypes As cVectorLng, _
    pEnteredInputs As Long)
On Error GoTo ErrSection:
    Dim Fnd             As Boolean
    Dim N               As Long
    Dim X               As Long
    Dim curType         As Byte
    Dim curPhrase       As String
    Dim Parens          As Integer
    Dim Last            As Long
    Dim curPhrases      As Long
    Dim curPhraseTypes  As Long
    Dim TotPhrases      As Long
    
    'Store address of phrase arrays
    Set pInputPhrases = New cVectorLng
    Set pInputPhraseTypes = New cVectorLng
    TotPhrases = gdGetSize(pPhrases) - 1
    
    'gdArray for first input
    pEnteredInputs = 1
    curPhrases = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
    curPhraseTypes = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
    pInputPhrases.Add curPhrases
    pInputPhraseTypes.Add curPhraseTypes
    
    'Get End of Function...
    For X = TotPhrases To 1 Step -1
        If gdGetNum(pPhraseTypes, X) = gPT_FRParen Then
            Last = X
            Exit For
        End If
    Next X
    
    'Check for "Junk" to right of FRight paren.  This could happen during
    'editing.  Return an Error.
    If Last < TotPhrases And TotPhrases > 1 Then
        If gdGetNum(pPhraseTypes, Last + 1) <> gPT_Offset Then
            SetErrorCode 2, pCurArrayElement
        End If
    End If
    
    'This function has no inputs
    If 3 > Last - 1 Then
        pEnteredInputs = 0
        Exit Sub
    End If
    
    '3=Pull inputs between function's left and right paren.  Don't process
    'the left and right paren of the function.
    For X = 3 To Last - 1
        curType = gdGetNum(pPhraseTypes, X)
        curPhrase = gdGetStr(pPhrases, X)
        
        Select Case curType
            Case gPT_FLParen, gPT_LeftPar
                Parens = Parens + 1
                N = N + 1
                gdSetStr curPhrases, N, curPhrase
                gdSetNum curPhraseTypes, N, curType
                
            Case gPT_FRParen, gPT_RightPar
                Parens = Parens - 1
                N = N + 1
                gdSetStr curPhrases, N, curPhrase
                gdSetNum curPhraseTypes, N, curType
                
            Case gPT_Comma
                If Parens = 0 Then
                    'Phrase(s) have been entered
                    If N > 0 Then
                        pEnteredInputs = pEnteredInputs + 1
                        gdFreeExtra curPhrases
                        gdFreeExtra curPhraseTypes
                        curPhrases = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
                        curPhraseTypes = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
                        pInputPhrases.Add curPhrases
                        pInputPhraseTypes.Add curPhraseTypes
                        N = 0
                    Else
                        'Phrase(s) have NOT been entered.
                        'gdSetStr curPhrases, 1, curPhrase
                        'gdSetNum curPhraseTypes, 1, curType
                        pEnteredInputs = pEnteredInputs + 1
                        gdFreeExtra curPhrases
                        gdFreeExtra curPhraseTypes
                        curPhrases = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
                        curPhraseTypes = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
                        pInputPhrases.Add curPhrases
                        pInputPhraseTypes.Add curPhraseTypes
                    End If
                Else
                    'Comma is part of an embedded fuction input list
                    N = N + 1
                    gdSetStr curPhrases, N, curPhrase
                    gdSetNum curPhraseTypes, N, curType
                End If
                
            Case Else
                'Build gdArray of input phrases...
                N = N + 1
                gdSetStr curPhrases, N, curPhrase
                gdSetNum curPhraseTypes, N, curType
                
        End Select
    Next X
    
    gdFreeExtra curPhrases
    gdFreeExtra curPhraseTypes
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
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
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
    Else
        Err.Raise Err.Number, eClass & " Function:StrPhrase ", Err.Description
    End If
End Function

Private Sub Class_Initialize()
On Error GoTo ErrSection:

    aPhraseType = gdCreateArray(eGDARRAY_TinyInts, 2000, USE_DEFAULT_NULL)
    aReturnType = gdCreateArray(eGDARRAY_TinyInts, 2000, USE_DEFAULT_NULL)
    aPhrase = gdCreateArray(eGDARRAY_Strings, 2000, USE_DEFAULT_NULL)
    aPhrasesArray = gdCreateArray(eGDARRAY_Longs, 2000, USE_DEFAULT_NULL)
    aPhraseTypesArray = gdCreateArray(eGDARRAY_Longs, 2000, USE_DEFAULT_NULL)
    aClassName = gdCreateArray(eGDARRAY_Strings, 2000, USE_DEFAULT_NULL)
    aClassPtr = gdCreateArray(eGDARRAY_Longs, 2000, USE_DEFAULT_NULL)
    aLate = gdCreateArray(eGDARRAY_TinyInts, 2000, USE_DEFAULT_NULL)
    aSize = gdCreateArray(eGDARRAY_Longs, 2000, USE_DEFAULT_NULL)
    aNodeFrom = gdCreateArray(eGDARRAY_Longs, 2000, USE_DEFAULT_NULL)
    aNodeTo = gdCreateArray(eGDARRAY_Longs, 2000, USE_DEFAULT_NULL)
    aParent = gdCreateArray(eGDARRAY_Longs, 2000, USE_DEFAULT_NULL)
    aRuleID = gdCreateArray(eGDARRAY_Longs, 2000, USE_DEFAULT_NULL)
    aPoolAddr = gdCreateArray(eGDARRAY_Longs, 2000, USE_DEFAULT_NULL)
    aParmList = gdCreateArray(eGDARRAY_Longs, 2000, USE_DEFAULT_NULL)
    aRebuildArray = gdCreateArray(eGDARRAY_TinyInts, 2000, USE_DEFAULT_NULL)
    aFID = gdCreateArray(eGDARRAY_Longs, 2000, USE_DEFAULT_NULL)
    aInputID = gdCreateArray(eGDARRAY_Longs, 2000, USE_DEFAULT_NULL)
    aReq = gdCreateArray(eGDARRAY_TinyInts, 2000, USE_DEFAULT_NULL)
    aEntered = gdCreateArray(eGDARRAY_TinyInts, 2000, USE_DEFAULT_NULL)
    aError = gdCreateArray(eGDARRAY_TinyInts, 2000, USE_DEFAULT_NULL)
    
    'Reserve 2000 spots for each data item
    gdReserve aPhraseType, 2000, 0
    gdReserve aReturnType, 2000, 0
    gdReserve aPhrase, 2000, 0
    gdReserve aPhrasesArray, 2000, 0
    gdReserve aPhraseTypesArray, 2000, 0
    gdReserve aClassName, 2000, 0
    gdReserve aClassPtr, 2000, 0
    gdReserve aLate, 2000, 0
    gdReserve aSize, 2000, 0
    gdReserve aNodeFrom, 2000, 0
    gdReserve aNodeTo, 2000, 0
    gdReserve aParent, 2000, 0
    gdReserve aRuleID, 2000, 0
    gdReserve aPoolAddr, 2000, 0
    gdReserve aParmList, 2000, 0
    gdReserve aRebuildArray, 2000, 0
    gdReserve aFID, 2000, 0
    gdReserve aInputID, 2000, 0
    gdReserve aReq, 2000, 0
    gdReserve aEntered, 2000, 0
    gdReserve aError, 2000, 0
    
    mLast = 0
    
    Set mCondPtr = New cVectorLng
    Set mActionPtr = New cVectorLng
    Set mCondTreeBottom = New cVectorLng
    Set mActionTreeBottom = New cVectorLng
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:Initialize ", Err.Description
    End If
End Sub

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
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:ValidateCodedText ", Err.Description
    End If
End Sub

'Convert the coded text string to a phrases and phrasetypes array
Private Sub ConvertCT(ByVal pCodedText As String, _
    pPhrasesArray As Long, pPhraseTypesArray As Long)
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
    
    'Create array to storing phrases and phrase types of coded text
    pPhrasesArray = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
    pPhraseTypesArray = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
    
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
            If PType = gPT_Offset Then
                
                'Get offset information
                OffsetValTokenPos = InStr(NextTokenPos + 1, pCodedText, "~")
                PType = Val(Mid(pCodedText, OffsetValTokenPos + 1, 2))
                PhraseLen = Val(Mid(pCodedText, OffsetValTokenPos + 3, 3))
                OffsetVal = Mid(pCodedText, OffsetValTokenPos + 6, PhraseLen)
                If OffsetVal <> "0" Then
                    X = X + 1
                    gdSetNum pPhraseTypesArray, X, gPT_Offset
                    gdSetStr pPhrasesArray, X, "."
                    X = X + 1
                    gdSetNum pPhraseTypesArray, X, PType
                    gdSetStr pPhrasesArray, X, OffsetVal
                End If
                NextTokenPos = OffsetValTokenPos
            Else
                If PType <> gPT_If Then
                    X = X + 1
                    gdSetNum pPhraseTypesArray, X, PType
                    gdSetStr pPhrasesArray, X, curPhrase
                End If
            End If
            
            NextTokenPos = NextTokenPos + 1
        End If
    Loop
    
    gdFreeExtra pPhrasesArray
    gdFreeExtra pPhraseTypesArray
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
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
        Case gPT_FCompiled: GetPhraseTypeDesc = "Compiled"
        Case gPT_FCompiledAction: GetPhraseTypeDesc = "Comp Act"
        Case gPT_FInternal: GetPhraseTypeDesc = "Internal"
        Case gPT_FTradeSense: GetPhraseTypeDesc = "TradeSense"
        Case gPT_Parm: GetPhraseTypeDesc = "Input"
        Case gPT_OuterParens: GetPhraseTypeDesc = "Parens"
        Case gPT_Trades: GetPhraseTypeDesc = "Trades"
        Case gPT_Bars: GetPhraseTypeDesc = "Bars"
        Case gPT_Variables: GetPhraseTypeDesc = "Variables"
        Case Else:
            GetPhraseTypeDesc = "None"
    End Select
    
ErrExit:
    Exit Function
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, , Err.Description
    Else
        Err.Raise Err.Number, eClass & " Sub:GetPhraseTypeDesc ", Err.Description
    End If
End Function

Private Sub Class_Terminate()
    Dim X   As Long
    
    'IMPORTANT!!!
    'Make sure the Trades, gdArrays are not
    'destroyed.  Also, make sure any Internal types (gPT_Internal) are not destroyed.
    'They will be destroyed in Trades.  Skip these.  The Trades gdArrays
    'were being destroyed here.  When this class was called a subsequent
    'time, it caused VB to crash since it's address was lost.
    Dim PoolAddress As Long
    For X = 1 To mLast
        PoolAddress = gdGetNum(aPoolAddr, X)
        If PoolAddress <> 0 And _
           PoolAddress <> mTradesAddr And _
           gdGetNum(aPhraseType, X) <> gPT_FInternal Then
            gdDestroyArray PoolAddress
        End If
        If gdGetNum(aParmList, X) <> 0 Then
            gdDestroyArray gdGetNum(aParmList, X)
        End If
    Next X
    
    gdDestroyArray aPhraseType
    gdDestroyArray aReturnType
    gdDestroyArray aPhrase
    gdDestroyArray aPhrasesArray
    gdDestroyArray aPhraseTypesArray
    gdDestroyArray aClassName
    gdDestroyArray aClassPtr
    gdDestroyArray aLate
    gdDestroyArray aSize
    gdDestroyArray aNodeFrom
    gdDestroyArray aNodeTo
    gdDestroyArray aParent
    gdDestroyArray aRuleID
    gdDestroyArray aFID
    gdDestroyArray aInputID
    gdDestroyArray aReq
    gdDestroyArray aEntered
    gdDestroyArray aError
    gdDestroyArray aPoolAddr
    gdDestroyArray aParmList
    gdDestroyArray aRebuildArray
    
    Set mCondPtr = Nothing
    Set mActionPtr = Nothing
    Set mCondTreeBottom = Nothing
    Set mActionTreeBottom = Nothing
    
End Sub

