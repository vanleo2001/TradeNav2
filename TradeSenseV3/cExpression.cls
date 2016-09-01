VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cExpression.cls
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
'' 04/17/2009   DAJ         Handle string constants better and stop comments
''                          from "walking" when there is an enter in expression
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text
'This class is responsible for validating a condition or action.

'Work variables
Private mEnterKey           As cEnterKey
Private mStrh               As cStrh
Private mFunctions          As cFunctions
Private mText               As String
Private mReturnTypes        As cParmTypes
Private mNextInput          As Long
Private mPortNavigator      As Boolean  'True=Called from Portfolio Navigator
                                        'Use Function ID's in codedtext.  If
                                        'False then use Function Names.

Private mEngineCodedtext    As Boolean  'True=Engine codedtext being created
Private mIfCnt              As Long
Private mEndIfCnt           As Long
Private mElseIfCnt          As Long
Private mThenCnt            As Long
Private mPrevPT             As Byte
Private mCommonDLL          As String
Private mFIDs               As cGdArray

'Returned to Caller
Private mEditText           As String
Private mPreview            As String
Private mCodedText          As String
Private mInputs             As Object
Private mLateCondition      As Boolean  'For compatibility with SysNav
Private mLateAction         As Boolean  'For compatibility with SysNav

'Trees database.  Addresses to phrases defining the rules.
Private aPhraseType         As Long
Private aReturnType         As Long
Private aPhrase             As Long
Private aFormatting         As Long
Private aNodeFrom           As Long
Private aNodeTo             As Long
Private aParent             As Long
Private aFID                As Long     'Function ID (if function)
Private aClassName          As Long
Private aInputID            As Long     'Input ID (if input of function)
Private aReq                As Long     '1=Function Input is required
Private aEntered            As Long     '1=Phrase was user entered, not defaulted
Private aError              As Long     '1=Phrase is error
Private mLast               As Long     'Total Nodes

'gdArray's containing the phrases in the rule
Private mPhrases            As Long
Private mPhraseTypes        As Long
Private mFormatting         As Long     'Store user formatting information
                                        'about each phrase

'tjr 2/03 - added to address comments '{}'
'note - this is not the best way to handle comments
'       but the above phrases in an array are modified in buildtrees - codedtext
'theory here is to store the comments and their original node
'then upon building formatted text, we will place the comments back in and tokenize
Private mComments           As cGdArray
Private commentNode         As Integer

Private mTrees              As cDbTrees

Private Type mPrivate
    astrMacros As cGdArray
End Type
Private m As mPrivate

Private Const gUserErr = vbObjectError + 1000
Private Const mClass = "Class:cExpression "

Property Let Functions(pData As cFunctions)
    Set mFunctions = pData
End Property
Property Let PortfolioNavigator(pData As Boolean)
    mPortNavigator = pData
    If mPortNavigator Then
        mCommonDLL = "PNCommon."
    Else
        mCommonDLL = "NavCommon2."
    End If
End Property

Property Get EditText() As String
    EditText = Trim(mEditText)
End Property
Property Get Preview() As String
    Preview = Trim(mPreview)
End Property
Property Get CodedText() As String
    CodedText = Trim(mCodedText)
End Property
Property Get Inputs() As Object
    Set Inputs = mInputs
End Property
Property Let Inputs(NewInputs As Object)
    Set mInputs = NewInputs
End Property

Property Get Trees() As cDbTrees
    Set Trees = mTrees
End Property

'Should phase out these properties.  Use mainly so System Navigator will
'work
Property Get LateCondition() As Boolean
    LateCondition = mLateCondition
End Property
Property Get LateAction() As Boolean
    LateAction = mLateAction
End Property

'Used by Function Mgr.  Search for and return first phrase with return type.
Property Get FunctionReturnType() As Byte
    Dim X       As Integer
    For X = 1 To mLast
        If PhraseType(X) > 0 And ReturnType(X) > 0 Then
            FunctionReturnType = ReturnType(X)
            Exit For
        End If
    Next X
End Property

Private Property Let FunctionReturnType(ByVal nReturnType As Byte)
    Dim X       As Integer
    For X = 1 To mLast
        If PhraseType(X) > 0 And ReturnType(X) > 0 Then
            ReturnType(X) = nReturnType
            Exit For
        End If
    Next X
End Property

'Used by RuleMgr in System Navigator.  Returns the Function ID of the
'compiled action used.
Property Get ActionFunctionID() As Long
    Dim X       As Long
    For X = 1 To mLast
        If PhraseType(X) = gPT_FCompiledAction Then
            ActionFunctionID = FID(X)
            Exit For
        End If
    Next X
End Property

'Allows easier syntax for retrieving an element from gdArrays
Property Get PhraseType(ByVal i As Long) As Byte
On Error GoTo ErrEnd
    
    PhraseType = gdGetNum(aPhraseType, i)

ErrExit:
    Exit Property
    
ErrEnd:
    'tjr added 1/03 for single character validation in functions
    PhraseType = 0
    
End Property
Property Get Formatting(ByVal i As Long) As String
    Formatting = gdGetStr(aFormatting, i)
End Property
Property Get ReturnType(ByVal i As Long) As Byte
    ReturnType = gdGetNum(aReturnType, i)
End Property
Private Property Let ReturnType(ByVal i As Long, ByVal nReturnType As Byte)
    gdSetNum aReturnType, i, nReturnType
End Property
Property Get Phrase(ByVal i As Long) As String
    Phrase = gdGetStr(aPhrase, i)
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
Property Get FID(ByVal i As Long) As Long
    FID = gdGetNum(aFID, i)
End Property
Property Get ClassName(ByVal i As Long) As Long
    ClassName = gdGetStr(aClassName, i)
End Property
Property Get InputID(ByVal i As Long) As Long
    InputID = gdGetNum(aInputID, i)
End Property
Property Get Entered(ByVal i As Long) As Byte
    Entered = gdGetNum(aEntered, i)
End Property
Property Get Req(ByVal i As Long) As Byte
    Req = gdGetNum(aReq, i)
End Property
Property Get PhsError(ByVal i As Long) As Byte
    PhsError = gdGetNum(aError, i)
End Property

Private Sub SetError(pPos As Long, pValue As Byte)
    gdSetNum aError, pPos, pValue
End Sub

'Returns the functionID's used in a codedtext string
Property Get GetFIDs() As Long
On Error GoTo ErrSection:

    Dim X               As Long
    Dim Skip            As Boolean
    Dim NextTokenPos    As Long
    Dim PType           As Integer
    Dim PhraseLen       As Long
    Dim curPhrase       As String
    Dim FID             As Long
    Dim Continue        As Boolean
    
    Set mFIDs = New cGdArray
    mFIDs.Create eGDARRAY_Longs

    'Scan through codedtext and find all FunctionID's used (Compiled, Compiled actions,
    'internal, and TradeSense functions)
    NextTokenPos = 1
    Do Until NextTokenPos = 0
        NextTokenPos = InStr(NextTokenPos, mCodedText, "~")
        If NextTokenPos > 0 Then
            PType = Val(Mid(mCodedText, NextTokenPos + 1, 2))
            PhraseLen = Val(Mid(mCodedText, NextTokenPos + 3, 3))
            curPhrase = Mid(mCodedText, NextTokenPos + 6, PhraseLen)
            
            'Added for System Navigator (Function tokens different from
            'new standard)
            Continue = False
            If mPortNavigator Then
                If PType = gPT_FInternal Or _
                   PType = gPT_FCompiled Or _
                   PType = gPT_FCompiledAction Or _
                   PType = gPT_FTradeSense Then
                    Continue = True
                End If
            Else
                If PType = gFUNC_NUMERIC Or _
                   PType = gFUNC_BOOLEAN Or _
                   PType = gFUNC_BOOLEAN_CONSTANT Or _
                   PType = gFUNC_TEXT_CONSTANT Or _
                   PType = gFUNC_NUMERIC_CONSTANT Then
                    Continue = True
                End If
            End If
            
            If Continue Then
                
                 'Get function ID of function name.  This was added to
                 'accomodate System Navigator codedtext.
                If mPortNavigator Then
                    FID = CLng(curPhrase)
                Else
                    'System navigator has function name stored, not ID.
                    'If mFunctions.NameFound(curPhrase, FID) Then
                    If CodedNameFound(curPhrase, FID) Then
                    Else
                        Err.Raise gUserErr, , "Function: " & UCase(curPhrase) & " was not found when building function cross-reference list for this rule or function.  Please verifying."
                    End If
                End If
                
                 'Search array of FunctionID's.  Don't add duplicates
                Skip = False
                For X = 0 To mFIDs.Size - 1
                    If FID = mFIDs.Num(X) Then Skip = True
                Next X
                 
                If Not Skip Then
                    mFIDs.Add FID
                End If
                
            End If
            NextTokenPos = NextTokenPos + 1
        End If
    Loop

    GetFIDs = mFIDs.ArrayHandle

ErrExit:
    Exit Property

ErrSection:
    Err.Raise Err.Number, mClass & " Sub:GetFIDs", Err.Description

End Property

'Check for matching parenthesis
Private Sub Chk01(pText As String)
On Error GoTo ErrSection:
    Dim Parens  As Integer
    Dim X       As Integer
    Dim CurChar As String
    Dim CommentLine     As Boolean
    'added to check comments
    Dim iBCnt  As Integer
    Const C_LBRACKET = "{"
    Const C_RBRACKET = "}"
    
    For X = 1 To Len(pText)
        CurChar = Mid(pText, X, 1)
        If CurChar = "'" Then
            CommentLine = True
        Else
            If CommentLine Then
                If CurChar = Chr(13) Then
                    CommentLine = False
                End If
            End If
        End If
        If Not CommentLine Then
            'first we need to count our comments
            If CurChar = C_LBRACKET Then
                iBCnt = iBCnt + 1
            ElseIf CurChar = C_RBRACKET Then
                iBCnt = iBCnt - 1
            End If
            
            'only counts if not between comment brackets
            If iBCnt = 0 Then
                If CurChar = "(" Then
                    Parens = Parens + 1
                Else
                    If CurChar = ")" Then
                        Parens = Parens - 1
                    End If
                End If
            End If
        End If
    Next X
    If Parens <> 0 Then
        Err.Raise gUserErr, , "Please make sure you have matching parenthesis"
    End If
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:Chk01 ", Err.Description
    End If
End Sub

'Check for matching double quotes (for text phrases)
Private Sub Chk02(pText As String)
On Error GoTo ErrSection:
    Dim Quotes  As Integer
    Dim X       As Integer
    Dim CurChar As String
    'added to check comments
    Dim iBCnt  As Integer
    Const C_LBRACKET = "{"
    Const C_RBRACKET = "}"
    
    For X = 1 To Len(pText)
        CurChar = Mid(pText, X, 1)
        'first we need to count our comments
        If CurChar = C_LBRACKET Then
            iBCnt = iBCnt + 1
        ElseIf CurChar = C_RBRACKET Then
            iBCnt = iBCnt - 1
        End If
        
        'only counts if not between comment brackets
        If iBCnt = 0 Then
            If CurChar = """" Then
                Quotes = Quotes + 1
            End If
        End If
    Next X
    If Quotes Mod 2 <> 0 Then
        Err.Raise gUserErr, , "Please make sure all text phrases are enclosed in double quotes."
    End If
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:Chk02 ", Err.Description
    End If
End Sub

'Check for matching comments '{}'
Private Sub ChkComments(pText As String)
On Error GoTo ErrSection:
    Dim iBCnt  As Integer
    Dim X       As Integer
    Dim CurChar As String
    Const C_LBRACKET = "{"
    Const C_RBRACKET = "}"
    
    For X = 1 To Len(pText)
        CurChar = Mid(pText, X, 1)
        If CurChar = C_LBRACKET Then
            iBCnt = iBCnt + 1
        ElseIf CurChar = C_RBRACKET Then
            iBCnt = iBCnt - 1
        End If
    Next X
    If Not iBCnt = 0 Then
        Err.Raise gUserErr, , "Please make sure all comment phrases have matching brackets."
    End If
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:ChkComments ", Err.Description
    End If
End Sub

Private Sub ValCondition(pNode As Long)
On Error GoTo ErrSection:
    Dim X       As Long
    Dim PT      As Byte
    
    If pNode = 0 Then Exit Sub
    
    ValGeneric pNode
    
    PT = PhraseType(pNode)
    
    'Verify Compiled Actions are not used in Conditions expressions.
    If PT = gPT_FCompiledAction Then
        SetError pNode, 1
        Err.Raise gUserErr, , "Action Functions are not allowed in conditions."
    End If
    
    Select Case PT
        Case gPT_FCompiled, gPT_FCompiledAction, _
             gPT_FTradeSense, gPT_FInternal
            
            If NodeFrom(pNode) > 0 Then
                For X = NodeFrom(pNode) To NodeTo(pNode)
                    ValCondition X
                Next X
            End If
            
        Case Else
            If NodeFrom(pNode) > 0 Then
                ValCondition (NodeFrom(pNode))
            End If
            If NodeTo(pNode) > 0 Then
                ValCondition (NodeTo(pNode))
            End If
    End Select
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:ValidateCondition ", Err.Description
    End If
End Sub

'Make sure one action function is specified with nothing to right of it.
Private Sub ValAction(pNode As Long)
On Error GoTo ErrSection:
    Dim X               As Long
    
    If pNode = 0 Then Exit Sub
    
    ValGeneric pNode
    
    Select Case PhraseType(pNode)
        Case gPT_FCompiled, gPT_FCompiledAction, _
             gPT_FTradeSense, gPT_FInternal
            
            If NodeFrom(pNode) > 0 Then
                For X = NodeFrom(pNode) To NodeTo(pNode)
                    ValAction X
                Next X
            End If
            
        Case Else
            If NodeFrom(pNode) > 0 Then
                ValAction (NodeFrom(pNode))
            End If
            If NodeTo(pNode) > 0 Then
                ValAction (NodeTo(pNode))
            End If
    End Select
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:ValAction ", Err.Description
    End If
End Sub
    
Private Sub ValGeneric(pInd As Long)
On Error GoTo ErrSection:
    Dim TmpNum      As Double
    Dim X           As Long
    
    'If this phrase is a function
    If FID(pInd) > 0 Then
    
        'Error - function needs to be reverified
        If mFunctions.Item(CStr(FID(pInd))).Reverify Then
            SetError pInd, 1
            If mFunctions.Item(CStr(FID(pInd))).ImplementationTypeID = gTradeSense Then
                Err.Raise gUserErr, , _
                    "Function: " & UCase(mFunctions.Item(CStr(FID(pInd))).FunctionName) & " needs to be reverified."
            Else
                Err.Raise gUserErr, , _
                    "Function: " & UCase(mFunctions.Item(CStr(FID(pInd))).FunctionName) & " needs to be resaved."
            End If
        End If
        
    End If

    'Numeric values must be between a range...
    If PhraseType(pInd) = gPT_Nbr Then
        'Ignore % sign if part of phrase.  Otherwise "Val" will bomb.
        If Right(Phrase(pInd), 1) = "%" Then
            TmpNum = Left(Phrase(pInd), Len(Phrase(pInd)) - 1)
            TmpNum = TmpNum / 100
        Else
            TmpNum = Val(Phrase(pInd))
        End If
        If TmpNum < -999999999999# Or _
           TmpNum > 999999999999# Then
            SetError pInd, 1
            Err.Raise gUserErr, , _
                "numbers must be between -999 billion to 999 billion"
        End If
    End If
    
    'Parameter names (entered inputs) cannot be longer than 50 characters.
    If PhraseType(pInd) = gPT_Parm Then
        If Len(Phrase(pInd)) > 50 Then
            SetError pInd, 1
            Err.Raise gUserErr, , _
                "Input Parameter names must be 50 characters for less."
        End If
    End If
    
    'Check for type mismatches with function inputs
    '- Series type inputs cannot have both Series and constant values
    '- Constant type inputs cannot only accept constants
    If InputID(pInd) > 0 Then
        Dim TableRetTypeID  As Byte
        Dim FunctionID      As Long
        Dim InpID           As Long
        Dim ErrFnd          As Boolean
        
        FunctionID = FID(Parent(pInd))
        InpID = InputID(pInd)
        TableRetTypeID = _
            mFunctions.Item(CStr(FunctionID)).Inputs.Item(CStr(InpID)).ParmTypeID
        
        ErrFnd = False
        Select Case TableRetTypeID
            
            'Numeric series expected.  Number series/constant both valid.
            Case gRetSeriesNbr, gRetSeriesBoolean
                If ReturnType(pInd) = gRetSeriesText Or _
                   ReturnType(pInd) = gRetConstantText Then
                    ErrFnd = True
                End If
                
            'Text series expected.  Text series/constant both valid.
            Case gRetSeriesText, gRetConstantText
                If ReturnType(pInd) <> gRetSeriesText And _
                   ReturnType(pInd) <> gRetConstantText Then
                    ErrFnd = True
                End If
        
        End Select
    
        If ErrFnd Then
            SetError pInd, 1
            Err.Raise gUserErr, , _
                "The value entered is not the type expected.  " & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                "Input Type expected: " & mReturnTypes.Item(CStr(TableRetTypeID)).ParmType & Chr(13) & Chr(10) & _
                "Input Type entered : " & mReturnTypes.Item(CStr(ReturnType(pInd))).ParmType
        End If
    End If
    
    'Phrases with a return type of constant cannot have an offset value.
    If PhraseType(pInd) = gPT_Offset Then
        If ReturnType(NodeFrom(pInd)) = gRetConstantNbr Or _
           ReturnType(NodeFrom(pInd)) = gRetConstantBoolean Or _
           ReturnType(NodeFrom(pInd)) = gRetConstantText Then
            SetError pInd, 1
            Err.Raise gUserErr, , _
                "Inputs that return a constant value cannot use the offset operator ('.')"
        End If
        
        If PhraseType(NodeTo(pInd)) = gPT_Nbr Then
            If Val(Phrase(NodeTo(pInd))) < 0 Then
                SetError NodeTo(pInd), 1
                Err.Raise gUserErr, , _
                    "Offset values must be positive numbers (no decimal points)."
            End If
        End If
    End If
    
    'OF operator not allowed in portfolio rules.
    If PhraseType(pInd) = gPT_Of And mPortNavigator Then
        SetError pInd, 1
        Err.Raise gUserErr, , _
            "OF operator not allowed.  This is only used in System Navigator to associate the function with a specific symbol."
    End If

    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:ValGeneric ", Err.Description
    End If
End Sub

Private Sub ReportAnyBuildErrors()
On Error GoTo ErrSection:
    Dim X       As Long
    
    For X = 1 To mLast
        If PhsError(X) > 0 Then
            Err.Raise gUserErr, , mTrees.ErrMsg
        End If
    Next X
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:ReportAnyBuildErrors ", Err.Description
    End If
End Sub

Private Sub ResetCommentNode()
    commentNode = 0
End Sub

'Due to the engine and the way we parse phrases - this is the most
'non-intrusive way to add comments back into the overall text.
'Make sure to call this method at any point that the node being looked
'at might change in BuildFormatted().
Private Function NextNodeComment() As String
On Error GoTo ErrSection:
    
    Dim Token       As String
    Dim strComment  As String
    
    commentNode = commentNode + 1
    
    'tjr 2/03 - set for comment
    'if there was a comment at this node before - lets put it back in
    strComment = mComments(commentNode)
    If Not strComment = "" Then
        Token = Tokenize(strComment, gPT_BracketComment, Len(strComment)) & " "
    End If
    NextNodeComment = Token


ErrExit:
    Exit Function

ErrSection:

End Function

'Driver to create Edit and Preview codedtext from arrays
Private Sub BuildFormatted(pNode As Long, pFillWords As Boolean, pShow As Boolean)
On Error GoTo ErrSection:
    
    Dim X           As Long
    Dim Y           As Long
    Dim Token       As String
    Dim PT          As Byte
    Dim curPhrase   As String
    Dim Value       As Double
    Dim F           As cFunction
    Dim i           As cInput
    Dim p           As Long
    Dim FillPre     As String
    Dim CurChar     As String
    Dim InsertComma As Boolean
    Dim InputNbr    As Integer
    
    mEngineCodedtext = False
    PT = PhraseType(pNode)
    curPhrase = Phrase(pNode)
    pShow = True
    
    'If phrase was defaulted then ignore
    If Entered(pNode) = 0 Then
        pShow = False
        GoTo ErrExit:
    End If
    
    'Precede input with any "Pre" words
    FillPre = ""
    If InputID(pNode) > 0 Then
    
        p = Parent(pNode)
        Set F = mFunctions.Item(CStr(FID(p)))
        
        If pFillWords Then
        
            'If all inputs are hidden(internal) then quit
            If F.HiddenInputs = F.Inputs.Count Then
            Else
                If InputID(pNode) > 0 And F.Inputs.Count > 0 Then
                
                    Set i = F.Inputs.Item(CStr(InputID(pNode)))
                    
                    If i.FillPre <> "" Then
                        FillPre = i.FillPre
                        curPhrase = FillPre & " " & curPhrase
                    End If
                End If
            End If
        End If
    End If
    
    Select Case PT
        
        'Functions...
        Case gPT_FCompiled, gPT_FCompiledAction, gPT_FTradeSense, gPT_FInternal
            
            'Get function name
            Set F = mFunctions.Item(CStr(FID(pNode)))
            curPhrase = F.FunctionName
            
            ApplyUserFormatting pNode
           
            'Handling Portfolio Navigator with Fill words
            If mPortNavigator Then
                If FillPre <> "" Then
                    curPhrase = FillPre & " " & curPhrase
                End If
                Token = Tokenize(curPhrase, PT, Len(curPhrase))
            
                If PhsError(pNode) > 0 Then
                    Token = Tokenize(curPhrase, gPT_Error, Len(curPhrase))
                End If
                
                mText = mText & NextNodeComment & Token & " "
            Else
                'System Navigator:
                If FillPre = "" Then
                    Token = Tokenize(curPhrase, PT, Len(curPhrase))
                    
                    If PhsError(pNode) > 0 Then
                        Token = Tokenize(curPhrase, gPT_Error, Len(curPhrase))
                    End If
                    
                    mText = mText & NextNodeComment & Token & " "
                Else
                    'Special handling for prefill words in System Navigator
                    'because Function Names are stored in codedtext. CurPhrase
                    'can't have the fill pre fill words included because the
                    'function is not found when this is included
                    Dim SysNavToken As String
                    SysNavToken = "~" & _
                        Format(CToken(PT, curPhrase), "00") & _
                        Format(Len(FillPre & " " & curPhrase), "000")
                    mText = mText & NextNodeComment & SysNavToken & FillPre & " " & curPhrase & " "
                End If
            End If
            
            'Function has one or more inputs...
            If NodeFrom(pNode) > 0 Then
            
                'If all inputs are hidden(internal) then don't show left parens
                'with comma's separating inputs
                If F.HiddenInputs = F.Inputs.Count Then GoTo ErrExit:
                
                Token = Tokenize("(", gPT_FLParen, 1)
                mText = mText & NextNodeComment & Token & " "
                For X = NodeFrom(pNode) To NodeTo(pNode)
                    'Skip hidden inputs and optional inputs that are not
                    'entered
                    InputNbr = InputNbr + 1
                    If Entered(X) Then
                        If InputNbr <= F.Inputs.Count Then
                            If F.Inputs.Item(InputNbr).ParmTypeUsage <> gInternalData Then
                                BuildFormatted X, pFillWords, pShow
                            End If
                        Else
                            BuildFormatted X, pFillWords, pShow
                        End If
                        
                        'If this is an optional input that wasn't entered,
                        'check to see if any other optional inputs were
                        'entered after this one.  If yes, then show the
                        'comma, otherwise don't show it.
                        InsertComma = False
                        For Y = X + 1 To NodeTo(pNode)
                            If Entered(Y) Then
                                InsertComma = True
                                Exit For
                            End If
                        Next Y
                        If InsertComma Then
                            Token = Tokenize(",", gPT_Comma, 1)
                            mText = mText & NextNodeComment & Token & " "
                        End If
                    End If
                    
                Next X
                
                'Attach function right parenthesis
                Token = Tokenize(")", gPT_FRParen, 1)
                mText = mText & NextNodeComment & Token & " "
            End If
        
        Case gPT_OuterParens
            If NodeFrom(pNode) > 0 Then
                ApplyUserFormatting pNode
                Token = Tokenize("(", gPT_LeftPar, 1)
                mText = mText & NextNodeComment & Token & " "
                For X = NodeFrom(pNode) To NodeTo(pNode)
                    BuildFormatted X, pFillWords, pShow
                Next X
                Token = Tokenize(")", gPT_RightPar, 1)
                mText = mText & NextNodeComment & Token & " "
            End If
        
        '(0 means couldn't be determined when building tree nodes)
        Case gPT_Parm, gPT_Nbr, gPT_Text, gPT_Bars, gPT_Comment, 0
            ApplyUserFormatting pNode
            Token = Tokenize(curPhrase, PT, Len(curPhrase))
            
            If PhsError(pNode) > 0 Or PT = 0 Then
                Token = Tokenize(curPhrase, gPT_Error, Len(curPhrase))
            End If
            
            mText = mText & NextNodeComment & Token & " "
        
        Case Else
            If NodeFrom(pNode) > 0 Then
                BuildFormatted NodeFrom(pNode), pFillWords, pShow
            End If
            ApplyUserFormatting pNode
            curPhrase = Symbol(PT)
            Token = Tokenize(curPhrase, PT, Len(curPhrase))
            
            If PhsError(pNode) > 0 Then
                Token = Tokenize(curPhrase, gPT_Error, Len(curPhrase))
            End If
            
            mText = mText & NextNodeComment & Token & " "
            If NodeTo(pNode) > 0 Then
                BuildFormatted NodeTo(pNode), pFillWords, pShow
            End If
            
    End Select
    
ErrExit:
    Set i = Nothing
    Set F = Nothing
    Exit Sub

ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:BuildFormatted ", Err.Description
    End If

End Sub
    
Private Sub ApplyUserFormatting(pNode As Long)
On Error GoTo ErrSection:
    
    Dim CurChar     As String
    Dim Token       As String
    Dim X           As Integer
    
    'If User formatting with the <tab> keys exists, prefix it now...
    For X = 1 To Len(Formatting(pNode))
        CurChar = Mid(Formatting(pNode), X, 1)
        If CurChar = "T" Then
            Token = Tokenize("<Tab>", gPT_Tab, 5)
        Else
            If CurChar = "E" Then
                Token = Tokenize("<EF>", gPT_EnterFormatting, 4)
            End If
        End If
        mText = mText & NextNodeComment & Token & " "
    Next X

ErrExit:
    Exit Sub

ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:ApplyUserFormatting ", Err.Description
    End If

End Sub

'Driver to create codedtext string.  If mPortNav is True then FunctionID's
'are generated in codedtext, otherwise, Function names are generated.
Private Sub BuildCoded(pNode As Long)
On Error GoTo ErrSection:
    
    Dim X           As Long
    Dim Token       As String
    Dim PT          As Byte
    Dim curPhrase   As String
    Dim Value       As Double
    Dim OfSymbol    As String
    Dim Z           As Long
    Dim FName       As String
    
    mEngineCodedtext = True
    PT = PhraseType(pNode)
    curPhrase = Phrase(pNode)
        
    Select Case PT
    
        'Convert Function Name to Function ID
        Case gPT_FCompiled, gPT_FCompiledAction, gPT_FTradeSense, gPT_FInternal
            
            'Tokenize FunctionID and Function Left Paren
            If mPortNavigator Then
                Token = Tokenize(FID(pNode), PT, Len(CStr(FID(pNode))))
            Else
                'System Navigator "NavCommon.cFunction" was modified to
                'move the Codedname to the ClassName to be compatible with
                'the new TradeSense.  This is a temporary fix to maintain
                'Binary compatibility
                FName = mFunctions.Item(CStr(CLng(FID(pNode)))).ClassName
                Token = Tokenize(FName, PT, Len(FName))
            End If
            mText = mText & Token & " "
            Token = Tokenize("(", gPT_FLParen, 1)
            mText = mText & Token & " "
            
            'Function with no inputs (include empty parens)
            If NodeFrom(pNode) = 0 Then
                Token = Tokenize(")", gPT_FRParen, 1)
                mText = mText & Token & " "
            Else
                'Separate inputs by commas
                For X = NodeFrom(pNode) To NodeTo(pNode)
                    BuildCoded X
                    If X <> NodeTo(pNode) Then
                        Token = Tokenize(",", gPT_Comma, 1)
                        mText = mText & Token & " "
                    End If
                Next X
                Token = Tokenize(")", gPT_FRParen, 1)
                mText = mText & Token & " "
            End If

        Case gPT_Bars
            ' TLB 3/11/05: don't replace if within double-quotes
            'If Left(curPhrase, 1) = Chr(34) And Right(curPhrase, 1) = Chr(34) Then
            ' TLB 4/13/05: now only allow replacing if it's "Market1"
            If UCase(curPhrase) <> "MARKET1" Then
                OfSymbol = curPhrase
            Else
                'Search back up tree looking for the most recent use of the
                'OF operator.  If not found, use Input #1 Defalt Value (Market1, GC, etc)
                OfSymbol = mFunctions.Item(CStr(FID(Parent(pNode)))).Inputs.Item(1).DefaultValue
            
                'Do this if the symbol after 'OF' is a function
                Z = pNode
                Do Until Parent(Z) = 0
                    If PhraseType(Parent(Z)) = gPT_Of Then
                        OfSymbol = Phrase(NodeTo(Parent(Z)))
                        Exit Do
                    End If
                    Z = Parent(Z)
                Loop
            End If
            
            Token = Tokenize(OfSymbol, gPT_Bars, Len(OfSymbol))
            mText = mText & Token & " "
        
        Case gPT_Nbr
            If Right(curPhrase, 1) = "%" Then
                Value = Left(curPhrase, Len(curPhrase) - 1)
                curPhrase = CStr(Value / 100)
            End If
            Token = Tokenize(curPhrase, PT, Len(curPhrase))
            mText = mText & Token & " "
        
        Case gPT_Parm
            Token = Tokenize(curPhrase, PT, Len(curPhrase))
            mText = mText & Token & " "
        
            'Parm must be broken down for Tokens in System Navigator
            If Not mPortNavigator Then
                Select Case ReturnType(pNode)
                    Case gRetSeriesNbr
                        Mid(mText, InStrRev(mText, "~", Len(mText)) + 1, 2) = _
                            Format(gPARM_NUMERIC_ARRAY, "00")
                    Case gRetSeriesBoolean
                        Mid(mText, InStrRev(mText, "~", Len(mText)) + 1, 2) = _
                            Format(gPARM_BOOLEAN_ARRAY, "00")
                    Case gRetConstantNbr
                        Mid(mText, InStrRev(mText, "~", Len(mText)) + 1, 2) = _
                            Format(gPARM_NUMERIC, "00")
                    Case gRetConstantBoolean
                        Mid(mText, InStrRev(mText, "~", Len(mText)) + 1, 2) = _
                            Format(gPARM_BOOLEAN, "00")
                    Case gRetSeriesText
                    Case gRetConstantText
                End Select
            End If
            
        Case gPT_Text
            If Len(curPhrase) > 0 Then
                curPhrase = Mid(curPhrase, 2, Len(curPhrase) - 2)
            End If
            Token = Tokenize(curPhrase, PT, Len(curPhrase))
            mText = mText & Token & " "
        
'        Case gPT_BracketComment
'            curPhrase = Mid(curPhrase, 1, Len(curPhrase))
'            Token = Tokenize(curPhrase, PT, Len(curPhrase))
'            mText = mText & Token & " "
        
        Case gPT_OuterParens
            Token = Tokenize("(", gPT_LeftPar, 1)
            mText = mText & Token & " "
            For X = NodeFrom(pNode) To NodeTo(pNode)
                BuildCoded X
            Next X
            Token = Tokenize(")", gPT_RightPar, 1)
            mText = mText & Token & " "
        
        Case gPT_Add, gPT_Sub, gPT_Div, gPT_Mult, gPT_GT, gPT_GE, gPT_LT, _
             gPT_LE, gPT_NE, gPT_And, gPT_Or, gPT_EQ, gPT_Offset
            BuildCoded NodeFrom(pNode)
            Token = Tokenize(Symbol(PT), PT, Len(Symbol(PT)))
            mText = mText & Token & " "
            BuildCoded NodeTo(pNode)
        
        Case gPT_Then
            If NodeFrom(pNode) > 0 Then
                BuildCoded NodeFrom(pNode)
            End If
            Token = Tokenize(Symbol(PT), PT, Len(Symbol(PT)))
            mText = mText & Token & " "
        
        Case gPT_Enter
            'If the Enterkey is for a comment, then don't include comment in
            'the codedtext
            If PhraseType(NodeFrom(pNode)) <> gPT_Comment Then
                If NodeFrom(pNode) > 0 Then
                    BuildCoded NodeFrom(pNode)
                End If
                Token = Tokenize(Symbol(PT), PT, Len(Symbol(PT)))
                mText = mText & Token & " "
            End If
            
        Case gPT_Comment    'Do nothing
        
        'Case gPT_Tab
        '    If NodeTo(pNode) > 0 Then
        '        BuildCoded NodeTo(pNode)
        '    End If
        
        'Case gPT_EnterFormatting
        '    If NodeFrom(pNode) > 0 Then
        '        BuildCoded NodeFrom(pNode)
        '    End If
            
        Case gPT_Of
            BuildCoded NodeFrom(pNode)
            
        Case gPT_If, gPT_Not, gPT_ElseIf, gPT_DoUntil
            Token = Tokenize(Symbol(PT), PT, Len(Symbol(PT)))
            mText = mText & Token & " "
            If NodeTo(pNode) > 0 Then
                BuildCoded NodeTo(pNode)
            End If
        
        Case gPT_EndIf, gPT_Else, gPT_EndDo
            Token = Tokenize(Symbol(PT), PT, Len(Symbol(PT)))
            mText = mText & Token & " "
    
        Case Else
            Token = Tokenize(curPhrase, PT, Len(curPhrase))
            mText = mText & Token & " "
        
    End Select
    
ErrExit:
    Exit Sub

ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:BuildCoded ", Err.Description
    End If

End Sub


'Returns true if the condition part of the rule has at least one late
'calculating function.  This is for compatibility with System Navigator.
Private Sub CalcLateCondition(pNode As Long)
On Error GoTo ErrSection:
    Dim X           As Long
    Dim PT          As Byte
    
    If pNode = 0 Then Exit Sub
    PT = PhraseType(pNode)
    
    Select Case PT
        Case gPT_FCompiled, gPT_FCompiledAction, _
             gPT_FTradeSense, gPT_FInternal
            If mFunctions.Item(CStr(CLng(FID(pNode)))).LateCalculating Then
                mLateCondition = True
                Exit Sub
            End If
            
            If NodeFrom(pNode) > 0 Then
                For X = NodeFrom(pNode) To NodeTo(pNode)
                    CalcLateCondition X
                    If mLateCondition Then Exit For
                Next X
            End If
            
        Case Else
            If NodeFrom(pNode) > 0 Then
                CalcLateCondition (NodeFrom(pNode))
            End If
            If NodeTo(pNode) > 0 Then
                CalcLateCondition (NodeTo(pNode))
            End If
            
    End Select
        
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:LateCondition ", Err.Description
    End If
End Sub

'Returns true if the action part of the rule has at least one late
'calculating function.  This is for compatibility with System Navigator.
Private Sub CalcLateAction(pNode As Long)
On Error GoTo ErrSection:
    Dim X           As Long
    Dim PT          As Byte
    
    If pNode = 0 Then Exit Sub
    PT = PhraseType(pNode)

    Select Case PT
        Case gPT_FCompiled, gPT_FTradeSense, gPT_FInternal
            If mFunctions.Item(CStr(CLng(FID(pNode)))).LateCalculating Then
                mLateAction = True
                Exit Sub
            End If
            
            If NodeFrom(pNode) > 0 Then
                For X = NodeFrom(pNode) To NodeTo(pNode)
                    CalcLateAction X
                    If mLateAction Then Exit For
                Next X
            End If
            
        Case Else
            If NodeFrom(pNode) > 0 Then
                CalcLateAction (NodeFrom(pNode))
            End If
            If NodeTo(pNode) > 0 Then
                CalcLateAction (NodeTo(pNode))
            End If
    End Select

    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:LateAction ", Err.Description
    End If
End Sub

'Returns true if the at least one late calculating function is found.
'This is for compatibility with System Navigator.
Private Sub CalcLateFunction(pNode As Long)
On Error GoTo ErrSection:
    Dim X           As Long
    Dim PT          As Byte
    
    If pNode = 0 Then Exit Sub
    PT = PhraseType(pNode)

    Select Case PT
        Case gPT_FCompiled, gPT_FTradeSense, gPT_FInternal
            If mFunctions.Item(CStr(CLng(FID(pNode)))).LateCalculating Then
                mLateCondition = True
                Exit Sub
            End If
            
            If NodeFrom(pNode) > 0 Then
                For X = NodeFrom(pNode) To NodeTo(pNode)
                    CalcLateFunction X
                    If mLateCondition Then Exit For
                Next X
            End If
            
        Case Else
            If NodeFrom(pNode) > 0 Then
                CalcLateFunction (NodeFrom(pNode))
            End If
            If NodeTo(pNode) > 0 Then
                CalcLateFunction (NodeTo(pNode))
            End If
            
    End Select
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:CalcLateFunction ", Err.Description
    End If
End Sub

'Determine the phrase type for the given characters passed here...
Private Function GetPhraseType(pPhrase As String) As Byte
On Error GoTo ErrSection:
    Dim FunctionID      As Long
    
    If mStrh.IsANumeric(pPhrase) And pPhrase <> "" And pPhrase <> "." Then
        GetPhraseType = gPT_Nbr
        GoTo ErrExit
    End If

    If Left(pPhrase, 2) = "If" And Left(pPhrase, 3) <> "Iff" Then
        GetPhraseType = gPT_If
        GoTo ErrExit
    End If

    If Left(pPhrase, 2) = "DoUntil" Then
        GetPhraseType = gPT_DoUntil
        GoTo ErrExit
    End If
    
    Select Case pPhrase
        Case ",": GetPhraseType = gPT_Comma
        Case ".": GetPhraseType = gPT_Offset
        Case "+": GetPhraseType = gPT_Add
        Case "-": GetPhraseType = gPT_Sub
        Case "*": GetPhraseType = gPT_Mult
        Case "/": GetPhraseType = gPT_Div
        Case ">": GetPhraseType = gPT_GT
        Case ">=": GetPhraseType = gPT_GE
        Case "<": GetPhraseType = gPT_LT
        Case "<=": GetPhraseType = gPT_LE
        Case "<>": GetPhraseType = gPT_NE
        Case "=": GetPhraseType = gPT_EQ
        Case "and": GetPhraseType = gPT_And
        Case "not": GetPhraseType = gPT_Not
        Case "or":  GetPhraseType = gPT_Or
        Case "of": GetPhraseType = gPT_Of
        Case "if": GetPhraseType = gPT_If
        Case "DoUntil": GetPhraseType = gPT_DoUntil
        Case "Else": GetPhraseType = gPT_Else
        Case "Elseif": GetPhraseType = gPT_ElseIf
        Case "then": GetPhraseType = gPT_Then
        Case "Endif": GetPhraseType = gPT_EndIf
        Case "EndDo": GetPhraseType = gPT_EndDo
        Case "(":  GetPhraseType = gPT_LeftPar
        Case ")": GetPhraseType = gPT_RightPar
        Case """": GetPhraseType = gPT_DoubleQuote
        Case "'": GetPhraseType = gPT_Comment
        Case "{": GetPhraseType = gPT_BracketComment
        Case Chr(13) & Chr(10): GetPhraseType = gPT_Enter
        Case Chr(9): GetPhraseType = gPT_Tab
        Case Is <> ""
            If mFunctions.NameFound(CStr(pPhrase), FunctionID) Then
                Select Case mFunctions.Item(CStr(FunctionID)).ImplementationTypeID
                    Case gInternalData: GetPhraseType = gPT_FInternal
                    Case gTradeSense: GetPhraseType = gPT_FTradeSense
                    Case gCompiled: GetPhraseType = gPT_FCompiled
                    Case gCompiledAction: GetPhraseType = gPT_FCompiledAction
                    Case Else
                        Err.Raise gUserErr, , "Function: " & pPhrase & " not found during parsing.  Please report this error."
                End Select
            Else
                'phrase not recognized.  Default to variable (parm).
                GetPhraseType = gPT_Parm
            End If
    End Select
    
'    'Capitalize the first letter
'    'pPhrase = LCase(pPhrase)
'    Mid(pPhrase, 1, 1) = UCase(Left(pPhrase, 1))
    
ErrExit:
    Exit Function
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:GetPhraseType ", Err.Description
    End If
End Function


Private Function Symbol(pPT As Byte) As String
On Error GoTo ErrSection:
    
    Select Case pPT
        Case gPT_Offset: Symbol = "."
        Case gPT_If: Symbol = "IF"
        Case gPT_DoUntil: Symbol = "DoUntil"
        Case gPT_Else: Symbol = "ELSE"
        Case gPT_ElseIf: Symbol = "ELSEIF"
        Case gPT_Then: Symbol = "THEN"
        Case gPT_EndIf: Symbol = "ENDIF"
        Case gPT_EndDo: Symbol = "EndDo"
        Case gPT_Add: Symbol = "+"
        Case gPT_Sub: Symbol = "-"
        Case gPT_Mult: Symbol = "*"
        Case gPT_Div: Symbol = "/"
        Case gPT_GT: Symbol = ">"
        Case gPT_GE: Symbol = ">="
        Case gPT_LT: Symbol = "<"
        Case gPT_LE: Symbol = "<="
        Case gPT_NE: Symbol = "<>"
        Case gPT_EQ: Symbol = "="
        Case gPT_And: Symbol = "And"
        Case gPT_Not: Symbol = "Not"
        Case gPT_Or: Symbol = "Or"
        Case gPT_Of: Symbol = "Of"
        Case gPT_Enter: Symbol = "<E>"
        Case gPT_EnterFormatting: Symbol = "<EF>"
        Case gPT_Tab: Symbol = "<Tab>"
    End Select
    
    Exit Function
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:Symbol ", Err.Description
    End If
End Function

'Validation for Advanced rules using multiple IF's and Action's.  This is
'only used by Portfolio Navigator right now.
Public Sub ValidateRuleAdv(pText As String)
On Error GoTo ErrSection:
    Dim svErr       As Long
    Dim svDesc      As String
    Dim svSource    As String
    Dim X           As Long
    
'-------------------------------------------------------------------------
'Parse text into Phrase and Phrase type gdArrays and then into a Tree...

    mEnterKey.ReFormat2 pText
    mStrh.InsertOneBlank pText
    VerifyBlanksAroundSeparators pText
    pText = RTrim(pText)
    
    'Pre-Array Build checks
    Chk01 pText     'Make sure parenthesis match
    Chk02 pText     'Make sure double quotes (for text) match
    ChkComments pText 'Make sure brackets match

    'Parse pText and build Phrase and Phrase type gdArrays for use in mTrees
    BuildPhrases pText
    
    Chk03   'Differentiate between function parens and regular parens
    
    'For X = 1 To gdGetSize(mPhrases) - 1
    '    Debug.Print gdGetStr(mPhrases, X)
    'Next X
    
    SaveUserFormatting   'Mark <Enterkeys> between IF/Then constucts
    
'-------------------------------------------------------------------------
'Build Tree and editor codedtext

    BuildTree
    
'------------------------------------------------------------------------
'IF/THEN/ELSE constructs checks
    
    ValIfThen01
    ValIfThen03
    ValDoUntil01
    
'---------------------------------------------------------------------------
'Build Codedtext and inputs

    'Build inputs used by Rule
    Set mInputs = CreateObject(mCommonDLL & "cInputs")
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            BuildInputs X
        End If
    Next X
    
    'Ensure Market1 is the first parm in inputs collection
    ReSortInputs
    
    'Build engine codedtext...
    mText = ""
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            BuildCoded X
        End If
    Next X
    mCodedText = mText

    
    'reset the comment engine
    ResetCommentNode
    'Build editor codedtext...
    mText = ""
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            BuildFormatted X, False, True
        End If
    Next X
    mEditText = mText
    
    'reset the comment engine
    ResetCommentNode
    'Create preview codedtext...
    mText = ""
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            BuildFormatted X, True, True
        End If
    Next X
    mPreview = mText
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then

        'Build the formatted text before raising the error.
        svErr = Err.Number
        svDesc = Err.Description
        svSource = Err.Source
        If mLast > 0 Then
            'reset the comment engine
            ResetCommentNode
            mText = ""
            For X = 1 To mLast
                If PhraseType(X) = gPT_Enter Then
                    BuildFormatted X, False, True
                End If
            Next X
            mEditText = mText
        End If
        Err.Raise svErr, svSource, svDesc
        
    Else
        Err.Raise Err.Number, mClass & " Sub:ValidateRule ", Err.Description
    End If
End Sub

'Driver to validate rule
'Keep this method available for System Navigator, Advanced mode which allows
'one IF and one ACTION function
Public Sub ValidateRule(strExpression As String)
On Error GoTo ErrSection:
    
    Dim svErr       As Long
    Dim svDesc      As String
    Dim svSource    As String
    Dim X           As Long
    Dim IFPos       As Long
    Dim IfCnt       As Long
    Dim CACnt       As Long
    Dim CompiledActionPos As Long
    
    Dim astrExpressions As cGdArray     ' Array of expressions
    Dim Expr As cExpression             ' Expression object for assignments
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim strVariable As String           ' Variable name for assignments
    Dim strAssignment As String         ' Expression for assignments
    Dim strFormatted As String          ' Formatted text for the whole expression
    Dim strPreview As String            ' Preview text for the whole expression
    Dim strCodedText As String          ' Coded text for assignments
    Dim strNewCoded As String           ' Coded text for assignments with parenthesis
    Dim strToken As String              ' Token for variable name
    Dim lPos As Long                    ' Position in an array
    Dim strMainExpr As String           ' Main expression
    
'-------------------------------------------------------------------------
'Parse out and handle any macro variable declarations...
    
    ' TLB 9/20/2012: rename "LW Sentiment" functions
    strExpression = Replace(strExpression, "LW Sentiment", "TN Consensus")
    
    ' DAJ 03/20/2009: Need to make sure that there are spaces around any math
    ' operator that isn't in comments or quotes...
    strExpression = SpacesAroundOperators(strExpression)
    
    Set astrExpressions = New cGdArray
    astrExpressions.SplitFields Trim(strExpression), vbLf
    
    Set mInputs = CreateObject(mCommonDLL & "cInputs")
    
    For lIndex = 0 To astrExpressions.Size - 1
        If Len(astrExpressions(lIndex)) = 0 Then
            strFormatted = strFormatted & " ~80003<E>"
            strPreview = strPreview & " ~80003<E>"
        ElseIf InStr(astrExpressions(lIndex), ":=") <> 0 Then
            If Left(astrExpressions(lIndex), 1) <> VARIABLE_PREFIX Then
                astrExpressions(lIndex) = VARIABLE_PREFIX & astrExpressions(lIndex)
            End If
            
            strVariable = Parse(astrExpressions(lIndex), ":=", 1)
            strAssignment = Parse(astrExpressions(lIndex), ":=", 2)
            
            Set Expr = New cExpression
            Expr.PortfolioNavigator = False
            Expr.Functions = mFunctions
            Expr.Inputs = mInputs
            Expr.ValidateFunctionRule strAssignment
            
            AddInputs Expr.Inputs
            
            strFormatted = strFormatted & " ~" & Format(gPARM_NUMERIC, "00") & Format(Len(strVariable), "000") & strVariable
            strFormatted = strFormatted & " ~" & Format(gCOMPARE, "00") & "002:= " & Expr.EditText
            
            strPreview = strPreview & " ~" & Format(gPARM_NUMERIC, "00") & Format(Len(strVariable), "000") & strVariable
            strPreview = strPreview & " ~" & Format(gCOMPARE, "00") & "002:= " & Expr.Preview
            
            strCodedText = Expr.CodedText
                        
            For lIndex2 = 0 To m.astrMacros.Size - 1
                strToken = Format(Len(Parse(m.astrMacros(lIndex2), vbTab, 1)), "000") & Parse(m.astrMacros(lIndex2), vbTab, 1)
                strNewCoded = "~14001( " & Parse(m.astrMacros(lIndex2), vbTab, 3) & " ~15001)"
                
                strCodedText = Replace(strCodedText, "~05" & strToken, strNewCoded)
                strCodedText = Replace(strCodedText, "~06" & strToken, strNewCoded)
                strCodedText = Replace(strCodedText, "~07" & strToken, strNewCoded)
                strCodedText = Replace(strCodedText, "~25" & strToken, strNewCoded)
                strCodedText = Replace(strCodedText, "~27" & strToken, strNewCoded)
                strCodedText = Replace(strCodedText, "~28" & strToken, strNewCoded)
            Next lIndex2
            
            If m.astrMacros.BinarySearch(strVariable & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = False Then
                m.astrMacros.Add strVariable & vbTab & strAssignment & vbTab & strCodedText, lPos
            End If
        ElseIf (Left(Trim(astrExpressions(lIndex)), 1) = "{") And (Right(Trim(astrExpressions(lIndex)), 1) = "}") Then
            strFormatted = strFormatted & " ~" & Format(gCOMMENT, "00") & Format(Len(Trim(astrExpressions(lIndex))), "000") & Trim(astrExpressions(lIndex)) & " ~80003<E>"
            strPreview = strPreview & " ~" & Format(gCOMMENT, "00") & Format(Len(Trim(astrExpressions(lIndex))), "000") & Trim(astrExpressions(lIndex)) & " ~80003<E>"
        End If
    Next lIndex

'-------------------------------------------------------------------------
'Parse text into Phrase and Phrase type gdArrays and then into a Tree...
    
    For lIndex = 0 To astrExpressions.Size - 1
        If (Len(astrExpressions(lIndex)) > 0) And (InStr(astrExpressions(lIndex), ":=") = 0) Then
            If (Left(Trim(astrExpressions(lIndex)), 1) <> "{") Or (Right(Trim(astrExpressions(lIndex)), 1) <> "}") Then
                strMainExpr = strMainExpr & astrExpressions(lIndex) & vbCrLf
            End If
        End If
    Next lIndex
    strMainExpr = Trim(strMainExpr)
    
    mEnterKey.ReFormat strMainExpr
    mStrh.InsertOneBlank strMainExpr
    VerifyBlanksAroundSeparators strMainExpr
    
    strMainExpr = RTrim(strMainExpr)
    
    Chk01 strMainExpr     'Make sure parenthesis match
    Chk02 strMainExpr     'Make sure double quotes (for text) match
    ChkComments strMainExpr 'Make sure brackets match
    
    'Parse strExpression and build Phrase and Phrase type gdArrays for use in mTrees
    BuildPhrases strMainExpr
    
    Chk03   'Differentiate between function parens and regular parens
    
    'For X = 1 To gdGetSize(mPhrases) - 1
    '    Debug.Print gdGetStr(mPhrases, X)
    'Next X
    
    SaveUserFormatting   'Mark <Enterkeys> between IF/Then constucts
    
'-------------------------------------------------------------------------
'Build Tree and phrase array's

    BuildTree
    
'------------------------------------------------------------------------
'IF/THEN/ELSE constructs checks
    
    'Verify that one IF statement is specified
    For X = 1 To mLast
        Select Case PhraseType(X)
            Case gPT_If
                IFPos = X
                IfCnt = IfCnt + 1
                
            Case gPT_FCompiledAction
                CompiledActionPos = X
                CACnt = CACnt + 1
                
            Case gPT_ElseIf, gPT_Else, gPT_DoUntil, gPT_EndDo
                SetError X, 1
                Err.Raise gUserErr, , "ELSEIF, ELSE, and ENDIF are not supported in System Navigator at this time."
        End Select
    Next X
    
    'Only 1 IF allowed in Regular editor
    If IfCnt > 1 Then
        SetError IFPos, 1
        Err.Raise gUserErr, , "Multiple IF's are not supported in this version."
    End If
    
    'Only 1 Compiled Action is allowed in Regular editor
    If CACnt > 1 Then
        SetError CompiledActionPos, 1
        Err.Raise gUserErr, , "Multiple Actions are not supported in this version."
    End If
    
    ValIfThen01
    ValIfThen03
    
'---------------------------------------------------------------------------
'Build coded text and inputs
    
    ' DAJ/TLB 09/16/2008: We have decided not to do this for now because of issues...
    ' If a parameter does not begin with the variable prefix, prepend it now...
    'For X = 1 To mLast
    '    If (PhraseType(X) = gPT_Parm) Or (PhraseType(X) = gPT_Bars) Then
    '        If Phrase(X) <> "Market1" Then
    '            If Left(Phrase(X), 1) <> VARIABLE_PREFIX Then
    '                gdSetStr aPhrase, X, VARIABLE_PREFIX & Phrase(X)
    '            End If
    '        End If
    '    End If
    'Next X
    
    'Build inputs used by Rule
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            BuildInputs X
        End If
    Next X

    For X = mInputs.Count To 1 Step -1
        If m.astrMacros.BinarySearch(mInputs.Item(X).ParmName & vbTab, , eGdSort_MatchUsingSearchStringLength) Then
            mInputs.Remove X
        End If
    Next X
    
    'Ensure Market1 is the first parm in inputs collection
    ReSortInputs
    
    'Determine whether late calculating functions are used in condition and/
    'or actions (for System Navigator)
    If Not mPortNavigator Then
        mLateCondition = False
        mLateAction = False
        CalcLateCondition IFPos
        CalcLateAction CompiledActionPos
    End If
    
    'Build engine codedtext...
    mText = ""
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            BuildCoded X
        End If
    Next X
        
    For lIndex2 = 0 To m.astrMacros.Size - 1
        strToken = Format(Len(Parse(m.astrMacros(lIndex2), vbTab, 1)), "000") & Parse(m.astrMacros(lIndex2), vbTab, 1)
        strNewCoded = "~14001( " & Parse(m.astrMacros(lIndex2), vbTab, 3) & " ~15001)"
        
        mText = Replace(mText, "~05" & strToken, strNewCoded)
        mText = Replace(mText, "~06" & strToken, strNewCoded)
        mText = Replace(mText, "~07" & strToken, strNewCoded)
        mText = Replace(mText, "~25" & strToken, strNewCoded)
        mText = Replace(mText, "~27" & strToken, strNewCoded)
        mText = Replace(mText, "~28" & strToken, strNewCoded)
    Next lIndex2
        
    mCodedText = mText

    'reset the comment engine
    ResetCommentNode
    'Build editor codedtext...
    mText = ""
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            BuildFormatted X, False, True
        End If
    Next X
    mEditText = strFormatted & " " & mText
    
    'reset the comment engine
    ResetCommentNode
    'Create preview codedtext...
    mText = ""
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            BuildFormatted X, True, True
        End If
    Next X
    mPreview = strPreview & " " & mText
        
    Exit Sub

ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then

        'Build the formatted text before raising the error.
        svErr = Err.Number
        svDesc = Err.Description
        svSource = Err.Source
        'reset the comment engine
        ResetCommentNode
        mText = ""
        For X = 1 To mLast
            If PhraseType(X) = gPT_Enter Then
                BuildFormatted X, False, True
            End If
        Next X
        mEditText = strFormatted & " " & mText
        Err.Raise svErr, svSource, svDesc
        
    Else
        Err.Raise Err.Number, mClass & " Sub:ValidateRule ", Err.Description
    End If

End Sub

Private Sub ReSortInputs()
On Error GoTo ErrSection:
    Dim NewInputs       As Object
    Dim NextInput       As Long
    Dim X               As Long
    
    If Not mPortNavigator Then
    
        'Holds the new inputs collection with Market1 first
        Set NewInputs = CreateObject(mCommonDLL & "cInputs")
        
        'Loop through existing inputs.  Move Market1 to top of list.
        For X = 1 To mInputs.Count
            If mInputs.Item(X).ParmName = "Market1" Then
                NextInput = NextInput + 1
                mInputs.Item(X).ParmSeq = NextInput
                
                With mInputs.Item(X)
                    NewInputs.Add .RuleName, NextInput, _
                        .ParmName, .ParmDesc, .ParmID, _
                        .Value, 0, 0, 0, 0, 0, 0, .ParmTypeID, _
                        .DefaultValue, .Required, .FromValue, _
                        .ToValue, .ListID, .FillPre, .FillPost
                End With
            End If
        Next X
        
        For X = 1 To mInputs.Count
            If mInputs.Item(X).ParmName <> "Market1" Then
                NextInput = NextInput + 1
                mInputs.Item(X).ParmSeq = NextInput
                With mInputs.Item(X)
                    NewInputs.Add .RuleName, NextInput, _
                        .ParmName, .ParmDesc, .ParmID, _
                        .Value, 0, 0, 0, 0, 0, 0, .ParmTypeID, _
                        .DefaultValue, .Required, .FromValue, _
                        .ToValue, .ListID, .FillPre, .FillPost
                End With
            End If
        Next X
        
        'Point newly ordered inputs to mInputs
        Set mInputs = CreateObject(mCommonDLL & "cInputs")
        Set mInputs = NewInputs
    
    End If
    
    Exit Sub
ErrSection:
    Err.Raise Err.Number, mClass & "Sub:ReSortInputs", Err.Description
End Sub

'Driver to validate the rule of a TradeSense Function
Public Sub ValidateFunctionRule(strExpression As String)
On Error GoTo ErrSection:
    
    Dim svErr       As Long
    Dim svDesc      As String
    Dim svSource    As String
    Dim X           As Long
    
    Dim astrExpressions As cGdArray     ' Array of expressions
    Dim Expr As cExpression             ' Expression object for assignments
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim strVariable As String           ' Variable name for assignments
    Dim strAssignment As String         ' Expression for assignments
    Dim strFormatted As String          ' Formatted text for the whole expression
    Dim strPreview As String            ' Preview text for the whole expression
    Dim strCodedText As String          ' Coded text for assignments
    Dim strNewCoded As String           ' Coded text for assignments with parenthesis
    Dim strToken As String              ' Token for variable name
    Dim lPos As Long                    ' Position in an array
    Dim strMainExpr As String           ' Main expression
    Dim strEnglishText As String        ' English version of the coded text
    Dim bMainExprHasMacros As Boolean   ' Does the main expression have macros?
    Dim strTempExpr As String           ' Temporary expression
    Dim bNeedToAppendLater As Boolean   ' Need to append to formatted and preview strings later
    
'-------------------------------------------------------------------------
'Parse out and handle any macro variable declarations...
    
    ' TLB 9/20/2012: rename "LW Sentiment" functions
    strExpression = Replace(strExpression, "LW Sentiment", "TN Consensus")
    
    ' DAJ 03/20/2009: Need to make sure that there are spaces around any math
    ' operator that isn't in comments or quotes...
    strExpression = SpacesAroundOperators(strExpression)
    
    Set astrExpressions = New cGdArray
    astrExpressions.SplitFields Trim(strExpression), vbLf
    
    Set mInputs = CreateObject(mCommonDLL & "cInputs")
    
    ' Strip any blank lines at the end of the expression...
    For lIndex = astrExpressions.Size - 1 To 0 Step -1
        If Len(Trim(astrExpressions(lIndex))) = 0 Then
            astrExpressions.Remove lIndex
        Else
            Exit For
        End If
    Next lIndex
    
    ' First, verify that there is only one main expression and it occurs on the last line...
    For lIndex = 0 To astrExpressions.Size - 1
        If lIndex < astrExpressions.Size - 1 Then
            If Len(Trim(astrExpressions(lIndex))) = 0 Then
                ' TLB 3/23/2015: need to replace a line of only spaces with an empty string
                astrExpressions(lIndex) = ""
            Else
                If InStr(astrExpressions(lIndex), ":=") = 0 Then
                    If (Left(Trim(astrExpressions(lIndex)), 1) <> "{") Or (Right(Trim(astrExpressions(lIndex)), 1) <> "}") Then
                        Err.Raise gUserErr, , "Main expression must be on the last line"
                    End If
                End If
            End If
        Else
            If InStr(astrExpressions(lIndex), ":=") <> 0 Then
                Err.Raise gUserErr, , "There must be one non-assignment expression"
            End If
        End If
    Next lIndex
    
    For lIndex = 0 To astrExpressions.Size - 1
        If Len(astrExpressions(lIndex)) = 0 Then
            strFormatted = strFormatted & " ~80003<E>"
            strPreview = strPreview & " ~80003<E>"
        ElseIf InStr(astrExpressions(lIndex), ":=") <> 0 Then
            If Left(astrExpressions(lIndex), 1) <> VARIABLE_PREFIX Then
                astrExpressions(lIndex) = VARIABLE_PREFIX & astrExpressions(lIndex)
            End If
            
            strVariable = Parse(astrExpressions(lIndex), ":=", 1)
            strAssignment = Parse(astrExpressions(lIndex), ":=", 2)
            strEnglishText = strAssignment
            
            Set Expr = New cExpression
            Expr.PortfolioNavigator = False
            Expr.Functions = mFunctions
            Expr.Inputs = mInputs
            Expr.ValidateFunctionRule strAssignment
            
            AddInputs Expr.Inputs
            
            strFormatted = strFormatted & " ~" & Format(gPARM_NUMERIC, "00") & Format(Len(strVariable), "000") & strVariable
            strFormatted = strFormatted & " ~" & Format(gCOMPARE, "00") & "002:= " & Expr.EditText
            
            strPreview = strPreview & " ~" & Format(gPARM_NUMERIC, "00") & Format(Len(strVariable), "000") & strVariable
            strPreview = strPreview & " ~" & Format(gCOMPARE, "00") & "002:= " & Expr.Preview
            
            strCodedText = Expr.CodedText
                        
            For lIndex2 = m.astrMacros.Size - 1 To 0 Step -1
                strToken = Format(Len(Parse(m.astrMacros(lIndex2), vbTab, 1)), "000") & Parse(m.astrMacros(lIndex2), vbTab, 1)
                strNewCoded = "~14001( " & Parse(m.astrMacros(lIndex2), vbTab, 3) & " ~15001)"
                
                strCodedText = Replace(strCodedText, "~05" & strToken, strNewCoded)
                strCodedText = Replace(strCodedText, "~06" & strToken, strNewCoded)
                strCodedText = Replace(strCodedText, "~07" & strToken, strNewCoded)
                strCodedText = Replace(strCodedText, "~25" & strToken, strNewCoded)
                strCodedText = Replace(strCodedText, "~27" & strToken, strNewCoded)
                strCodedText = Replace(strCodedText, "~28" & strToken, strNewCoded)
                
                strEnglishText = Replace(strEnglishText, Parse(m.astrMacros(lIndex2), vbTab, 1), Parse(m.astrMacros(lIndex2), vbTab, 4))
            Next lIndex2
            
            If m.astrMacros.BinarySearch(strVariable & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = False Then
                m.astrMacros.Add strVariable & vbTab & strAssignment & vbTab & strCodedText & vbTab & strEnglishText, lPos
            End If
        ElseIf (Left(Trim(astrExpressions(lIndex)), 1) = "{") And (Right(Trim(astrExpressions(lIndex)), 1) = "}") Then
            strFormatted = strFormatted & " ~" & Format(gCOMMENT, "00") & Format(Len(Trim(astrExpressions(lIndex))), "000") & Trim(astrExpressions(lIndex)) & " ~80003<E>"
            strPreview = strPreview & " ~" & Format(gCOMMENT, "00") & Format(Len(Trim(astrExpressions(lIndex))), "000") & Trim(astrExpressions(lIndex)) & " ~80003<E>"
        ElseIf (lIndex = astrExpressions.Size - 1) Then
            If m.astrMacros.Size = 0 Then
                bNeedToAppendLater = True
                strEnglishText = astrExpressions(lIndex)
            Else
                bNeedToAppendLater = False
                
                Set Expr = New cExpression
                Expr.PortfolioNavigator = False
                Expr.Functions = mFunctions
                Expr.Inputs = mInputs
                Expr.ValidateFunctionRule astrExpressions(lIndex)
                
                If Len(strFormatted) > 0 Then
                    strFormatted = strFormatted & " " & Expr.EditText
                Else
                    strFormatted = Expr.EditText
                End If
                If Len(strPreview) > 0 Then
                    strPreview = strPreview & " " & Expr.Preview
                Else
                    strPreview = Expr.Preview
                End If
            
                strEnglishText = astrExpressions(lIndex)
                For lIndex2 = m.astrMacros.Size - 1 To 0 Step -1
                    strEnglishText = Replace(strEnglishText, Parse(m.astrMacros(lIndex2), vbTab, 1), "(" & Parse(m.astrMacros(lIndex2), vbTab, 4) & ")")
                Next lIndex2
            End If
        End If
    Next lIndex
    
'-------------------------------------------------------------------------
'Parse text into Phrase and Phrase type gdArrays and then into a Tree...
    
    'Trim off enter key at end (occurs during library reverifying)
    strMainExpr = Trim(astrExpressions(astrExpressions.Size - 1))
    'strMainExpr = strEnglishText
    If Right(strMainExpr, 3) = "<E>" Then
        strMainExpr = Left(strMainExpr, Len(strMainExpr) - 3)
    End If
    
    mEnterKey.ReFormat strMainExpr
    mStrh.InsertOneBlank strMainExpr
    VerifyBlanksAroundSeparators strMainExpr
    
    strMainExpr = RTrim(strMainExpr)
    
    Chk01 strMainExpr     'Make sure parenthesis match
    Chk02 strMainExpr     'Make sure double quotes (for text) match
    ChkComments strMainExpr 'Make sure brackets match
    
    'Parse strMainExpr and build Phrase and Phrase type gdArrays for use in mTrees
    BuildPhrases strMainExpr
    
    Chk03   'Differentiate between function parens and regular parens
    
    'For X = 1 To gdGetSize(mPhrases) - 1
    '    Debug.Print gdGetStr(mPhrases, X)
    'Next X
    
    SaveUserFormatting   'Mark <Enterkeys> between IF/Then constucts
    
'-------------------------------------------------------------------------
'Build Tree and phrase array's

    BuildTree
    
'-------------------------------------------------------------------------
'Post Tree building validations
    
    'Don't allow IF/THEN structures
    For X = 1 To mLast
        ValGeneric X
        Select Case PhraseType(X)
            
            Case gPT_If, gPT_Else, gPT_ElseIf, gPT_Then, gPT_EndIf
                Err.Raise gUserErr, , "IF/THEN/ELSE operators cannot be used."
            
            Case gPT_DoUntil, gPT_EndDo
                Err.Raise gUserErr, , "DoUntil/EndDo operators cannot be used."
            
            Case gPT_FCompiledAction
                Err.Raise gUserErr, , "Action Functions are not allowed in TradeSense Functions."
            
        End Select
    Next X
         
'--------------------------------------------------------------------------
'Build codedtext and inputs
    
    ' DAJ/TLB 09/16/2008: We have decided not to do this for now because of issues...
    ' If a parameter does not begin with the variable prefix, prepend it now...
    'For X = 1 To mLast
    '    If (PhraseType(X) = gPT_Parm) Or (PhraseType(X) = gPT_Bars) Then
    '        If Phrase(X) <> "Market1" Then
    '            If Left(Phrase(X), 1) <> VARIABLE_PREFIX Then
    '                gdSetStr aPhrase, X, VARIABLE_PREFIX & Phrase(X)
    '            End If
    '        End If
    '    End If
    'Next X
    
    'Build inputs used by Rule
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            BuildInputs X
        End If
    Next X
    
    For X = mInputs.Count To 1 Step -1
        If m.astrMacros.BinarySearch(mInputs.Item(X).ParmName & vbTab, , eGdSort_MatchUsingSearchStringLength) Then
            mInputs.Remove X
        End If
    Next X
    
    'Ensure Market1 is the first parm in inputs collection
    ReSortInputs
    
    'Determine whether late calculating functions are used
    '(for System Navigator)
    If Not mPortNavigator Then
        mLateCondition = False
        CalcLateFunction 1
    End If
    
    'Build engine codedtext...
    mText = ""
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            BuildCoded X
        End If
    Next X
    
    'Strip out Enter key at end of codedtext for system navigator
    If Not mPortNavigator Then
        mText = Trim(mText)
        If Right(mText, 9) = "~80003<E>" Then
            mText = Left(mText, Len(mText) - 9)
        End If
    End If
    
    bMainExprHasMacros = False
    For lIndex2 = 0 To m.astrMacros.Size - 1
        strToken = Format(Len(Parse(m.astrMacros(lIndex2), vbTab, 1)), "000") & Parse(m.astrMacros(lIndex2), vbTab, 1)
        strNewCoded = "~14001( " & Parse(m.astrMacros(lIndex2), vbTab, 3) & " ~15001)"
        
        strTempExpr = mText
        
        mText = Replace(mText, "~05" & strToken, strNewCoded)
        mText = Replace(mText, "~06" & strToken, strNewCoded)
        mText = Replace(mText, "~07" & strToken, strNewCoded)
        mText = Replace(mText, "~25" & strToken, strNewCoded)
        mText = Replace(mText, "~27" & strToken, strNewCoded)
        mText = Replace(mText, "~28" & strToken, strNewCoded)
        
        If mText <> strTempExpr Then
            bMainExprHasMacros = True
        End If
    Next lIndex2
        
    mCodedText = mText
    
'--------------------------------------------------------------------------
'Build formatted text
    
    'reset the comment engine
    ResetCommentNode
    'Build editor codedtext...
    mText = ""
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            BuildFormatted X, False, True
        End If
    Next X
    If Len(strFormatted) > 0 Then
        mEditText = strFormatted '& " " & mText
        If bNeedToAppendLater = True Then
            mEditText = mEditText & " " & mText
        End If
    Else
        mEditText = mText
    End If
    
'--------------------------------------------------------------------------
'Build preview text
    
    'reset the comment engine
    ResetCommentNode
    'Create preview codedtext...
    mText = ""
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            BuildFormatted X, True, True
        End If
    Next X
    If Len(strPreview) > 0 Then
        mPreview = strPreview '& " " & mText
        If bNeedToAppendLater = True Then
            mPreview = mPreview & " " & mText
        End If
    Else
        mPreview = mText
    End If
    
    ' If the main expression has a macro in it, but the return type for the function has
    ' been determined to be a constant, change the return type to the appropriate series...
    If bMainExprHasMacros = True Then
        Select Case FunctionReturnType
            Case gRetConstantNbr
                FunctionReturnType = gRetSeriesNbr
            Case gRetConstantBoolean
                FunctionReturnType = gRetSeriesBoolean
        End Select
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        
        'Build the formatted text before raising the error.
        svErr = Err.Number
        svDesc = Err.Description
        svSource = Err.Source
        If mLast > 0 Then
            'reset the comment engine
            ResetCommentNode
            BuildFormatted 1, False, True
            mEditText = strFormatted & " " & mText
        End If
        Err.Raise svErr, svSource, svDesc
        
    Else
        Err.Raise Err.Number, mClass & " Sub:ValidateFunctionRule ", Err.Description
    End If

End Sub

Private Sub BuildTree()
On Error GoTo ErrSection:

    'Create new Tree and retrieve array addresses for quick access
    Set mTrees = New cDbTrees
    With mTrees
        aPhraseType = .AddrPhraseType
        aReturnType = .AddrReturnType
        aPhrase = .AddrPhrase
        aFormatting = .AddrFormatting
        aNodeFrom = .AddrNodeFrom
        aNodeTo = .AddrNodeTo
        aParent = .AddrParent
        aFID = .AddrFID
        aClassName = .AddrClassName
        aInputID = .AddrInputID
        aEntered = .AddrEntered
        aReq = .AddrReq
        aError = .AddrError
        .Functions = mFunctions
        .AddEnglishText mPhrases, mPhraseTypes, mFormatting
        .BuildTree
    End With
    mLast = mTrees.Last
    
    ReportAnyBuildErrors    'Show any errors found during Tree building
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:BuildTree ", Err.Description
    End If
End Sub

'Make sure 1 blank surrounds all separator symbols.
'Except for negative numbers -100 and Functions followed by offsets)
Private Sub VerifyBlanksAroundSeparators(pText As String)
On Error GoTo ErrSection:
    Dim Symbol      As String
    Dim CurChar     As String
    Dim curPos      As Integer
    Dim CommentMode As Boolean
    
    If Len(pText) < 2 Then Exit Sub
    
    'Insert one blank before symbols (if not already there)
    curPos = 1
    Do Until curPos > Len(pText)
        curPos = mStrh.NextSepSymbolPos(pText, curPos, Len(pText), Symbol)
        If curPos = 0 Then Exit Do
        
        If Symbol = "'" Or Symbol = """" Then
            CommentMode = Not CommentMode ' True
        End If
        If CommentMode And Symbol = Chr(13) & Chr(10) Then
            CommentMode = False
        End If
        
        If Symbol <> "" Then
        
            'Check character before symbol.  If not blank and not a period
            'then insert a blank.
            If curPos > 1 Then
                CurChar = Mid(pText, curPos, 1)
                If (Mid(pText, curPos - 1, 1) <> " ") And (CurChar <> ".") And (CurChar <> Chr(34)) And (Not CommentMode) Then
                    pText = Left(pText, curPos - 1) & " " & _
                            Right(pText, Len(pText) - curPos + 1)
                    curPos = curPos + 1
                End If
            End If
            
            'Check character after symbol
            If Mid(pText, curPos + Len(Symbol), 1) <> " " And Not CommentMode Then
                'Don't insert blank after minus signs immediately followed by
                'a numeric value
                If Symbol = "-" And IsNumeric(Mid(pText, curPos + 1, 1)) Then
                Else
                    pText = Left(pText, curPos + Len(Symbol) - 1) & " " & _
                            Right(pText, Len(pText) - curPos - Len(Symbol) + 1)
                    curPos = curPos + 1
                End If
            End If
            curPos = curPos + Len(Symbol)
        Else
            curPos = curPos + 1
        End If
    Loop
    
    Exit Sub
ErrSection:
    If Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:VerifyBlanksAroundSeparators ", Err.Description
    End If
End Sub

'Parsed Text into Phrases gdArrays
Private Sub BuildPhrases(pText As String)
On Error GoTo ErrSection:
    
    Dim svPos               As Integer
    Dim curPos              As Integer
    Dim curPhrase           As String
    Dim curPhraseType       As Integer
    Dim NextSymbol          As String
    Dim NextSepBegPos       As Integer
    Dim X                   As Long
    Dim DoubleQuoteFnd      As Boolean
    Dim additionalPhraseCtr As Integer
    
    'Create gdArrays to store phrases and phrase types
    'NOTE:  These gdArray address are passed into the Tree and stored at the
    'top of the tree.
    gdSetSize mPhrases, 0, 0
    gdSetSize mPhraseTypes, 0, 0
    gdSetSize mFormatting, 0, 0
    mComments.Size = 0
    
    additionalPhraseCtr = 0
    curPos = 1
    Do Until curPos > Len(pText)
        
        'Skip blanks...
        curPos = mStrh.FindNonBlank(pText, curPos)
        
        'cur group of characters form a separator symbol...
        NextSepBegPos = _
            mStrh.NextSepSymbolPos(pText, curPos, curPos, NextSymbol)
        If NextSymbol <> "" Then
            curPhraseType = GetPhraseType(NextSymbol)
            Select Case curPhraseType
            
                'If the Separate symbol is a Double quote, then everying follow
                'ing it is text.  Extract it now...
                Case gPT_DoubleQuote
                    
                    'Find next double quote
                    svPos = curPos
                    curPos = InStr(curPos + 1, pText, NextSymbol)
                    
                    'Nothing found between double quotes
                    X = X + 1
                    If curPos - 1 = svPos Then
                        gdSetStr mPhrases, X, ""
                        gdSetNum mPhraseTypes, X, gPT_Text
                    Else
                        'Text found between double quotes
                        curPhrase = Trim(Mid(pText, svPos + 1, curPos - svPos - 1))
                        curPhrase = """" & curPhrase & """"
                        gdSetStr mPhrases, X, curPhrase
                        gdSetNum mPhraseTypes, X, gPT_Text
                    End If
                    curPos = curPos + 1
            
                'If the Separate symbol is a bracket comment, then everying following
                'it up to the matching bracket is text.  Extract it now...
                Case gPT_BracketComment
                    
                    'Find the matching bracket
                    Dim iBCnt As Integer, bMatchFound As Boolean
                    iBCnt = 1 'already found one bracket
                    svPos = curPos
                    bMatchFound = False
                    While Not bMatchFound And svPos <= Len(pText)
                        svPos = svPos + 1 'move to next char
                        'either we found the matching bracket
                        If Mid(pText, svPos, 1) = "}" Then
                            iBCnt = iBCnt - 1
                            bMatchFound = iBCnt < 1
                        'or found a nested bracket
                        ElseIf Mid(pText, svPos, 1) = "{" Then
                            iBCnt = iBCnt + 1
                        'or found the comment text
                        Else
                        End If
                    Wend
                    'if the next non-space char is a '{' we have another comment following
                    '   need to add it to this one for the engine as it sits to work
                    
                    'Text found between brackets
                    curPhrase = Trim(Mid(pText, curPos, svPos - (curPos - 1)))
                    Dim iCommentNode As Integer
                    'due to positioning, since much of the code looks at the 'users' inputed text
                    'and not our additional formatted text, we keep a count of what we have
                    'added and set the comment node to where the user actually placed it
                    ' - note: since we are dealing with position in text string, remove 1
                    iCommentNode = X - (additionalPhraseCtr - 1)
                    curPos = svPos + 1
                
                    'tjr 2/03 - set for comment
                    mComments(iCommentNode) = mComments(iCommentNode) & curPhrase
                    
                'Single Quote (comment).  Everything right up to enterkey
                'is part of the comment
                Case gPT_Comment
                    svPos = curPos
                    curPos = InStr(curPos + 1, pText, Chr(13) & Chr(10))
                    If curPos <= 0 Then
                        curPos = Len(pText)
                    End If
                    
                    'No comments found
                    X = X + 1
                    If curPos - 1 = svPos Then
                        gdSetStr mPhrases, X, ""
                        gdSetNum mPhraseTypes, X, gPT_Comment
                    Else
                        'Text found between double quotes
                        curPhrase = Trim(Mid(pText, svPos + 1, curPos - svPos - 1))
                        curPhrase = "'" & curPhrase
                        gdSetStr mPhrases, X, curPhrase
                        gdSetNum mPhraseTypes, X, gPT_Comment
                    End If
                    curPos = curPos - 1 'Set back one so enterkey is picked up on next pass
                    
                'Regular separate handling.  Add to phrases arrays...
                'Notice: IF/THEN and EnterKeys are added to the phrases array
                'when building the tree
                Case Else
                    
                    curPos = NextSepBegPos + Len(NextSymbol)
                    
                    Select Case curPhraseType
                        Case gPT_Enter
                            NextSymbol = "<E>"
                            'If Previous phrase was also enter, then this is
                            'a formatting enter (blank line)
                            If gdGetStr(mPhrases, X) = "<E>" Then
                                NextSymbol = "<EF>"
                                additionalPhraseCtr = additionalPhraseCtr + 1
                            End If
                           
                        Case gPT_Tab
                            NextSymbol = "<Tab>"
                            additionalPhraseCtr = additionalPhraseCtr + 1
                    End Select
                    X = X + 1
                    gdSetStr mPhrases, X, NextSymbol
                    gdSetNum mPhraseTypes, X, curPhraseType
            End Select
        Else
            'current group of characters form a numeric or text value...
            NextSepBegPos = _
                mStrh.NextSepSymbolPos(pText, curPos, Len(pText), NextSymbol)
            If NextSepBegPos = Len(pText) Then
                If NextSymbol = "" Then
                    curPhrase = _
                        Trim(Mid(pText, curPos, NextSepBegPos - curPos + 1))
                Else
                    curPhrase = _
                        Trim(Mid(pText, curPos, NextSepBegPos - (curPos + 1)))
                End If
            Else
                curPhrase = Trim(Mid(pText, curPos, NextSepBegPos - curPos))
            End If
            
            curPhraseType = GetPhraseType(curPhrase)
            X = X + 1
            
            'MT Dec/2001:
            'Inputs must be formatted the same way so they are sorted
            'and grouped correctly in cdbtrees.  This is necessary when
            'assigning gdArrays to inputs.
'            'tjr - 8/03: allow ANY case user puts in
'            If curPhraseType = gPT_Parm Then
'                curPhrase = FileNameDisplay(LCase(curPhrase))
'            End If
            
            gdSetStr mPhrases, X, curPhrase
            gdSetNum mPhraseTypes, X, curPhraseType
            
            curPos = NextSepBegPos
            If curPos = Len(pText) And NextSymbol = "" Then Exit Do
        End If
        
        'Default formatting information to ""
        gdSetStr mFormatting, X, ""
    Loop
    
    gdFreeExtra mPhrases
    gdFreeExtra mPhraseTypes
    gdFreeExtra mFormatting
    
ErrExit:
    Exit Sub

ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:BuildPhrases ", Err.Description
    End If

End Sub

'===================== BUILT-IN CODEDTEXT ================================
Public Function BuiltinCodedText(pCodedName As String, pParmTypeID As Long, _
        pInputs As cInputs) As String
On Error GoTo ErrSection:
    Dim X               As Integer
    Dim Token           As String
    Dim wrkCodedText    As String
    
    BuiltinCodedText = ""
    
    'Coded name of function goes first...
    Token = TokenizeBuiltin(pCodedName, ConvertParmTypeID_ToToken(pParmTypeID), Len(pCodedName))
    wrkCodedText = wrkCodedText & Token & " "
    
    If pInputs.Count <> 0 Then
        
        'Insert a parenthesis for the beginning of the inputs list
        Token = TokenizeBuiltin("(", gFLEFTPAREN, 1)
        wrkCodedText = wrkCodedText & Token & " "
        
        For X = 1 To pInputs.Count
            'Insert input name...
            With pInputs.Item(X)
                Token = TokenizeBuiltin(.ParmName, _
                    ConvertParmTypeID_ToToken(.ParmTypeID), _
                    Len(.ParmName))
                wrkCodedText = wrkCodedText & Token & " "
            End With
            
            'Insert comma...
            If X + 1 <= pInputs.Count Then
                Token = TokenizeBuiltin(",", gCOMMA, 1)
                wrkCodedText = wrkCodedText & Token & " "
            End If
        Next X
        
        'Insert a right parenthesis to end the inputs list
        Token = TokenizeBuiltin(")", gFRIGHTPAREN, 1)
        wrkCodedText = wrkCodedText & Token
    End If

    BuiltinCodedText = wrkCodedText
    
    Exit Function
ErrSection:
    Err.Raise Err.Number, "Class:cExpression  Sub:BuiltinCodedText" & Err.Number & " " & Err.Description
End Function
Private Function ConvertParmTypeID_ToToken(pParmTypeID As Long) As Byte
    Select Case pParmTypeID
        Case gRetConstantNbr: ConvertParmTypeID_ToToken = gFUNC_NUMERIC_CONSTANT
        Case gRetConstantText: ConvertParmTypeID_ToToken = gFUNC_TEXT_CONSTANT 'gTextT
        Case gRetSeriesBoolean: ConvertParmTypeID_ToToken = gFUNC_BOOLEAN
        Case gRetSeriesNbr: ConvertParmTypeID_ToToken = gFUNC_NUMERIC
        Case gRetBars: ConvertParmTypeID_ToToken = gPARM_BARS
        Case gRetConstantBoolean: ConvertParmTypeID_ToToken = gFUNC_BOOLEAN_CONSTANT
        Case gRetTrades: ConvertParmTypeID_ToToken = gPARM_TRADES
        Case gRetPortfolio: ConvertParmTypeID_ToToken = gPT_Portfolio
        Case gRetSystems: ConvertParmTypeID_ToToken = gPT_Systems
    End Select
End Function
Private Function TokenizeBuiltin(pText As String, pToken As Byte, _
            pLen As Integer) As String
    TokenizeBuiltin = "~" & Format(pToken, "00") & Format(pLen, "000") & pText
End Function
'===================== BUILT-IN CODEDTEXT ================================

'Differentiate between funtion parenthesis and regular parenthesis
Private Sub Chk03()
On Error GoTo ErrSection:
    Dim svPos   As Long
    Dim Cnt     As Long
    Dim PT      As Byte
    Dim PTPrev  As Byte
    Dim Parens  As Integer
    Dim X       As Long
    
    Cnt = gdGetSize(mPhraseTypes) - 1
    
    For X = 1 To Cnt
        If X > 1 Then
            PT = gdGetNum(mPhraseTypes, X)
            If PT = gPT_LeftPar Then
                PTPrev = gdGetNum(mPhraseTypes, X - 1)
                If PTPrev = gPT_FCompiled Or _
                   PTPrev = gPT_FCompiledAction Or _
                   PTPrev = gPT_FTradeSense Or _
                   PTPrev = gPT_FInternal Then
                    svPos = X
                    gdSetNum mPhraseTypes, X, gPT_FLParen
                
                    'Find and set right paren to function
                    Parens = 1
                    Do Until Parens = 0
                        X = X + 1
                        If X > Cnt Then Exit Do
                        Select Case gdGetNum(mPhraseTypes, X)
                            Case gPT_LeftPar: Parens = Parens + 1
                            Case gPT_RightPar: Parens = Parens - 1
                        End Select
                    Loop
                    
                    'Set right paren to function right paren...
                    gdSetNum mPhraseTypes, X, gPT_FRParen
                    X = svPos
                End If
            End If
        End If
    Next X
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:Chk03 ", Err.Description
    End If
End Sub

Private Sub ValIfThen01()
On Error GoTo ErrSection:
    Dim X       As Long
    
    For X = 1 To mLast
        Select Case PhraseType(X)
            Case gPT_If, gPT_ElseIf
                If PhraseType(NodeTo(X)) <> gPT_Then Then
                    SetError X, 1
                    Err.Raise gUserErr, , "IF (or ELSEIF) does not " & _
                        "have a matching THEN"
                End If
                If ReturnType(NodeTo(X)) <> gRetSeriesBoolean And _
                   ReturnType(NodeTo(X)) <> gRetConstantBoolean Then
                    SetError X, 1
                    Err.Raise gUserErr, , _
                        "IF Conditions must result in a true/false value.  " & _
                        "Please review the expression.  Make sure you have comparison operator(s) and/or function(s) that return true/false."
                End If
        
            Case gPT_DoUntil
                If ReturnType(NodeTo(X)) <> gRetSeriesBoolean And _
                   ReturnType(NodeTo(X)) <> gRetConstantBoolean Then
                    SetError X, 1
                    Err.Raise gUserErr, , _
                        "DoUntil Condition must result in a true/false value.  " & _
                        "Please review the expression.  Make sure you have comparison operator(s) and/or function(s) that return true/false."
                End If
        
            Case gPT_Then
                If PhraseType(Parent(X)) <> gPT_If And _
                   PhraseType(Parent(X)) <> gPT_ElseIf Then
                    SetError X, 1
                    Err.Raise gUserErr, , "This THEN operator does not " & _
                        "have a matching IF or ELSEIF"
                End If
        
            Case gPT_Enter
                Select Case PhraseType(NodeFrom(X))
                    Case gPT_Comment, gPT_If, gPT_ElseIf, gPT_Else, _
                         gPT_EndIf, gPT_FCompiledAction, gPT_DoUntil, gPT_EndDo
                    Case Else
                        SetError NodeFrom(X), 1
                        Err.Raise gUserErr, , "This phrase cannot begin a line"
                End Select
                
        End Select
    Next X
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:ValIfThen01 ", Err.Description
    End If
End Sub

'Make sure IF/THEN/ELSE constructs are used correctly in rule
Private Sub ValIfThen02(pNode As Long)
On Error GoTo ErrSection:
    Dim X           As Long
    Dim Fnd         As Boolean
    Dim PT          As Byte
    
    PT = PhraseType(pNode)
    
    Select Case PT
    
        Case gPT_If, gPT_ElseIf, gPT_Else, gPT_EndIf, gPT_Then
        
            If PT = gPT_If Then mIfCnt = mIfCnt + 1
            If PT = gPT_ElseIf Then mElseIfCnt = mElseIfCnt + 1
            If PT = gPT_Then Then mThenCnt = mThenCnt + 1
            If PT = gPT_EndIf Then mEndIfCnt = mEndIfCnt + 1
            
            If mPrevPT = 0 Then
                mPrevPT = PT
            Else
                Select Case PT
                    Case gPT_If
                        If mPrevPT <> gPT_EndIf And mPrevPT <> gPT_Then And _
                           mPrevPT <> gPT_Else Then
                            SetError pNode, 1
                            Err.Raise gUserErr, , "IF is not being used correctly.  IF can follow THEN and ELSE or group a new condition."
                        End If
                    Case gPT_ElseIf
                        If mPrevPT <> gPT_Then And mPrevPT <> gPT_EndIf And mPrevPT <> gPT_EndDo Then
                            SetError pNode, 1
                            Err.Raise gUserErr, , "ELSEIF is not being used correctly.  THEN must match up with a corresponding THEN."
                        End If
                    Case gPT_Else
                        If mPrevPT <> gPT_Then And mPrevPT <> gPT_EndIf And mPrevPT <> gPT_EndDo Then
                            SetError pNode, 1
                            Err.Raise gUserErr, , "ELSE is not being used correctly.  ELSE must match up with a corresponding THEN."
                        End If
                    Case gPT_Then
                        If mPrevPT <> gPT_If And mPrevPT <> gPT_ElseIf Then
                            SetError pNode, 1
                            Err.Raise gUserErr, , "THEN is not being used correctly.  THEN must match up with a corresponding IF or ELSEIF."
                        End If
                    Case gPT_EndIf
                        If mPrevPT <> gPT_Then And mPrevPT <> gPT_Else And mPrevPT <> gPT_EndIf And mPrevPT <> gPT_EndDo Then
                            SetError pNode, 1
                            Err.Raise gUserErr, , "ENDIF is not being used correctly.  ENDIF must match up with a corresponding THEN or ELSE."
                        End If
                End Select
                mPrevPT = PT
            End If
        
    End Select
    
    If NodeFrom(pNode) > 0 Then
        ValIfThen02 NodeFrom(pNode)
    End If
    If NodeTo(pNode) > 0 Then
        ValIfThen02 NodeTo(pNode)
    End If
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:ValIfThen02 ", Err.Description
    End If
End Sub

Private Sub ValIfThen03()
On Error GoTo ErrSection:
    Dim X       As Long
    
    'Process each rule statement (enter key lines)
    For X = 1 To mLast
        If PhraseType(X) = gPT_Enter Then
            ValIfThen02 X
        Else
            Exit For
        End If
    Next X
    
    'Report If/Then/Else usage errors from ValIfThen02
    If mIfCnt + mElseIfCnt > mThenCnt Then
        Err.Raise gUserErr, , "An extra IF was found without a matching THEN"
    End If
    If mIfCnt + mElseIfCnt < mThenCnt Then
        Err.Raise gUserErr, , "An extra THEN was found without a matching IF"
    End If
    If mIfCnt > mEndIfCnt Then
        Err.Raise gUserErr, , "An extra IF was found without a matching ENDIF"
    End If
    If mIfCnt < mEndIfCnt Then
        Err.Raise gUserErr, , "An extra ENDIF was found without a matching IF"
    End If
    
    'Verify that IFs and ELSEIF's all return true/false
    'Standard Condition/Action validations
    For X = 1 To mLast
        Select Case PhraseType(X)
            Case gPT_If, gPT_ElseIf
                
                'Standard validation for IF
                ValCondition NodeTo(X)
                
            Case gPT_FCompiledAction
            
                'Offset values not allowed with action functions
                If PhraseType(Parent(X)) = gPT_Offset Then
                    SetError NodeTo(Parent(X)), 1
                    Err.Raise gUserErr, , _
                        "Offsets cannot be used with Action functions"
                End If
                
                'Standard validation for ACTION
                ValAction NodeTo(X)
                
        End Select
    Next X
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:ValIfThen03 ", Err.Description
    End If
End Sub

Private Sub ValDoUntil01()
On Error GoTo ErrSection:
    Dim X       As Long
    Dim R       As Long
    
    For X = 1 To mLast
        Select Case PhraseType(X)
            Case gPT_DoUntil: R = R + 1
            Case gPT_EndDo: R = R - 1
        End Select
    Next X
    
    If R <> 0 Then
        Err.Raise gUserErr, , "One or more DoUntil/EndDo blocks were found."
    End If
    
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:ValDoUntil01 ", Err.Description
    End If
End Sub

'Driver to Inputs collection (for parms entered in rule)
Private Sub BuildInputs(pNode As Long)
On Error GoTo ErrSection:
    Dim X           As Long
    Dim curPhrase   As String
    Dim Skip        As Boolean
    Dim Y           As Integer
    Dim xInput      As Object
    Dim NewID       As Long
    
    curPhrase = Phrase(pNode)
    
    'If phrase was defaulted then ignore
    If Entered(pNode) = 0 And PhraseType(pNode) <> gPT_Bars Then GoTo ErrExit
    
    If PhraseType(pNode) = gPT_Parm Or PhraseType(pNode) = gPT_Bars Then
    
        'Skip duplicate inputs
        Skip = False
        For Y = 1 To mInputs.Count
            
            'Dup parm found.  Make sure it's return type is the same for
            'both parms
            If curPhrase = mInputs.Item(Y).ParmName Then
                If Not mInputs.Compare(mInputs.Item(Y).ParmTypeID, ReturnType(pNode), NewID) And _
                   PhraseType(pNode) <> gPT_Bars Then
                    SetError pNode, 1
                    Err.Raise gUserErr, , "This input is not being used consistently in the rule.  For example, an input can't represent a number in one part of the rule, and as a text value in a different part of the rule."
                End If
                Skip = True
            End If
        Next Y
     
        'Skip "Market1" bar phrases if OF was used instead.  Search back up
        'the tree looking for a parent of phrase type OF.  If found then
        'an OF symbol was specified.  Don't add market1 to inputs.
        If Not Skip And PhraseType(pNode) = gPT_Bars And Phrase(pNode) = "Market1" Then
            Y = Parent(pNode)
            Do Until Y <= 0
                If PhraseType(Y) = gPT_Of Then
                    Skip = True
                    Exit Do
                End If
                Y = Parent(Y)
            Loop
        End If
        
        If Not Skip Then
        
            Set xInput = CreateObject(mCommonDLL & "cInput")
            
            '===========PORTFOLIO NAVIGATOR VERSION ==============
            '- Comment out code block below (SYSTEM NAVIGATOR VERSION)
            '- Uncomment code block below
            '- Set references to "PVCommon","PVVectors"
            '- Compile with Binary Comptability
            If mPortNavigator Then
            
                'If the FID has a value then this is an input.  The inputID
                'of the function input is stored here.
                mNextInput = mNextInput + 1
                'Set xInput = CreateObject(mCommonDLL & "cInput")
                If InputID(pNode) > 0 Then
                    With mFunctions.Item(CStr(FID(Parent(pNode)))).Inputs.Item(CStr(InputID(pNode)))
                        xInput.DefaultValue = .DefaultValue
                        xInput.ParmDesc = .ParmDesc
                        xInput.FillPre = .FillPre
                        xInput.ParmID = .ParmID
                        xInput.ListID = .ListID
                        xInput.ListName = .ListName
                        xInput.Required = .Required
                        xInput.ParmTypeName = .ParmTypeName
                        xInput.ParmTypeID = .ParmTypeID
                        xInput.ParmTypeUsage = .ParmTypeUsage
                        xInput.Value = .Value
                    End With
                Else
                    'This is a parm outside a function
                    With xInput
                        .ParmTypeID = ReturnType(pNode)
                        .Required = True
                        .Value = 0
                    End With
                End If
                xInput.ParmName = curPhrase
                xInput.ParmSeq = mNextInput
                mInputs.Add xInput, CStr(mNextInput)
                '===========PORTFOLIO NAVIGATOR VERSION ==============
            Else
                '===========SYSTEM NAVIGATOR VERSION ==============
                '- Comment out code block above (PORTFOLIO NAVIGATOR VERSION)
                '- Uncomment code block below
                '- Set references to "NavCommon"
                '- Compile with Binary Comptability
                
                'The "Add" method requires each field to passed in (for Binary
                'Compatibility).  Borrow out dated fields to pass in values
                'required for TradeSense
                '  - FromVal Maps to .ParmTypeUsage
                '  - ToVal   maps to .ParmTypeName
                '  - FillPos maps to .ListName
                mNextInput = mNextInput + 1
                If InputID(pNode) > 0 Then
                    With mFunctions.Item(CStr(FID(Parent(pNode)))).Inputs.Item(CStr(InputID(pNode)))
                        mInputs.Add "", mNextInput, _
                            curPhrase, .ParmDesc, mNextInput, _
                            .Value, 0, 0, 0, 0, 0, 0, .ParmTypeID, _
                            .DefaultValue, .Required, .ParmTypeUsage, _
                            .ParmTypeName, .ListID, "", .ListName
                    End With
                Else
                    mInputs.Add "", mNextInput, _
                        curPhrase, "", mNextInput, 0, 0, 0, 0, 0, 0, 0, ReturnType(pNode), _
                        "", True, 0, "", 0, "", ""
                End If
            End If
            '==========SYSTEM NAVIGATOR VERSION ==============

        End If
    End If
    
    Select Case PhraseType(pNode)
        Case gPT_FCompiled, gPT_FCompiledAction, _
             gPT_FTradeSense, gPT_FInternal
            
            If NodeFrom(pNode) > 0 Then
                For X = NodeFrom(pNode) To NodeTo(pNode)
                    BuildInputs X
                Next X
            End If
            
        Case Else
            If NodeFrom(pNode) > 0 Then
                BuildInputs NodeFrom(pNode)
            End If
            If NodeTo(pNode) > 0 Then
                BuildInputs NodeTo(pNode)
            End If
            
    End Select
    
ErrExit:
    Set xInput = Nothing
    Exit Sub
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:BuildInputs ", Err.Description
    End If
End Sub

Private Function Tokenize(pText As String, pToken As Byte, pLen As Integer) As String
On Error GoTo ErrSection:

    If mPortNavigator Then
        Tokenize = "~" & Format(pToken, "00") & Format(pLen, "000") & pText
    Else
        Tokenize = "~" & Format(CToken(pToken, pText), "00") & Format(pLen, "000") & pText
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:Tokenize ", Err.Description
    End If

End Function

'Translate Tokens from New TradeSense to Old System Navigator format for now
Private Function CToken(pTradeSenseToken As Byte, pFunctionName As String) As String
On Error GoTo ErrSection:
    Dim ID      As Long
    Dim Fnd     As Boolean
    
    Select Case pTradeSenseToken
        Case gPT_Nbr: CToken = gNUMERIC
        Case gPT_Add: CToken = gADD
        Case gPT_Sub: CToken = gADD
        Case gPT_Mult: CToken = gMULTI
        Case gPT_Div: CToken = gMULTI
        Case gPT_LeftPar: CToken = gLEFTPAREN
        Case gPT_RightPar: CToken = gRIGHTPAREN
        Case gPT_GT: CToken = gCOMPARE
        Case gPT_GE: CToken = gCOMPARE
        Case gPT_LT: CToken = gCOMPARE
        Case gPT_LE: CToken = gCOMPARE
        Case gPT_NE: CToken = gCOMPARE
        Case gPT_EQ: CToken = gCOMPARE
        Case gPT_And: CToken = gAND
        Case gPT_Or: CToken = gOR
        Case gPT_Not: CToken = gNOT
        Case gPT_Parm: CToken = gPARM_NUMERIC       '<<????
        Case gPT_Text: CToken = gTextT
        Case gPT_FLParen: CToken = gFLEFTPAREN
        Case gPT_FRParen: CToken = gFRIGHTPAREN
        Case gPT_Offset: CToken = gOFFSET
        Case gPT_Comma: CToken = gCOMMA
        Case gPT_BracketComment: CToken = gCOMMENT
        
        Case gPT_FInternal, gPT_FCompiled, gPT_FCompiledAction, gPT_FTradeSense
            
            'pFunctionName will have the codedname if engine codedtext is
            'being generated, otherwise the function name will be specified
            'for Editor codedtext.
            Fnd = False
            If mEngineCodedtext Then
                If CodedNameFound(pFunctionName, ID) Then
                    Fnd = True
                End If
            Else
                If mFunctions.NameFound(pFunctionName, ID) Then
                    Fnd = True
                End If
            End If
            If Fnd Then
                Select Case mFunctions.Item(CStr(ID)).ReturnTypeID
                    Case gRetSeriesNbr: CToken = gFUNC_NUMERIC
                    Case gRetSeriesBoolean: CToken = gFUNC_BOOLEAN
                    Case gRetConstantBoolean: CToken = gFUNC_BOOLEAN_CONSTANT
                    Case gRetConstantNbr: CToken = gFUNC_NUMERIC_CONSTANT
                    Case gRetConstantText: CToken = gFUNC_TEXT_CONSTANT
                    Case Else
                        CToken = gFUNC_NUMERIC
                End Select
            Else
                Err.Raise gUserErr, , "Token Translation error occurred for Function: " & UCase(pFunctionName) & ".  Unexpected results may occur."
            End If
            
        Case gPT_If: CToken = gIF
        Case gPT_DoUntil: CToken = gPT_DoUntil
        Case gPT_EndDo: CToken = gPT_EndDo
        Case gPT_DoubleQuote: CToken = gPT_DoubleQuote
        Case gPT_Error: CToken = gError
        Case gPT_OuterParens: CToken = gPT_OuterParens
        Case gPT_Then: CToken = gPT_Then
        Case gPT_Else: CToken = gPT_Else
        Case gPT_EndIf: CToken = gPT_EndIf
        Case gPT_ElseIf: CToken = gPT_ElseIf
        Case gPT_Tab: CToken = gPT_Tab
        Case gPT_Comment: CToken = gPT_Comment
        Case gPT_Enter, gPT_EnterFormatting: CToken = gPARAGRAPH
        Case gPT_Of: CToken = gOF
        Case gPT_Trades: CToken = gPARM_TRADES
        Case gPT_Bars: CToken = gPARM_BARS
        Case gPT_Portfolio: CToken = gPT_Portfolio
        Case gPT_Systems: CToken = gPT_Systems
    End Select
    
    Exit Function
ErrSection:
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, mClass & " Sub:CToken ", Err.Description
    End If
End Function

Private Function CodedNameFound(pName As String, pID As Long) As Boolean
On Error GoTo ErrSection:
    Dim X       As Integer
    Dim Cnt     As Integer
    
    CodedNameFound = False
    Cnt = mFunctions.Count
    For X = 1 To Cnt
        With mFunctions.Item(X)
            If .ClassName = pName Then
                pID = .FunctionID
                CodedNameFound = True
                Exit Function
            End If
        End With
    Next X
    
    Exit Function
ErrSection:
    Err.Raise Err.Number, mClass & " Sub:CodedNameFound", Err.Number & " " & Err.Number & " " & Err.Description
End Function

'Search for <Enter> and <Tab> formatting.  Delete these from the phrases
'arrays and insert an entry into the Formatting array attached to the next
'phrase in sequence
Private Sub SaveUserFormatting()
On Error GoTo ErrSection:
    Dim X               As Integer
    Dim Y               As Integer
    Dim UserFormat      As Boolean
    Dim Cnt             As Integer
    
    For X = 1 To gdGetSize(mPhrases) - 1
        
        If gdGetNum(mPhraseTypes, X) = gPT_Enter Then
            
            'Now, search for to see what the next line begins with.  It must be
            'with: Comment, If, Else, ElseIf, Endif, Action.  Otherwise, it
            'is assumed to be a formatting Enter key.
            UserFormat = True
            For Y = X + 1 To gdGetSize(mPhrases) - 1
                Select Case gdGetNum(mPhraseTypes, Y)
                    Case gPT_If, gPT_ElseIf, gPT_Else, _
                         gPT_EndIf, gPT_Comment, gPT_FCompiledAction, _
                         gPT_DoUntil, gPT_EndDo
                        UserFormat = False
                        Exit For
                    Case gPT_Tab
                    Case Else
                        Exit For
                End Select
            Next Y
            
            If X < gdGetSize(mPhrases) - 1 Then
                'Tag the phrase AFTER <enter> with an "E" to signify that
                'Enter should be inserted before the phrase during
                'formatting later.  Then, delete the enter key phrase.
                If UserFormat Then
                    gdSetStr mFormatting, X + 1, _
                            Trim(gdGetStr(mFormatting, X) & "E")
                    gdDeleteItems mPhrases, X, 1
                    gdDeleteItems mPhraseTypes, X, 1
                    gdDeleteItems mFormatting, X, 1
                    X = X - 1
                End If
            End If
            
        End If
    Next X
                            
    'Associate formatting tabs with the previous phrase found.
    'Assign 0 + n per tab
    For X = 1 To gdGetSize(mPhrases) - 2
        If gdGetNum(mPhraseTypes, X) = gPT_Tab Then
            gdSetStr mFormatting, X + 1, Trim(gdGetStr(mFormatting, X) & "T")
            gdDeleteItems mPhrases, X, 1
            gdDeleteItems mPhraseTypes, X, 1
            gdDeleteItems mFormatting, X, 1
            X = X - 1
        End If
    Next X
    
    'For X = 1 To gdGetSize(mPhrases) - 1
    '    Debug.Print gdGetStr(mPhrases, X) & ", " & gdGetStr(mFormatting, X)
    'Next X
    
ErrExit:
    Exit Sub

ErrSection:
    Err.Raise Err.Number, mClass & " Sub:SaveUserFormatting", Err.Number & " " & Err.Number & " " & Err.Description

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddInputs
'' Description: Add inputs to the main collection from the given collection
'' Inputs:      Inputs Collection
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddInputs(ByVal Inputs As Object)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim bFound As Boolean               ' Was the input found in the collection?
    
    For lIndex = 1 To Inputs.Count
        bFound = False
        
        For lIndex2 = 1 To mInputs.Count
            If Inputs.Item(lIndex).ParmName = mInputs.Item(lIndex2).ParmName Then
                bFound = True
                Exit For
            End If
        Next lIndex2

        If bFound = False Then
            mNextInput = mNextInput + 1
            
            With Inputs.Item(lIndex)
                mInputs.Add .RuleName, mNextInput, .ParmName, .ParmDesc, mNextInput, .Value, .IfOptimize, _
                            .OptFromValue, .OptToValue, .OptStepValue, .OptListID, .RuleID, .ParmTypeID, _
                            .DefaultValue, .Required, .FromValue, .ToValue, .ListID, .FillPre, .FillPost
            End With
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExpression.AddInputs", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnglishFromCoded
'' Description: Convert a coded text string back into an English string
'' Inputs:      Coded Text
'' Returns:     English Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnglishFromCoded(ByVal strCodedText As String) As String
On Error GoTo ErrSection:

    Dim astrTokens As cGdArray          ' Array of tokens
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrTokens = New cGdArray
    astrTokens.SplitFields strCodedText, "~"
    
    For lIndex = 0 To astrTokens.Size - 1
        astrTokens(lIndex) = Mid(astrTokens(lIndex), 6)
    Next lIndex
    
    EnglishFromCoded = astrTokens.JoinFields(" ")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cExpression.EnglishFromCoded", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SpacesAroundOperators
'' Description: Put spaces around any operators found outside quotes and comments
'' Inputs:      Expression
'' Returns:     Fixed Expression
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SpacesAroundOperators(ByVal strExpression As String) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bInComment As Boolean           ' Are we in a comment?
    Dim bInQuotes As Boolean            ' Are we in quotes?
    Dim strCharacter As String          ' Current character in the string
    Dim strPrevChar As String           ' Previous character in the string
    Dim strReturn As String             ' Return value for the function
    
    bInComment = False
    bInQuotes = False
    strPrevChar = " "
    
    For lIndex = 1 To Len(strExpression)
        If lIndex > 1 Then
            strPrevChar = strCharacter
        End If
        strCharacter = Mid(strExpression, lIndex, 1)
        
        Select Case strCharacter
            Case "{"
                strReturn = strReturn & strCharacter
                bInComment = True
            
            Case "}"
                strReturn = strReturn & strCharacter
                bInComment = False
            
            Case Chr(34)
                strReturn = strReturn & strCharacter
                bInQuotes = Not bInQuotes
            
            Case "+", "-", "*", "/"
                If bInComment Or bInQuotes Then
                    strReturn = strReturn & strCharacter
                ElseIf (IsAlpha(strPrevChar) = True) Or (IsDigit(strPrevChar) = True) Then
                    strReturn = strReturn & " " & strCharacter & " "
                Else
                    strReturn = strReturn & strCharacter
                End If
                
            Case Else
                strReturn = strReturn & strCharacter
        
        End Select
    Next lIndex
    
    SpacesAroundOperators = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cExpression.SpacesAroundOperators", eGDRaiseError_Raise, g.strAppPath
    
End Function

Private Sub Class_Initialize()
    
    Set mStrh = New cStrh
    Set mEnterKey = New cEnterKey
    Set mReturnTypes = New cParmTypes
    mReturnTypes.Load
    
    mPhrases = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
    mPhraseTypes = gdCreateArray(eGDARRAY_TinyInts, 0, USE_DEFAULT_NULL)
    mFormatting = gdCreateArray(eGDARRAY_Strings, 0, USE_DEFAULT_NULL)
    
    Set mComments = New cGdArray
    mComments.Create eGDARRAY_Strings
    
    Set m.astrMacros = New cGdArray
    m.astrMacros.Create eGDARRAY_Strings
    
End Sub

Private Sub Class_Terminate()
    
    gdDestroyArray mPhrases
    gdDestroyArray mPhraseTypes
    gdDestroyArray mFormatting
    
    mComments.Destroy
    
    Set mEnterKey = Nothing
    Set mStrh = Nothing
    Set mFunctions = Nothing
    Set mInputs = Nothing
    Set mReturnTypes = Nothing
    Set mTrees = Nothing
    
    Set m.astrMacros = Nothing
    
End Sub
