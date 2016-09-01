VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cFunction.cls
'' Description: Wrapper class for an function
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/23/2010   DAJ         Make sure to do a Str around default value for parm
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Private Type mPrivate
    lFunctionID As Long
    strFunctionName As String
    strDescription As String
    lFunctionCategoryID As Long
    bImplementationTypeID As Byte
    bLateCalculating As Boolean
    dLastModified As Date
    strCodedName As String
    strFormatted As String
    strFormattedWithFillWords As String
    strCodedText As String
    bUsage As Byte
    bBarParms As Byte      'Of can be used outside function
    lDataTypeID As Long
    lLibraryID As Long
    Inputs As cInputs
    strTradeSenseUsage As String
    bSecurityLevel As Byte
    bCannotDelete As Boolean
    strPassword As String
    alFunctionIDs As cGdArray
    bReverify As Boolean
    strAbbrev As String
    bUsesOpenNextBar As Boolean
    bUsesHLCNextBar As Boolean
    strRequiredMod As String
    
    'MT 6/2001: TradeSense compatibility =======================
    lListID As Long
    strDLLName As String
    strClassName As String
    strPreview As String
    'MT 6/2001: TradeSense compatibility =======================
    
    'Error processing...
    strErrDesc As String
    bErrNbr As Byte
End Type
Private m As mPrivate

'MT 6/2001: TradeSense compatibility =======================
Property Get ReturnTypeID() As Long
    ReturnTypeID = m.lDataTypeID
End Property
Property Get HiddenInputs() As Byte
    HiddenInputs = m.bBarParms
End Property
Property Get ListID() As Long
    ListID = m.lListID
End Property
Property Get DLLName() As String
    DLLName = m.strDLLName
End Property
Property Get ClassName() As String
    ClassName = m.strCodedName
End Property
Property Get Preview() As String
    Preview = m.strFormattedWithFillWords
End Property
Property Let ReturnTypeID(pData As Long)
    m.lDataTypeID = pData
End Property
Property Let HiddenInputs(pData As Byte)
    m.bBarParms = pData
End Property
Property Let ListID(pData As Long)
    m.lListID = pData
End Property
Property Let DLLName(pData As String)
    m.strDLLName = pData
End Property
Property Let ClassName(pData As String)
    m.strCodedName = pData
End Property
Property Let Preview(pData As String)
    m.strPreview = pData
End Property

Property Get ErrDesc() As String
    ErrDesc = m.strErrDesc
End Property
Property Get ErrNbr() As Byte
    ErrNbr = m.bErrNbr
End Property

Property Get Inputs() As cInputs
    Set Inputs = m.Inputs
End Property
Property Let Inputs(pData As cInputs)
    Set m.Inputs = pData
End Property

Property Get FunctionID() As Long
    FunctionID = m.lFunctionID
End Property
Property Get FunctionName() As String
    FunctionName = m.strFunctionName
End Property
Property Get Description() As String
    Description = m.strDescription
End Property
Property Get FunctionCategoryID() As Long
    FunctionCategoryID = m.lFunctionCategoryID
End Property
Property Get ImplementationTypeID() As Byte
    ImplementationTypeID = m.bImplementationTypeID
End Property
Property Get LateCalculating() As Boolean
    LateCalculating = m.bLateCalculating
End Property
Property Get LastModified() As Date
    LastModified = m.dLastModified
End Property
Property Get CodedName() As String
    CodedName = m.strCodedName
End Property
Property Get Formatted() As String
    Formatted = m.strFormatted
End Property
Property Get FormattedWithFillWords() As String
    FormattedWithFillWords = m.strFormattedWithFillWords
End Property
Property Get CodedText() As String
    CodedText = m.strCodedText
End Property
Property Get Usage() As Byte
    Usage = m.bUsage
End Property
Property Get BarParms() As Byte
    BarParms = m.bBarParms
End Property
Property Get DataTypeID() As Long
    DataTypeID = m.lDataTypeID
End Property
Property Get LibraryID() As Long
    LibraryID = m.lLibraryID
End Property
Property Get SecurityLevel() As Long
    SecurityLevel = m.bSecurityLevel
End Property
Property Get CannotDelete() As Boolean
    CannotDelete = m.bCannotDelete
End Property
Property Get Password() As String
    Password = m.strPassword
End Property
Property Get TradeSenseUsage() As String
    TradeSenseUsage = m.strTradeSenseUsage
End Property
Property Get Reverify() As Boolean
    Reverify = m.bReverify
End Property
Property Get Abbrev() As String
    Abbrev = m.strAbbrev
End Property
Property Get UsesNextBarHLC() As Boolean
    UsesNextBarHLC = m.bUsesHLCNextBar
End Property
Property Get UsesOpenNextBar() As Boolean
    UsesOpenNextBar = m.bUsesOpenNextBar
End Property
Property Get RequiredMod() As String
    RequiredMod = m.strRequiredMod
End Property

Property Let FunctionID(pData As Long)
    m.lFunctionID = pData
End Property
Property Let FunctionName(pData As String)
    m.strFunctionName = pData
End Property
Property Let Description(pData As String)
    m.strDescription = pData
End Property
Property Let FunctionCategoryID(pData As Long)
    m.lFunctionCategoryID = pData
End Property
Property Let ImplementationTypeID(pData As Byte)
    m.bImplementationTypeID = pData
End Property
Property Let LateCalculating(pData As Boolean)
    m.bLateCalculating = pData
End Property
Property Let LastModified(pData As Date)
    m.dLastModified = pData
End Property
Property Let CodedName(pData As String)
    m.strCodedName = pData
End Property
Property Let Formatted(pData As String)
    m.strFormatted = pData
End Property
Property Let FormattedWithFillWords(pData As String)
    m.strFormattedWithFillWords = pData
End Property
Property Let CodedText(pData As String)
    m.strCodedText = pData
End Property
Property Let Usage(pData As Byte)
    m.bUsage = pData
End Property
Property Let BarParms(pData As Byte)
    m.bBarParms = pData
End Property
Property Let DataTypeID(pData As Long)
    m.lDataTypeID = pData
End Property
Property Let SecurityLevel(pData As Long)
    m.bSecurityLevel = pData
End Property
Property Let CannotDelete(pData As Boolean)
    m.bCannotDelete = pData
End Property
Property Let Password(pData As String)
    m.strPassword = pData
End Property
Property Let TradeSenseUsage(pData As String)
    m.strTradeSenseUsage = pData
End Property
Property Let LibraryID(pData As Long)
    m.lLibraryID = pData
End Property
Property Let FunctionIDs(pData As Long)
    m.alFunctionIDs.CopyFromHandle pData
End Property
Property Let Reverify(pData As Boolean)
    m.bReverify = pData
End Property
Property Let Abbrev(pData As String)
    m.strAbbrev = pData
End Property
Property Let UsesNextBarHLC(pData As Boolean)
    m.bUsesHLCNextBar = pData
End Property
Property Let UsesOpenNextBar(ByVal bData As Boolean)
    m.bUsesOpenNextBar = bData
End Property
Property Let RequiredMod(ByVal strRequiredMod As String)
    m.strRequiredMod = strRequiredMod
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetRTF
'' Description: Get the RTF version of some rule text
'' Inputs:      Text to convert
'' Returns:     RTF text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetRTF(strRuleText As String) As String
On Error GoTo ErrSection:

    GetRTF = Color(strRuleText)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFunction.GetRTF", eGDRaiseError_Raise, g.strAppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Validate
'' Description: Validate the function information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Validate()
On Error GoTo ErrSection:

    Dim strErrDesc As String            ' Error description
    
    ' Function name must be 50 characters or less
    m.bErrNbr = 1
    If (Len(m.strFunctionName) <= 0 Or Len(m.strFunctionName) > 50) Then
        Err.Raise vbObjectError + 1000, , _
            "Function name must be 1 to 50 characters in length and " & _
                "consist of letters, numbers, and spaces."
    End If
    
    ' First character must be alphabetic
    If Not FirstCharValid(m.strFunctionName, strErrDesc) Then
        Err.Raise vbObjectError + 1000, , strErrDesc
    End If
        
    ' Make sure single quotes not in system name
    If InStr(1, m.strFunctionName, "'") > 0 Then
        Err.Raise vbObjectError + 1000, , _
            "Single quotes not allowed in the Function Name"
    End If

    ' Remaining characters must be alphabetic or numeric
    If Not RemainCharsValid(m.strFunctionName, strErrDesc) Then
        Err.Raise vbObjectError + 1000, , strErrDesc
    End If
    
    ' Alphabetic operators are not allowed
    If OperatorsFound(m.strFunctionName, strErrDesc) Then
        Err.Raise vbObjectError + 1000, , strErrDesc
    End If
    
    ' Codedname validation
    m.bErrNbr = 3
    If m.bImplementationTypeID = 1 Then
        ' Name must be 50 characters or less
        If (Len(m.strCodedName) = 0 Or Len(m.strCodedName) > 50) Then
            Err.Raise vbObjectError + 1000, , _
                "DLL name is invalid.  It cannot be longer than 50 characters."
        End If
        
        ' No imbedded blanks allowed
        If InStr(1, m.strCodedName, " ") > 0 Then
            Err.Raise vbObjectError + 1000, , _
                "DLL name is invalid.  Spaces are not allowed."
        End If
        
        ' First character must be alphabetic
        If Not FirstCharValid(m.strCodedName, strErrDesc) Then
            Err.Raise vbObjectError + 1000, , strErrDesc
        End If
        
        ' Remaining characters must be alphabetic or numeric
        If Not RemainCharsValid(m.strCodedName, strErrDesc) Then
            Err.Raise vbObjectError + 1000, , strErrDesc
        End If
        
        m.strFormatted = "N/A"
        m.strFormattedWithFillWords = "N/A"
    Else
#If 0 Then
        '(Type is 2 for custom coded text functions)
        'Determine usage (System testing function or MM function)
        m.bUsage = C_Both
        For X = 1 To m.Inputs.Count
            If m.Inputs.Item(X).ParmTypeID = C_RetBars Then
                m.bUsage = C_System
                Exit For
            End If
            If m.Inputs.Item(X).ParmTypeID = C_RetTrades Then
                m.bUsage = C_MM
                Exit For
            End If
        Next X
#End If
    End If

    ' Coded Text was not entered for a Coded Text function...
    m.bErrNbr = 4
    If m.bImplementationTypeID = 2 Then
        If Len(m.strFormatted) = 0 Then
            Err.Raise vbObjectError + 1000, , "Please enter an expression for the function"
        End If
    End If

    ' Function Category ID validation
    m.bErrNbr = 5
    If m.lFunctionCategoryID = 0 Then
        Err.Raise vbObjectError + 1000, , "Function Category must be selected"
    End If

    ' Library ID validation
    m.bErrNbr = 6
    If m.lLibraryID = 0 Then
        Err.Raise vbObjectError + 1000, , "Library must be selected"
    End If

    ' Make sure no duplicate inputs exist
    ' make sure no more than 10 inputs exist (for money management)
    m.bErrNbr = 7
    'If m.Inputs.Count > 10 And m.bUsage = 1 Then
    '    Err.Raise vbObjectError + 1000, , _
    '        "No more than 10 inputs are allowed for a money " & _
    '            "management function"
    'End If
    
    ' Short cut already exists
    m.bErrNbr = 8
    If Len(m.strAbbrev) > 0 Then
        'Set rs = dbNav.OpenRecordset("Select * from [tblFunctions] " & _
            "Where [Abbrev]='" & m.strAbbrev & "' Order by [Abbrev];", dbOpenSnapshot)
        'If Not rs.EOF Then
        '    Err.Raise vbObjectError + 1000, , _
        '        "Function Short cut " & UCase(m.strAbbrev) & " already exists."
        'End If
        
        'Function short cuts (abbreviations) cannot be numeric
        If IsNumeric(m.strAbbrev) Then
            Err.Raise vbObjectError + 1000, , "Function short cuts (abbreviations) cannot be numeric"
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFunction.Validate", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidateInputs
'' Description: Validate each function input
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateInputs()
On Error GoTo ErrSection:
    
    Dim lIndex As Integer               ' Index into a for loop
    
    For lIndex = 1 To m.Inputs.Count
        With m.Inputs.Item(lIndex)
            ' Input name...
            m.bErrNbr = 20
            If Len(.ParmName) = 0 Then
                Err.Raise vbObjectError + 1001, , "An Input name must be entered"
            End If
            If Len(.ParmName) > 50 Then
                Err.Raise vbObjectError + 1001, , "Input name cannot exceed 50 characters"
            End If
        
            ' Make sure delimiter for builtin in functions is not specified in
            ' input name
            If InStr(1, .ParmName, "^") > 0 Then
                Err.Raise vbObjectError + 1001, , "Input name cannot contain the '^' character"
            End If
            
            'Verify that backets are around optional inputs
            'If Not mRequired Then
            '    If Left(.ParmName, 1) <> "[" Then .ParmName = "[" & .ParmName
            '    If Right(.ParmName, 1) <> "]" Then .ParmName = .ParmName & "]"
            'End If
            
            Select Case .ParmTypeID
                Case C_RetNumericConstant, C_RetNumeric, _
                     C_RetTrueFalseConstant
                    If .Required Then
                        If .FromValue <> 0 And .ToValue <> 0 Then
                            If Val(.DefaultValue) < .FromValue Or _
                               Val(.DefaultValue) > .ToValue Then
                                m.bErrNbr = 21
                                Err.Raise vbObjectError + 1002, , _
                                    "The Default Value must be between " & Format(.FromValue, "general number") & _
                                    " and " & Format(.ToValue, "general number")
                            End If
                        End If
                    End If
                
                Case C_RetTrades
                    If .ParmName = "Trades" Then .DefaultValue = "Trades"
                    .DefaultValue = .ParmName
                
                Case C_RetBars
                    If .ParmName = "Market1" Then .DefaultValue = "Market1"
                    .DefaultValue = .ParmName
                
                Case C_RetText
                    'If m.lListID = 0 Then
                    '    m.bErrNbr = 4
                    '    Err.Raise vbObjectError + 1000, , "Please select a list"
                    'End If
                    
                Case 0
                    Err.Raise vbObjectError + 1000, , "Please select an Input type"
            End Select
            
            ' Fill pre/pos words
            If Len(.FillPre) > 30 Then
                m.bErrNbr = 22
                Err.Raise vbObjectError + 1003, , _
                    "Fill words before value must be 30 characters or less"
            End If
            If Len(.FillPost) > 30 Then
                m.bErrNbr = 23
                Err.Raise vbObjectError + 1004, , _
                    "Fill words after value must be 30 characters or less"
            End If
        
        End With
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFunction.ValidateInputs", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load a function from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
            "WHERE [FunctionID]=" & Str(m.lFunctionID) & ";", dbOpenSnapshot)
    If rs.EOF Then
        m.lFunctionID = 0
    Else
        If rs!CheckSum = BuildCheckSum(rs, "tblFunctions") Then
            m.strFunctionName = rs!FunctionName
            m.strDescription = rs!Description
            m.lFunctionCategoryID = rs!FunctionCategoryID
            m.bImplementationTypeID = rs!ImplementationTypeID
            m.bLateCalculating = rs!LateCalculating
            m.dLastModified = rs!LastModified
            m.strCodedName = rs!CodedName
            m.strFormatted = DecryptField(rs!Formatted)
            m.strFormattedWithFillWords = DecryptField(rs!FormattedWithFillWords)
            m.strCodedText = DecryptField(rs!CodedText)
            m.bReverify = rs!Reverify
            m.bUsage = rs!Usage
            m.lDataTypeID = rs!ReturnTypeID
            m.lLibraryID = rs!LibraryID
            m.bSecurityLevel = NullChk(rs!SecurityLevel, 0)
            m.bCannotDelete = rs!CannotDelete
            m.strPassword = DecryptField(rs!Password)
            m.strTradeSenseUsage = NullChk(rs!TradeSenseUsage)
            m.strAbbrev = NullChk(rs!Abbrev)
            m.bUsesOpenNextBar = rs!UsesOpenNextBar
            m.bUsesHLCNextBar = rs!UsesHLCNextBar
            
            'MT 6/2001: TradeSense compatibility =======================
            m.lListID = rs!ListID
            m.strPreview = NullChk(rs!FormattedWithFillWords)
            
            m.strRequiredMod = NullChk(rs!RequiredMod)
            
            LoadInputs
        Else
            Err.Raise vbObjectError + 1000, , "Function cannot be loaded because it is no longer valid"
        End If
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cFunction.Load", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save a function to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim bNewRec As Boolean              ' Is this a new function?

    Validate
    ValidateInputs
    
    'Update the table first...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
        "WHERE [FunctionID]=" & Str(m.lFunctionID) & ";", dbOpenDynaset)
    With rs
        If .EOF Then
            .AddNew
            bNewRec = True
        Else
            .Edit
            bNewRec = False
        End If
            
        !FunctionName = m.strFunctionName
        !Description = m.strDescription
        !FunctionCategoryID = m.lFunctionCategoryID
        !ImplementationTypeID = m.bImplementationTypeID
        !LateCalculating = m.bLateCalculating
        If m.dLastModified = 0 Then !LastModified = Now() Else !LastModified = m.dLastModified
        !CodedName = m.strCodedName
        EncryptField !Formatted, m.strFormatted
        EncryptField !FormattedWithFillWords, m.strFormattedWithFillWords
        EncryptField !CodedText, m.strCodedText
        !ReturnTypeID = m.lDataTypeID
        !Usage = m.bUsage
        !LibraryID = m.lLibraryID
        !SecurityLevel = m.bSecurityLevel
        !CannotDelete = m.bCannotDelete
        EncryptField !Password, m.strPassword
        !TradeSenseUsage = m.strTradeSenseUsage
        !Reverify = m.bReverify
        !Abbrev = m.strAbbrev
        !UsesOpenNextBar = m.bUsesOpenNextBar
        !UsesHLCNextBar = m.bUsesHLCNextBar
        
        'MT 6/2001: TradeSense compatibility =======================
        !ListID = m.lListID
        
        If Len(m.strRequiredMod) = 0 Then rs!RequiredMod = Null Else rs!RequiredMod = m.strRequiredMod
        
        !CheckSum = BuildCheckSum(rs, "tblFunctions")
        .Update
        
        'Add new item to collection or update existing item...
        If bNewRec Then
            rs.MoveLast
            m.lFunctionID = rs!FunctionID
        End If
    End With
    rs.Close
        
    UpdateReferences
    SaveInputs
    
    ' For now, reload the function so that all of the properties get set correctly
    Load
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Select Case Err.Number
        Case 3022
            Err.Raise vbObjectError + 1000, , "This function already exists.  Verify that the function name and coded name are unique.  " & Chr(13) & Chr(10) & _
                "Function Name: " & m.strFunctionName & Chr(13) & Chr(10) & _
                "Coded Name: " & m.strCodedName
        Case Else
            RaiseError "cFunction.Save", eGDRaiseError_Raise, g.strAppPath
    End Select
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveInputs
'' Description: Save the function inputs
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveInputs()
On Error GoTo ErrSection:
    
    Dim lIndex As Integer               ' Index into a for loop
    Dim bFound As Boolean               ' Was the input found?
    Dim rs As Recordset                 ' Recordset into the database
    Dim bChanged As Boolean             ' Is this a new input?
    
    'No inputs found, exit...
    If m.Inputs Is Nothing Then Exit Sub
        
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionParms] " & _
        "WHERE [FunctionID]=" & Str(m.lFunctionID) & ";", dbOpenDynaset)
    
    ' Write all new inputs from Inputs collection to tblRuleParms
    For lIndex = 1 To m.Inputs.Count
        With m.Inputs.Item(lIndex)
            rs.FindFirst "[ParmText]='" & .ParmName & "'"
            If rs.NoMatch Then
                bChanged = True
                rs.AddNew
                rs!FunctionID = m.lFunctionID
            Else
                If rs!ParmNbr <> .ParmSeq Then bChanged = True
                If rs!ParmTypeID <> .ParmTypeID Then bChanged = True
                If rs!DefaultValue <> Str(.DefaultValue) Then bChanged = True
                'If rs!FromValue <> .FromValue Then bChanged = True
                'If rs!ToValue <> .ToValue Then bChanged = True
                If rs!Required <> .Required Then bChanged = True
                rs.Edit
            End If
            rs!ParmNbr = .ParmSeq
            rs!ParmText = .ParmName
            rs!ParmDesc = .ParmDesc
            rs!ParmTypeID = .ParmTypeID
            rs!DefaultValue = Str(.DefaultValue)
            'rs!FromValue = .FromValue
            'rs!ToValue = .ToValue
            rs!Required = .Required
            rs!ListID = .ListID
            rs!FillPre = .FillPre
            'rs!FillPost = .FillPost
            rs.Update
        End With
    Next lIndex
    
    ' Delete old inputs from table
    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            bFound = False
            For lIndex = 1 To m.Inputs.Count
                If rs!ParmText = m.Inputs.Item(lIndex).ParmName Then
                    bFound = True
                    Exit For
                End If
            Next lIndex
            If Not bFound Then
                bChanged = True
                rs.Delete
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    If bChanged Then
        'Only if inputs have changed do we need to flag the dependent functions to be reverified
        ReverifyParents
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cFunction.SaveInputs", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateReferences
'' Description: Update Function references table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateReferences()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Integer               ' Index into a for loop
    Dim bFound As Boolean               ' Was the function found?
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionRefs] " & _
        "WHERE [FunctionID]=" & Str(m.lFunctionID) & " " & _
        "ORDER BY [FunctionID],[FunctionIDref];", dbOpenDynaset)
    
    If m.alFunctionIDs Is Nothing Then Exit Sub
    
    ' Add new functions to the references table
    For lIndex = 0 To m.alFunctionIDs.Size - 1
        rs.FindFirst "[FunctionID]=" & Str(m.lFunctionID) & " and " & _
                     "[FunctionIDRef]=" & Str(m.alFunctionIDs(lIndex))
        If rs.NoMatch Then
            rs.AddNew
            rs!FunctionID = m.lFunctionID
            rs!FunctionIDRef = m.alFunctionIDs(lIndex)
            rs.Update
        End If
    Next lIndex
    
    ' If there are records in the recordset, go back to the beginning of the
    ' recordset so that we can walk through it to see if there are any old references
    ' that can be removed (04/30/2008 DAJ)...
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        
        'Delete any old references
        Do Until rs.EOF
            bFound = False
            For lIndex = 0 To m.alFunctionIDs.Size - 1
                If rs!FunctionIDRef = m.alFunctionIDs(lIndex) Then
                    bFound = True
                    Exit For
                End If
            Next lIndex
            If Not bFound Then
                rs.Delete
            End If
            rs.MoveNext
        Loop
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set rs = Nothing
    RaiseError "cFunction.UpdateReferences", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReverifyParents
'' Description: Set Reverify flag to current functions or rules that use this
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ReverifyParents()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    ' Set on the Reverify flag of the functions using this functions
    Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.*, tblFunctionRefs.* " & _
                "FROM tblFunctions INNER JOIN tblFunctionRefs ON tblFunctions.FunctionID = tblFunctionRefs.FunctionID " & _
                "WHERE (((tblFunctionRefs.FunctionIDRef)=" & Str(m.lFunctionID) & "));", dbOpenDynaset)
    ValidateCheckSums rs, "tblFunctions"
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do While Not rs.EOF
        If rs!CheckSum <> 0.5 Then
            rs.Edit
            rs!Reverify = True
            rs!CheckSum = BuildCheckSum(rs, "tblFunctions")
            rs.Update
        End If
        rs.MoveNext
    Loop
    rs.Close

    ' Set on the Reverify flag of the rules using this function
    Set rs = g.dbNav.OpenRecordset("SELECT tblRules.*, tblFunctionRules.* " & _
                "FROM tblRules INNER JOIN tblFunctionRules ON tblRules.RuleID = tblFunctionRules.RuleID " & _
                "WHERE (((tblFunctionRules.FunctionIDRef)=" & Str(m.lFunctionID) & "));", dbOpenDynaset)
    ValidateCheckSums rs, "tblRules"
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do While Not rs.EOF
        If rs!CheckSum <> 0.5 Then
            rs.Edit
            rs!Reverify = True
            rs!CheckSum = BuildCheckSum(rs, "tblRules")
            rs.Update
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    ' Set on the Reverify flag of the systems using rules with the
    ' reverify flag set on
    Set rs = g.dbNav.OpenRecordset("SELECT tblSystems.*, tblSystemRules.*, tblRules.* " & _
                "FROM tblSystems INNER JOIN (tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID) ON tblSystems.SystemNumber = tblSystemRules.SystemNumber " & _
                "WHERE (((tblRules.Reverify)=True));", dbOpenDynaset)
    ValidateCheckSums rs, "tblSystems"
    ValidateCheckSums rs, "tblSystemRules"
    ValidateCheckSums rs, "tblRules"
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do While Not rs.EOF
        If rs![tblSystems.CheckSum] <> 0.5 And rs![tblSystemRules.CheckSum] <> 0.5 And rs![tblRules.CheckSum] <> 0.5 Then
            rs.Edit
            rs![tblSystems.Reverify] = True
            rs![tblSystems.CheckSum] = BuildCheckSum(rs, "tblSystems")
            rs.Update
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    ' Set on the Reverify flag of the mm strategies using rules with the
    ' reverify flag set on
    g.dbNav.Execute "UPDATE tblMM " & _
        "INNER JOIN (tblRules INNER JOIN tblMMRules ON tblRules.RuleID = tblMMRules.RuleID) ON tblMM.MMID = tblMMRules.MMid " & _
        "SET tblMM.Reverify = True " & _
        "WHERE (((tblRules.Reverify)=True));"
      
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    '3078 is "table not found".  This allows the money management app
    'to use this DLL.  The table "tblMM" will not exist in System Nav.
    'and will bomb with 3078 when executing a query for money mgt tables.
    If Err.Number = 3078 Then Resume Next
    Set rs = Nothing
    RaiseError "cFunction.ReverifyParents", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadInputs
'' Description: Load function inputs from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadInputs()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset from the database
    
    Set m.Inputs = New cInputs
    
    Set rs = g.dbNav.OpenRecordset("SELECT tblFunctionParms.*, tblLists.ListName, tblParmTypes.ParmType, tblParmTypes.ParmType, tblParmTypes.ParmUsage " & _
        "FROM tblLists RIGHT JOIN (tblParmTypes RIGHT JOIN tblFunctionParms ON tblParmTypes.ParmTypeID = tblFunctionParms.ParmTypeID) ON tblLists.ListID = tblFunctionParms.ListID " & _
        "WHERE (((tblFunctionParms.FunctionID) = " & Str(m.lFunctionID) & ")) " & _
        "ORDER BY tblFunctionParms.FunctionID, tblFunctionParms.ParmNbr;", dbOpenSnapshot)
    
    With rs
        m.bBarParms = 0
        Do Until .EOF
            'Count the number bar structure type parms and save this number at the
            'function level.  It is used by TradeSense to determine when the operator
            '"OF" is valid
            If !ParmTypeID = C_RetBars Or !ParmTypeID = C_RetTrades Then
                m.bBarParms = m.bBarParms + 1
            End If
            m.Inputs.Add "", !ParmNbr, !ParmText, !ParmDesc, !ParmID, 0, False, 0, 0, 0, 0, 0, !ParmTypeID, _
                  !DefaultValue, !Required, !ParmUsage, !ParmType, !ListID, !FillPre, !ListName
            
            .MoveNext
        Loop
    End With
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cFunction.LoadInputs", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadInput
'' Description: Add an input to the collection
'' Inputs:      Input Properties
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadInput(pParmID As Long, pParmNbr As Long, _
    pParmText As Variant, pParmDesc As Variant, _
    pParmTypeID As Variant, pListID As Long, pRequired As Boolean, _
    pFromValue As Double, pToValue As Double, _
    pDefaultValue As Variant, pFillPre As Variant, pFillPost As Variant)
On Error GoTo ErrSection:
    
    m.Inputs.Add "", pParmNbr, pParmText, pParmDesc, pParmID, 0, False, 0, 0, 0, 0, 0, pParmTypeID, _
                  pDefaultValue, pRequired, pFromValue, pToValue, pListID, pFillPre, pFillPost
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFunction.LoadInput", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveInput
'' Description: Remove the given input from the collection
'' Inputs:      ID of the Input to Remove
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveInput(ByVal lParmID As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionParms] " & _
                "WHERE [ParmID]=" & Str(lParmID) & ";", dbOpenDynaset)
    
    If rs.EOF Then
        Err.Raise vbObjectError + 1000, , "Delete failed because input does not exist"
    Else
        rs.Delete
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cFunction.RemoveInput", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Delete
'' Description: Delete the function from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
            "WHERE [FunctionID]=" & Str(m.lFunctionID) & ";", dbOpenDynaset)
    With rs
        If rs.EOF Then
            Err.Raise vbObjectError + 1000, , "Delete failed because function does not exist"
        Else
            ReverifyParents
            rs.Delete
        End If
    End With
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cFunction.Delete", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub Class_Initialize()
    
    Set m.Inputs = New cInputs
    Set m.alFunctionIDs = New cGdArray
    m.alFunctionIDs.Create eGDARRAY_Longs

End Sub

Private Sub Class_Terminate()
    
    Set m.Inputs = Nothing
    Set m.alFunctionIDs = Nothing

End Sub

Public Function ValidName(ByVal strFunctionName As String) As String
On Error GoTo ErrSection:

    Dim strErrDesc As String            ' Error description
    
    ' Function name must be 50 characters or less
    If (Len(strFunctionName) <= 0 Or Len(strFunctionName) > 50) Then
        ValidName = "Function name must be 1 to 50 characters in length and " & _
                "consist of letters, numbers, and spaces."
    
    ' First character must be alphabetic
    ElseIf Not FirstCharValid(strFunctionName, strErrDesc) Then
        ValidName = strErrDesc
        
    ' Make sure single quotes not in name
    ElseIf InStr(1, strFunctionName, "'") > 0 Then
        ValidName = "Single quotes not allowed in the Function Name"

    ' Make sure comma not in name
    ElseIf InStr(strFunctionName, ",") > 0 Then
        ValidName = "Comma is not allowed in the Function Name"
    
    ' Remaining characters must be alphabetic or numeric
    ElseIf Not RemainCharsValid(strFunctionName, strErrDesc) Then
        ValidName = strErrDesc
    
    ' Alphabetic operators are not allowed
    ElseIf OperatorsFound(strFunctionName, strErrDesc) Then
        ValidName = strErrDesc
    
    ' Otherwise, the name is valid...
    Else
        ValidName = ""
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFunction.ValidName", eGDRaiseError_Raise, g.strAppPath

End Function
