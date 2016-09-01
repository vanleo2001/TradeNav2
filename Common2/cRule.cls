VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cRule.cls
'' Description: Wrapper class for an rule
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/23/2010   DAJ         Make sure to do a Str around default value for a parm
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Private Type mPrivate
    ' System Rule stuff...
    bSelected As Boolean
    bAlternate As Boolean
    iRuleUse As Integer
    dLastModKnown As Date
    lUnitsID As Long
    strLinkedRules As String
    bSysExitBasedOnEachTrade As Boolean
    lSysNumContracts As Long
    bSysAsPercentOfPosition As Boolean
    
    ' Rule stuff...
    lRuleID As Long
    strName As String
    strCond As String
    strCondFillWords As String
    strCondCoded As String
    strActionCodedName As String
    lLibraryID As Long
    dCreated As Date
    dLastModified As Date
    bDeleted As Boolean
    bBuySell As Boolean
    iRuleType As Byte
    bLateCondition As Boolean
    bLateAction As Boolean
    strOrderPlacement As String
    strPrice As String
    strPriceCoded As String
    nSecurityLevel As Byte
    bCannotDelete As Boolean
    strPassword As String
    iSeq As Integer
    strPrice2RTF As String
    strPrice2Coded As String
    bReverify As Boolean
    bExitOnEntryBar As Boolean
    bExitBasedOnEachTrade As Boolean
    lNumberContracts As Long
    bAsPercentOfPosition As Boolean
    lSystemNumber As Long
    lCategoryID As Long
    
    alFunctionIDs As cGdArray
    alCondRefs As cGdArray
    alPrice1Refs As cGdArray   'Out dated
    alPrice2Refs As cGdArray   'Out dated
    
    'Error Handling
    nError As Byte
    
    'Collections required for building parameter list...
    Inputs As cInputs
End Type
Private m As mPrivate

Property Get ErrNbr() As Byte
    ErrNbr = m.nError
End Property

'Properties used to control/expose parameters...
Property Get Inputs() As cInputs
    Set Inputs = m.Inputs
End Property
Property Let Inputs(pData As cInputs)
    Set m.Inputs = pData
End Property


Property Get Selected() As Boolean
    Selected = m.bSelected
End Property
Property Get Alternate() As Boolean
    Alternate = m.bAlternate
End Property
Property Get LastModKnown() As Date
    LastModKnown = m.dLastModKnown
End Property
Property Get LastModified() As Date
    LastModified = m.dLastModified
End Property
Property Get RuleID() As Long
    RuleID = m.lRuleID
End Property
Property Get Name() As String
    Name = m.strName
End Property
Property Get Cond() As String
    Cond = m.strCond
End Property
Property Get CondFillWords() As String
    CondFillWords = m.strCondFillWords
End Property
Property Get CondCoded() As String
    CondCoded = m.strCondCoded
End Property
Property Get ActionCodedName() As String
    ActionCodedName = m.strActionCodedName
End Property
Property Get LibraryID() As Long
    LibraryID = m.lLibraryID
End Property
Property Get Created() As Date
    Created = m.dCreated
End Property
Property Get Deleted() As Boolean
    Deleted = m.bDeleted
End Property
Property Get RuleUse() As Integer
    RuleUse = m.iRuleUse
End Property
Property Get LinkedRules() As String
    LinkedRules = m.strLinkedRules
End Property
Property Get BuySell() As Boolean
    BuySell = m.bBuySell
End Property
Property Get RuleType() As Byte
    RuleType = m.iRuleType
End Property
Property Get LateCondition() As Boolean
    LateCondition = m.bLateCondition
End Property
Property Get LateAction() As Boolean
    LateAction = m.bLateAction
End Property
Property Get OrderPlacement() As String
    OrderPlacement = m.strOrderPlacement
End Property
Property Get Price() As String
    Price = m.strPrice
End Property
Property Get PriceCoded() As String
    PriceCoded = m.strPriceCoded
End Property
Property Get SecurityLevel() As Byte
    SecurityLevel = m.nSecurityLevel
End Property
Property Get CannotDelete() As Boolean
    CannotDelete = m.bCannotDelete
End Property
Property Get Password() As String
    Password = m.strPassword
End Property
Property Get Seq() As Integer
    Seq = m.iSeq
End Property
Property Get Price2RTF() As String
    Price2RTF = m.strPrice2RTF
End Property
Property Get Price2Coded() As String
    Price2Coded = m.strPrice2Coded
End Property
Property Get Reverify() As Boolean
    Reverify = m.bReverify
End Property
Property Get ExitOnEntryBar() As Boolean
    ExitOnEntryBar = m.bExitOnEntryBar
End Property
Property Get ExitBasedOnEachTrade() As Boolean
    ExitBasedOnEachTrade = m.bExitBasedOnEachTrade
End Property
Property Get NumberContracts() As Long
    NumberContracts = m.lNumberContracts
End Property
Property Get AsPercentOfPosition() As Boolean
    AsPercentOfPosition = m.bAsPercentOfPosition
End Property
Property Get SystemNumber() As Long
    SystemNumber = m.lSystemNumber
End Property
Public Property Get PreviewRTF() As String
    PreviewRTF = RulePreview(m.bBuySell, m.strOrderPlacement, m.strCond, m.strPrice, m.strPrice2RTF)
End Property
Public Property Get SysExitBasedOnEachTrade() As Boolean
    SysExitBasedOnEachTrade = m.bSysExitBasedOnEachTrade
End Property
Public Property Get SysNumContracts() As Long
    SysNumContracts = m.lSysNumContracts
End Property
Public Property Get SysAsPercentOfPosition() As Boolean
    SysAsPercentOfPosition = m.bSysAsPercentOfPosition
End Property
Public Property Get CondRefs() As Long
    CondRefs = m.alCondRefs.ArrayHandle
End Property
Public Property Get CategoryID() As Long
    CategoryID = m.lCategoryID
End Property

Property Let Selected(pData As Boolean)
    m.bSelected = pData
End Property
Property Let Alternate(pData As Boolean)
    m.bAlternate = pData
End Property
Property Let LastModKnown(pData As Date)
    m.dLastModKnown = pData
End Property
Property Let LastModified(pData As Date)
    m.dLastModified = pData
End Property
Property Let RuleID(pData As Long)
    m.lRuleID = pData
End Property
Property Let Name(pData As String)
    m.strName = pData
End Property
Property Let Cond(pData As String)
    m.strCond = pData
End Property
Property Let CondFillWords(pData As String)
    m.strCondFillWords = pData
End Property
Property Let CondCoded(pData As String)
    m.strCondCoded = pData
End Property
Property Let ActionCodedName(pData As String)
    m.strActionCodedName = pData
End Property
Property Let LibraryID(pData As Long)
    m.lLibraryID = pData
End Property
Property Let Created(pData As Date)
    m.dCreated = pData
End Property
Property Let Deleted(pData As Boolean)
    m.bDeleted = pData
End Property
Property Let RuleUse(pData As Integer)
    m.iRuleUse = pData
End Property
Property Let LinkedRules(pData As String)
    m.strLinkedRules = pData
End Property
Property Let BuySell(pData As Boolean)
    m.bBuySell = pData
End Property
Property Let RuleType(pData As Byte)
    m.iRuleType = pData
End Property
Property Let LateCondition(pData As Boolean)
    m.bLateCondition = pData
End Property
Property Let LateAction(pData As Boolean)
    m.bLateAction = pData
End Property
Property Let OrderPlacement(pData As String)
    m.strOrderPlacement = pData
End Property
Property Let Price(pData As String)
    m.strPrice = pData
End Property
Property Let PriceCoded(pData As String)
    m.strPriceCoded = pData
End Property
Property Let SecurityLevel(pData As Byte)
    m.nSecurityLevel = pData
End Property
Property Let CannotDelete(pData As Boolean)
    m.bCannotDelete = pData
End Property
Property Let Password(pData As String)
    m.strPassword = pData
End Property
Property Let Seq(pData As Integer)
    m.iSeq = pData
End Property
Property Let Price2RTF(pData As String)
    m.strPrice2RTF = pData
End Property
Property Let Price2Coded(pData As String)
    m.strPrice2Coded = pData
End Property
Property Let Reverify(pData As Boolean)
    m.bReverify = pData
End Property

'ID's of functions used in condition and price expressions...
Property Let CondRefs(ByVal pData As Long)
    Set m.alCondRefs = New cGdArray
    m.alCondRefs.CopyFromHandle pData
End Property
Property Let Price1Refs(ByVal pData As Long)
    m.alPrice1Refs.CopyFromHandle pData
End Property
Property Let Price2Refs(ByVal pData As Long)
    m.alPrice2Refs.CopyFromHandle pData
End Property
Property Let ExitOnEntryBar(pData As Boolean)
    m.bExitOnEntryBar = pData
End Property
Property Let ExitBasedOnEachTrade(pData As Boolean)
    m.bExitBasedOnEachTrade = pData
End Property
Property Let NumberContracts(pData As Long)
    m.lNumberContracts = pData
End Property
Property Let AsPercentOfPosition(pData As Boolean)
    m.bAsPercentOfPosition = pData
End Property
Property Let SystemNumber(pData As Long)
    m.lSystemNumber = pData
End Property
Public Property Let SysExitBasedOnEachTrade(ByVal bSysExitBasedOnEachTrade As Boolean)
    m.bSysExitBasedOnEachTrade = bSysExitBasedOnEachTrade
End Property
Public Property Let SysNumContracts(ByVal lNumContracts As Long)
    m.lSysNumContracts = lNumContracts
End Property
Public Property Let SysAsPercentOfPosition(ByVal bAsPercentOfPosition As Boolean)
    m.bSysAsPercentOfPosition = bAsPercentOfPosition
End Property
Public Property Let CategoryID(ByVal lCategoryID As Long)
    m.lCategoryID = lCategoryID
End Property

Public Function GetRTF(ByVal strRuleText As String) As String
On Error GoTo ErrSection:

    GetRTF = Color(strRuleText)
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cRule.GetRTF", eGDRaiseError_Raise, g.strAppPath

End Function

Public Sub Validate()
On Error GoTo ErrSection:
    
    Dim strErrorMsg As String
    
    'Rule name
    If (Len(m.strName) <= 0 Or Len(m.strName) > 50) Then
        Err.Raise vbObjectError + 1000, , _
            "Rule name must be 1 to 50 characters in length and " & _
            "consist of letters, numbers, and blanks."
        m.nError = 1
    End If
    
    'First character must be alphabetic
    If Not FirstCharValid(m.strName, strErrorMsg) Then
        Err.Raise vbObjectError + 1000, , strErrorMsg
        m.nError = 1
    End If
    
    'Single quote's not allowed in rule names
    If InStr(1, m.strName, "'") > 0 Then
        Err.Raise vbObjectError + 1000, , "Rule name cannot contain single quotes (')"
        m.nError = 1
    End If

    'Condition
    If Len(m.strCond) = 0 Then
        Err.Raise vbObjectError + 1000, , "A condition must be entered."
        m.nError = 2
    End If
    
    'Price validation
    If m.strOrderPlacement = "Stop" Or m.strOrderPlacement = "Limit" Or _
       m.strOrderPlacement = "Stop with Limit" Then
        If Len(m.strPrice) = 0 Then
            Err.Raise vbObjectError + 1000, , "A price must be entered."
            m.nError = 3
        End If
    End If
    
    'Limit Price validation
    If m.strOrderPlacement = "Stop with Limit" Then
        If Len(m.strPrice2RTF) = 0 Then
            Err.Raise vbObjectError + 1000, , _
                "A limit price must be entered."
            m.nError = 4
        End If
    End If

' Because this is impossible with the user interface and we may want this to
' happen when importing a library, take this out for now.  9/13/2002 DAJ
#If 0 Then
    'If Rule is NOT global, then verify that it only is part of one sys
    Dim rs As Recordset
    Dim MultipleSystemsFnd  As Boolean
    Dim svNbr As Long
    
    If m.lSystemNumber <> 0 Then
        Set rs = dbNav.OpenRecordset("Select * from [tblSystemRules] Where [RuleID]=" & m.lRuleID & " Order by [RuleID],[SystemNumber]", dbOpenSnapshot)
        Do Until rs.EOF
            If svNbr <> 0 Then
                If rs!SystemNumber <> svNbr Then
                    MultipleSystemsFnd = True
                    Exit Do
                End If
            End If
            svNbr = rs!SystemNumber
            rs.MoveNext
        Loop
        If MultipleSystemsFnd Then
            Err.Raise vbObjectError + 1000, , "Rule: " & UCase(m.strName) & " is being used by two or more strategies.  The 'Make available to other strategies' check box must be checked.  Remove the rule from the other strategies using it in order to use this option."
        End If
    End If
#End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRule.Validate", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub Load()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    
    'Update the table first...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
        "WHERE [RuleID]=" & Str(m.lRuleID) & ";", dbOpenSnapshot)
    With rs
        If .EOF Then
            m.lRuleID = 0
        Else
            If rs!CheckSum = BuildCheckSum(rs, "tblRules") Then
                m.strName = rs!Name
                m.strCond = DecryptField(rs!CondRTF)
                m.strCondFillWords = DecryptField(rs!PreviewRTF)
                m.strCondCoded = DecryptField(rs!CondCoded)
                m.strActionCodedName = rs!ActionCodedName
                m.bBuySell = rs!BuySell
                m.iRuleType = rs!RuleType
                m.strOrderPlacement = rs!OrderPlacement
                m.strPrice = DecryptField(rs!PriceRTF)
                m.strPriceCoded = DecryptField(rs!PriceCoded)
                m.lLibraryID = rs!LibraryID
                m.dLastModified = rs!LastModified
                m.bDeleted = False
                m.bLateCondition = rs!LateCondition
                m.bLateAction = rs!LateAction
                m.nSecurityLevel = NullChk(rs!SecurityLevel, 0)
                m.bCannotDelete = rs!CannotDelete
                m.strPassword = DecryptField(rs!Password)
                m.strPrice2RTF = DecryptField(rs!Price2RTF)
                m.strPrice2Coded = DecryptField(rs!Price2Coded)
                m.bReverify = rs!Reverify
                m.bExitOnEntryBar = rs!ExitOnEntryBar
                m.bExitBasedOnEachTrade = rs!ExitBasedOnEachTrade
                m.lNumberContracts = NullChk(rs!NumberContracts, 0&)
                m.bAsPercentOfPosition = rs!AsPercentOfPosition
                If rs.Fields.Count >= 26 Then m.lSystemNumber = NullChk(rs!SystemNumber, 0&)
                If rs.Fields.Count >= 28 Then m.lCategoryID = NullChk(rs!CategoryID, 0&)
            Else
                Err.Raise vbObjectError + 1000, , "Rule cannot be loaded because it is no longer valid"
            End If
        End If
    End With
        
    LoadInputs
    
    ' Load up the Function References...
    Set m.alCondRefs = New cGdArray
    m.alCondRefs.Create eGDARRAY_Longs
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionRules] " & _
        "WHERE [RuleID] = " & Str(m.lRuleID) & ";", dbOpenDynaset)
    Do Until rs.EOF
        m.alCondRefs.Add rs!FunctionIDRef
        rs.MoveNext
    Loop
    
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set rs = Nothing
    RaiseError "cRule.Load", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub Save()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    Dim bNewRec As Boolean
    
    'Validate rule and it's inputs...
    Validate
    ValidateInputs
    
    'Update the table first...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                "WHERE [RuleID]=" & Str(m.lRuleID) & ";", dbOpenDynaset)
    With rs
        If .EOF Then
            .AddNew
            bNewRec = True
        Else
            .Edit
            bNewRec = False
        End If
                
        !Name = m.strName
        EncryptField !CondRTF, m.strCond
        EncryptField !PreviewRTF, RulePreview(m.bBuySell, m.strOrderPlacement, m.strCond, _
                m.strPrice, m.strPrice2RTF)
        m.strCondFillWords = RulePreview(m.bBuySell, m.strOrderPlacement, m.strCond, _
                m.strPrice, m.strPrice2RTF)
        EncryptField !CondCoded, m.strCondCoded
        !ActionCodedName = m.strActionCodedName
        !BuySell = m.bBuySell
        !RuleType = m.iRuleType
        !OrderPlacement = m.strOrderPlacement
        EncryptField !PriceRTF, m.strPrice
        EncryptField !PriceCoded, m.strPriceCoded
        !LibraryID = m.lLibraryID
        !LastModified = m.dLastModified
        !LateCondition = m.bLateCondition
        !LateAction = m.bLateAction
        !SecurityLevel = m.nSecurityLevel
        !CannotDelete = m.bCannotDelete
        EncryptField !Password, m.strPassword
        EncryptField !Price2RTF, m.strPrice2RTF
        EncryptField !Price2Coded, m.strPrice2Coded
        !ExitOnEntryBar = m.bExitOnEntryBar
        !ExitBasedOnEachTrade = m.bExitBasedOnEachTrade
        !NumberContracts = m.lNumberContracts
        !AsPercentOfPosition = m.bAsPercentOfPosition
        If rs.Fields.Count >= 26 Then !SystemNumber = m.lSystemNumber
        If rs.Fields.Count >= 28 Then !CategoryID = m.lCategoryID
        !Reverify = False
        !CheckSum = BuildCheckSum(rs, "tblRules")
                
        .Update
    
        'Note new RuleID and load parameters of rule
        If bNewRec Then
            .MoveLast
            m.lRuleID = !RuleID
        End If
    End With
    
    SaveInputs
    UpdateRefs
    'ReverifyParents
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Select Case Err.Number
        Case 3022
            Err.Raise vbObjectError + 1000, , "This rule already exists.  Use a different name."
        Case Else
            RaiseError "cRule.Save", eGDRaiseError_Raise, g.strAppPath
    End Select

End Sub

'This routines combines the condition and price code text and creates the
'rich text for previewing
Private Function RulePreview(pBuySell As Boolean, pOrderPlacement As String, _
    pCond As String, pPrice1 As String, pPrice2 As String) As String
On Error GoTo ErrSection:
    
    Dim strTemp As String

    'Insert Carriage return (~80001), "THEN", and "BUY/SELL"
    If pBuySell Then
        strTemp = pCond & " ~80001 ~24005THEN ~24016BUY Next Bar at "
    Else
        strTemp = pCond & " ~80001 ~24005THEN ~24017SELL Next Bar at "
    End If
    
    'Insert Order type...
    'Use token 50 (which doesn't translate to anything).  This will
    'default to normal windows text.
    Select Case UCase(pOrderPlacement)
        Case "MARKET", "MARKET ON CLOSE"
            strTemp = strTemp & "~50" & _
                    Format(Len(pOrderPlacement), "000") & _
                    UCase(pOrderPlacement)
        
        Case "STOP WITH LIMIT", "STOP WITH LIMIT CLOSE ONLY"
            strTemp = strTemp & pPrice1 & "~50007 STOP, " & _
                pPrice2
            If Right(UCase(pOrderPlacement), 10) = "CLOSE ONLY" Then
                strTemp = strTemp & "~50017 LIMIT Close Only"
            Else
                strTemp = strTemp & "~50006 LIMIT"
            End If
        
        Case Else
            strTemp = strTemp & pPrice1 & "~50001 " & "~50" & _
                Format(Len(pOrderPlacement), "000") & _
                UCase(pOrderPlacement)
    End Select
    
    'Return reformatting rule...
    RulePreview = strTemp
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRule.RulePreview", eGDRaiseError_Raise, g.strAppPath
    
End Function

'Set the reverify flag of the systems and mm strategies that use this rule
Private Sub ReverifyParents()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    ' Set on the Reverify flag of the systems using this rule
    Set rs = g.dbNav.OpenRecordset("SELECT tblSystems.*, tblSystemRules.* " & _
                "FROM tblSystems INNER JOIN tblSystemRules ON tblSystems.SystemNumber = tblSystemRules.SystemNumber " & _
                "WHERE (((tblSystemRules.RuleID)=" & Str(m.lRuleID) & "));", dbOpenDynaset)
    ValidateCheckSums rs, "tblSystems"
    ValidateCheckSums rs, "tblSystemRules"
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do While Not rs.EOF
        If rs![tblSystems.CheckSum] <> 0.5 And rs![tblSystemRules.CheckSum] <> 0.5 Then
            rs.Edit
            rs!Reverify = True
            rs![tblSystems.CheckSum] = BuildCheckSum(rs, "tblSystems")
            rs.Update
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    
    ' Set on the Reverify flag of the mm strategies using this rule
    g.dbNav.Execute "UPDATE tblMM " & _
        "INNER JOIN tblMMRules " & _
            "ON tblMM.MMid = tblMMRules.MMid " & _
        "SET tblMM.Reverify = True " & _
        "WHERE (((tblMMRules.RuleID)=" & Str(m.lRuleID) & "));"
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    '3078 is "table not found".  This allows the money management app
    'to use this DLL.  The table "tblMM" will not exist in System Nav.
    'and will bomb with 3078 when executing a query for money mgt tables.
    If Err.Number = 3078 Then Resume Next
    Set rs = Nothing
    RaiseError "cRule.ReverifyParents", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub Delete()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
            "WHERE [RuleID]=" & Str(m.lRuleID) & ";", dbOpenDynaset)
    With rs
        If rs.EOF Then
            Err.Raise vbObjectError + 1000, , "Delete was cancelled because Rule does not exist"
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
    RaiseError "cRule.Delete", eGDRaiseError_Raise, g.strAppPath

End Sub

'Load Parameters collection of this rule
Public Sub LoadInputs()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    
    Set m.Inputs = New cInputs
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleParms] " & _
        "WHERE [RuleID]=" & Str(m.lRuleID) & " ORDER BY [RuleID],[ParmSeq];", dbOpenSnapshot)
    Do Until rs.EOF
        m.Inputs.Add m.strName, rs!ParmSeq, rs!ParmName, rs!ParmDesc, rs!ParmID, rs!DefaultValue, _
            False, 0, 0, 0, 0, m.lRuleID, rs!ParmTypeID, rs!DefaultValue, True, rs!FromValue, _
            rs!ToValue, rs!ListID, "", ""
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set rs = Nothing
    RaiseError "cRule.LoadInputs", eGDRaiseError_Raise, g.strAppPath

End Sub

'Validate each rule input...
Private Sub ValidateInputs()
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    
    For lIndex = 1 To m.Inputs.Count
        With m.Inputs.Item(lIndex)
            Select Case .ParmTypeID
                Case C_RetNumericConstant, C_RetNumeric, _
                     C_RetTrueFalseConstant
                    If .Required Then
                        If .FromValue <> 0 And .ToValue <> 0 Then
                            If Val(.DefaultValue) < .FromValue Or _
                               Val(.DefaultValue) > .ToValue Then
                                m.nError = 3
                                Err.Raise vbObjectError + 1000, , _
                                    "The Input Value must be between " & _
                                    Format(.FromValue, "general number") & _
                                    " and " & Format(.ToValue, "general number")
                            End If
                        End If
                    End If
            End Select
        End With
    Next lIndex
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRule.ValidateInputs", eGDRaiseError_Raise, g.strAppPath

End Sub

'This routine loops through the Inputs collection and updates
'to tblRuleParms.  Old inputs are removed.
Private Sub SaveInputs()
On Error GoTo ErrSection:
    
    Dim lIndex As Integer
    Dim bFound As Boolean
    Dim rs As Recordset
    Dim bChanged As Boolean
    
    'No inputs found, exit...
    If m.Inputs Is Nothing Then Exit Sub
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleParms] " & _
        "WHERE [RuleID]=" & Str(m.lRuleID) & ";", dbOpenDynaset)
    
    'Write all new inputs from vsInputs to tblRuleParms
    bChanged = False
    For lIndex = 1 To m.Inputs.Count
        With m.Inputs.Item(lIndex)
            rs.FindFirst "[ParmName]='" & .ParmName & "'"
            If rs.NoMatch Then
                bChanged = True
                rs.AddNew
                rs!RuleID = m.lRuleID
                rs!ParmName = .ParmName
                rs!ParmDesc = .ParmDesc
                rs!ParmTypeID = .ParmTypeID
                rs!ListID = .ListID
                rs!FromValue = .FromValue
                rs!ToValue = .ToValue
                .ParmID = rs!ParmID
            Else
                rs.Edit
                If rs!ParmSeq <> lIndex Then bChanged = True
                If rs!DefaultValue <> Str(.DefaultValue) Then bChanged = True
            End If
            rs!ParmSeq = lIndex
            rs!DefaultValue = Str(.DefaultValue)
            rs.Update
        End With
    Next lIndex
    
    'Delete old inputs from table
    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            bFound = False
            For lIndex = 1 To m.Inputs.Count
                If rs!ParmName = m.Inputs.Item(lIndex).ParmName Then
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
    
    ' Add Market1 if it does not exist
    rs.FindFirst "[ParmName] = 'Market1'"
    If rs.NoMatch Then
        bChanged = True
        rs.AddNew
        rs!RuleID = m.lRuleID
        rs!ParmName = "Market1"
        rs!ParmDesc = ""
        rs!ParmSeq = 1
        rs!ParmTypeID = 5
        rs!ListID = 0
        rs!DefaultValue = "Market1"
        rs!FromValue = 0
        rs!ToValue = 0
        rs.Update
    End If
    
    ' Only mark the parents for reverify if the inputs have changed
    If bChanged Then ReverifyParents
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set rs = Nothing
    RaiseError "cRule.SaveInputs", eGDRaiseError_Raise, g.strAppPath

End Sub

'Remove item from collection and delete table entry...
Public Sub RemoveInput(lParmID As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleParms] " & _
               "WHERE [ParmID]=" & Str(lParmID) & ";", dbOpenDynaset)
    
    If rs.EOF Then
        Err.Raise vbObjectError + 1000, , "Delete cancelled because Input does not exist"
    Else
        rs.Delete
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set rs = Nothing
    RaiseError "cRule.RemoveInput", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub UpdateRefs()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    Dim lIndex As Integer
    
    Set m.alFunctionIDs = New cGdArray
    m.alFunctionIDs.Create eGDARRAY_Longs
    BuildRefs m.alCondRefs
    BuildRefs m.alPrice1Refs
    BuildRefs m.alPrice2Refs
    
    'If the verify button is not pressed, there will be no funtionid
    'lists.  Exit and keep the existing ID's...
    If m.alFunctionIDs.Size = 0 Then Exit Sub
    
    'Delete existing References...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionRules] " & _
        "WHERE [RuleID] = " & Str(m.lRuleID) & ";", dbOpenDynaset)
    Do Until rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
    
    'Add FunctionID's of functions used in condition, price, and/or
    'action expression to tblFunctionRules...
    For lIndex = 0 To m.alFunctionIDs.Size - 1
        rs.AddNew
        rs!RuleID = m.lRuleID
        rs!FunctionIDRef = m.alFunctionIDs(lIndex)
        rs.Update
    Next lIndex
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set rs = Nothing
    RaiseError "cRule.UpdateRefs", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub BuildRefs(pFIDs As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim bFound As Boolean
    
    If pFIDs Is Nothing Then Exit Sub
    For lIndex = 0 To pFIDs.Size - 1
        bFound = False
        For lIndex2 = 0 To m.alFunctionIDs.Size - 1
            If m.alFunctionIDs(lIndex2) = pFIDs(lIndex) Then
                bFound = True
                Exit For
            End If
        Next lIndex2
        If Not bFound Then
            m.alFunctionIDs.Add pFIDs(lIndex)
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRule.BuildRefs", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Function ValidName(ByVal strRuleName As String) As String
On Error GoTo ErrSection:

    Dim strErrorMsg As String
    
    'Rule name
    If (Len(strRuleName) <= 0 Or Len(strRuleName) > 50) Then
        ValidName = "Rule name must be 1 to 50 characters in length and " & _
            "consist of letters, numbers, and blanks."
    
    'First character must be alphabetic
    ElseIf Not FirstCharValid(strRuleName, strErrorMsg) Then
        ValidName = strErrorMsg
    
    'Single quote's not allowed in rule names
    ElseIf InStr(1, strRuleName, "'") > 0 Then
        ValidName = "Rule name cannot contain single quotes (')"
    
    ' Otherwise the Rule Name is valid...
    Else
        ValidName = ""
    
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRule.ValidName", eGDRaiseError_Raise, g.strAppPath
    
End Function

Public Function LoadWithSystemInfo(ByVal lRuleID As Long) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] WHERE [RuleID]=" & Str(lRuleID) & ";", dbOpenDynaset)
    If Not rs.EOF Then
        If rs!CheckSum = BuildCheckSum(rs, "tblRules") Then
            m.lRuleID = lRuleID
            m.strName = rs!Name
            m.strCond = DecryptField(rs!CondRTF)
            m.strCondFillWords = DecryptField(rs!PreviewRTF)
            m.strCondCoded = DecryptField(rs!CondCoded)
            m.strActionCodedName = rs!ActionCodedName
            m.bBuySell = rs!BuySell
            m.iRuleType = rs!RuleType
            m.strOrderPlacement = rs!OrderPlacement
            m.strPrice = DecryptField(rs!PriceRTF)
            m.strPriceCoded = DecryptField(rs!PriceCoded)
            m.lLibraryID = rs!LibraryID
            m.dLastModified = rs!LastModified
            m.bDeleted = False
            m.bLateCondition = rs!LateCondition
            m.bLateAction = rs!LateAction
            m.nSecurityLevel = NullChk(rs!SecurityLevel, 0)
            m.bCannotDelete = rs!CannotDelete
            m.strPassword = DecryptField(rs!Password)
            m.strPrice2RTF = DecryptField(rs!Price2RTF)
            m.strPrice2Coded = DecryptField(rs!Price2Coded)
            m.bReverify = rs!Reverify
            m.bExitOnEntryBar = rs!ExitOnEntryBar
            m.bExitBasedOnEachTrade = rs!ExitBasedOnEachTrade
            m.lNumberContracts = NullChk(rs!NumberContracts, 0&)
            m.bAsPercentOfPosition = rs!AsPercentOfPosition
            If rs.Fields.Count >= 26 Then m.lSystemNumber = rs!SystemNumber
            If rs.Fields.Count >= 28 Then m.lCategoryID = NullChk(rs!CategoryID, 0&)
            
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemRules] WHERE [RuleID]=" & Str(lRuleID) & ";", dbOpenDynaset)
            If Not rs.EOF Then
                m.iSeq = rs!Seq
                m.bSelected = rs!Selected
                m.bAlternate = rs!Alternate
                m.iRuleUse = rs!RuleUse
                m.dLastModKnown = rs!LastModifiedKnown
                m.lUnitsID = rs!UnitsID
                m.strLinkedRules = NullChk(rs!LinkedRules)
                m.bSysExitBasedOnEachTrade = rs!ExitBasedOnEachTrade
                m.lSysNumContracts = rs!NumberContracts
                m.bSysAsPercentOfPosition = rs!AsPercentOfPosition
            Else
                m.iSeq = 0
                m.bSelected = True
                m.bAlternate = False
                m.iRuleUse = m.iRuleType
                m.dLastModKnown = m.dLastModified
                m.strLinkedRules = ""
                m.bSysExitBasedOnEachTrade = m.bExitBasedOnEachTrade
                m.lSysNumContracts = m.lNumberContracts
                m.bSysAsPercentOfPosition = m.bAsPercentOfPosition
            End If
            
            Set m.Inputs = New cInputs
            m.Inputs.LoadRuleParms m.lRuleID
            
            ' Load up the Function References...
            Set m.alCondRefs = New cGdArray
            m.alCondRefs.Create eGDARRAY_Longs
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionRules] " & _
                "WHERE [RuleID] = " & Str(m.lRuleID) & ";", dbOpenDynaset)
            Do Until rs.EOF
                m.alCondRefs.Add rs!FunctionIDRef
                rs.MoveNext
            Loop
            
            LoadWithSystemInfo = True
        Else
            Err.Raise vbObjectError + 1000, , "Rule cannot be loaded because it is no longer valid"
        End If
    End If

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "cRule.LoadWithSystemInfo", eGDRaiseError_Raise, g.strAppPath
    
End Function

Public Sub SaveWithSystemInfo(Optional ByVal lNewSystemID& = 0&)
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim lNewRuleID As Long
    
    If lNewSystemID <> 0& Then m.lSystemNumber = lNewSystemID
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                "WHERE [RuleID]=" & Str(m.lRuleID) & ";", dbOpenDynaset)
    If rs.EOF And rs.BOF Then
        rs.AddNew
        m.lRuleID = rs!RuleID
        lNewRuleID = rs!RuleID
    Else
        rs.Edit
        lNewRuleID = 0&
    End If
    
    rs!Name = m.strName
    EncryptField rs!CondRTF, m.strCond
    m.strCondFillWords = RulePreview(m.bBuySell, m.strOrderPlacement, m.strCond, _
            m.strPrice, m.strPrice2RTF)
    EncryptField rs!PreviewRTF, m.strCondFillWords
    EncryptField rs!CondCoded, m.strCondCoded
    rs!ActionCodedName = m.strActionCodedName
    rs!BuySell = m.bBuySell
    rs!RuleType = m.iRuleType
    rs!OrderPlacement = m.strOrderPlacement
    EncryptField rs!PriceRTF, m.strPrice
    EncryptField rs!PriceCoded, m.strPriceCoded
    rs!LibraryID = m.lLibraryID
    rs!LastModified = m.dLastModified
    rs!LateCondition = m.bLateCondition
    rs!LateAction = m.bLateAction
    rs!SecurityLevel = m.nSecurityLevel
    rs!CannotDelete = m.bCannotDelete
    EncryptField rs!Password, m.strPassword
    EncryptField rs!Price2RTF, m.strPrice2RTF
    EncryptField rs!Price2Coded, m.strPrice2Coded
    rs!Reverify = m.bReverify
    rs!ExitOnEntryBar = m.bExitOnEntryBar
    rs!ExitBasedOnEachTrade = m.bExitBasedOnEachTrade
    rs!NumberContracts = m.lNumberContracts
    rs!AsPercentOfPosition = m.bAsPercentOfPosition
    If rs.Fields.Count >= 26 Then rs!SystemNumber = m.lSystemNumber
    If rs.Fields.Count >= 28 Then rs!CategoryID = m.lCategoryID
    rs!CheckSum = BuildCheckSum(rs, "tblRules")
    rs.Update
    
    If m.lSystemNumber <> 0 Then
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemRules] " & _
                "WHERE [SystemNumber]=" & Str(m.lSystemNumber) & " AND [RuleID]=" & Str(m.lRuleID) & ";", dbOpenDynaset)
        If rs.EOF And rs.BOF Then
            rs.AddNew
            rs!SystemNumber = m.lSystemNumber
            rs!RuleID = m.lRuleID
        Else
            rs.Edit
        End If
        
        rs!Seq = m.iSeq '0&
        rs!Selected = m.bSelected
        rs!Alternate = m.bAlternate
        rs!RuleUse = m.iRuleUse
        rs!LastModifiedKnown = m.dLastModKnown
        rs!UnitsID = m.lUnitsID
        rs!LinkedRules = m.strLinkedRules
        rs!ExitBasedOnEachTrade = m.bSysExitBasedOnEachTrade
        rs!NumberContracts = m.lSysNumContracts
        rs!AsPercentOfPosition = m.bSysAsPercentOfPosition
        rs!CheckSum = BuildCheckSum(rs, "tblSystemRules")
        rs.Update
    End If
    
    If Not m.Inputs Is Nothing Then m.Inputs.SaveRuleParms lNewRuleID, lNewSystemID
    UpdateRefs

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cRule.SaveWithSystemInfo", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Friend Sub SetPrivateData(mCopyFrom As mPrivate, Optional ByVal lNewRuleID As Long = 0&, _
            Optional ByVal lNewSystemID As Long = 0&)
    
    ' get copy of all private data
    m = mCopyFrom
    m.lRuleID = lNewRuleID
    m.lSystemNumber = lNewSystemID

    Set m.alFunctionIDs = mCopyFrom.alFunctionIDs.MakeCopy
    Set m.alCondRefs = mCopyFrom.alCondRefs.MakeCopy
    Set m.alPrice1Refs = mCopyFrom.alPrice1Refs.MakeCopy
    Set m.alPrice2Refs = mCopyFrom.alPrice2Refs.MakeCopy
    Set m.Inputs = mCopyFrom.Inputs.MakeCopy(lNewRuleID, lNewSystemID)

End Sub

Public Function MakeCopy(Optional ByVal lNewRuleID As Long = 0&, _
            Optional ByVal lNewSystemID As Long = 0&) As cRule
    
    Dim aCopy As New cRule
    
    aCopy.SetPrivateData m, lNewRuleID, lNewSystemID
    Set MakeCopy = aCopy

End Function

Private Sub Class_Initialize()

    Set m.alFunctionIDs = New cGdArray
    Set m.alCondRefs = New cGdArray
    Set m.alPrice1Refs = New cGdArray
    Set m.alPrice2Refs = New cGdArray
    Set m.Inputs = New cInputs

End Sub

Private Sub Class_Terminate()

    Set m.alFunctionIDs = Nothing
    Set m.alCondRefs = Nothing
    Set m.alPrice1Refs = Nothing
    Set m.alPrice2Refs = Nothing
    Set m.Inputs = Nothing

End Sub

Public Sub CopyRuleInfo(ByVal Rule As cRule)
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim bFound As Boolean
    Dim lNewID As Long
    
    m.strCond = Rule.Cond
    m.strCondFillWords = Rule.CondFillWords
    m.strCondCoded = Rule.CondCoded
    m.strActionCodedName = Rule.ActionCodedName
    m.bBuySell = Rule.BuySell
    m.iRuleType = Rule.RuleType
    m.strOrderPlacement = Rule.OrderPlacement
    m.strPrice = Rule.Price
    m.strPriceCoded = Rule.PriceCoded
    m.dLastModified = Rule.LastModified
    m.bDeleted = Rule.Deleted
    m.bLateCondition = Rule.LateCondition
    m.bLateAction = Rule.LateAction
    m.strPrice2RTF = Rule.Price2RTF
    m.strPrice2Coded = Rule.Price2Coded
    m.bReverify = Rule.Reverify
    m.bExitOnEntryBar = Rule.ExitOnEntryBar
    m.bExitBasedOnEachTrade = Rule.ExitBasedOnEachTrade
    m.lNumberContracts = Rule.NumberContracts
    m.bAsPercentOfPosition = Rule.AsPercentOfPosition
    m.lCategoryID = Rule.CategoryID
    Set m.alCondRefs = New cGdArray
    m.alCondRefs.CopyFromHandle Rule.CondRefs
    
    For lIndex = 1 To Rule.Inputs.Count
        bFound = False
        For lIndex2 = 1 To m.Inputs.Count
            If m.Inputs.Item(lIndex2).ParmName = Rule.Inputs.Item(lIndex).ParmName Then
                bFound = True
                Exit For
            End If
        Next lIndex2
        
        If bFound Then
            m.Inputs.Item(lIndex).CopyRuleInfo Rule.Inputs.Item(lIndex)
        Else
            lNewID = m.Inputs.NextParmID
            m.Inputs.AddRuleParm Rule.Inputs.Item(lIndex).MakeCopy(lNewID, m.lRuleID, m.lSystemNumber), lNewID
        End If
    Next lIndex
    
    For lIndex = 1 To m.Inputs.Count
        bFound = False
        For lIndex2 = 1 To Rule.Inputs.Count
            If Rule.Inputs.Item(lIndex2).ParmName = m.Inputs.Item(lIndex).ParmName Then
                bFound = True
                Exit For
            End If
        Next lIndex2
        
        If Not bFound Then m.Inputs.Remove Str(m.Inputs.Item(lIndex).ParmID)
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRule.CopyRuleInfo", eGDRaiseError_Raise, g.strAppPath
    
End Sub
