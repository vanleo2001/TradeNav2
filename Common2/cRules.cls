VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private Type mPrivate
    Rules As cGdTree
    
    lSystemID As Long
End Type
Private m As mPrivate

'Adds a rule to the collection
Public Sub Add(ByVal lRuleID As Long, Rule As cRule)
    m.Rules.Add Rule, Str(lRuleID)
End Sub

Public Sub Load()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    Dim Rule As cRule
    
    Set m.Rules = New cGdTree
    
    'Update the table first...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] ORDER BY [Name];", dbOpenDynaset)
    ValidateCheckSums rs, "tblRules"
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do Until rs.EOF
        If rs!CheckSum <> 0.5 Then
            Set Rule = New cRule
            With Rule
                .RuleID = rs!RuleID
                .Name = rs!Name
                .Cond = DecryptField(rs!CondRTF)
                .CondFillWords = DecryptField(rs!PreviewRTF)
                .CondCoded = DecryptField(rs!CondCoded)
                .ActionCodedName = rs!ActionCodedName
                .BuySell = rs!BuySell
                .RuleType = rs!RuleType
                .OrderPlacement = rs!OrderPlacement
                .Price = DecryptField(rs!PriceRTF)
                .PriceCoded = DecryptField(rs!PriceCoded)
                .LibraryID = rs!LibraryID
                .LastModified = rs!LastModified
                .Deleted = False
                .LateCondition = rs!LateCondition
                .LateAction = rs!LateAction
                .SecurityLevel = NullChk(rs!SecurityLevel, 0)
                .CannotDelete = rs!CannotDelete
                .Password = DecryptField(rs!Password)
                .Price2RTF = DecryptField(rs!Price2RTF)
                .Price2Coded = DecryptField(rs!Price2Coded)
                .ExitOnEntryBar = rs!ExitOnEntryBar
                .ExitBasedOnEachTrade = rs!ExitBasedOnEachTrade
                .NumberContracts = NullChk(rs!NumberContracts, 0&)
                .AsPercentOfPosition = rs!AsPercentOfPosition
                If rs.Fields.Count >= 26 Then .SystemNumber = rs!SystemNumber
                If rs.Fields.Count >= 28 Then .CategoryID = NullChk(rs!CategoryID, 0&)
                .Reverify = rs!Reverify
            End With
            Add rs!RuleID, Rule
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    
ErrExit:
    Set Rule = Nothing
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set Rule = Nothing
    Set rs = Nothing
    RaiseError "cRules.Load", eGDRaiseError_Raise, g.strAppPath

End Sub

Property Get Found(ByVal vKey As Variant) As Boolean
    Found = m.Rules.Exists(vKey)
End Property

Property Get Item(ByVal vKey As Variant) As cRule
    Set Item = m.Rules(vKey)
End Property

Property Let Item(ByVal vKey As Variant, ByVal Rule As cRule)
    Set m.Rules(vKey) = Rule
    If m.Rules.Key(vKey) = Str(Rule.RuleID) Then
        m.Rules.Key(vKey) = Str(Rule.RuleID)
    End If
End Property

Property Get Count() As Long
    Count = m.Rules.Count
End Property

Public Sub Remove(ByVal vKey As Variant)
    m.Rules.Remove vKey
End Sub

Private Sub Class_Initialize()
    Set m.Rules = New cGdTree
End Sub

Private Sub Class_Terminate()
    Set m.Rules = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NameFound
'' Description: Is the given name in the collection?
'' Inputs:      Name to search for
'' Returns:     Index where found or zero if not found
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NameFound(ByVal strName As String) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    
    NameFound = 0&
    For lIndex = 1 To Count
        If Item(lIndex).Name = strName Then
            NameFound = lIndex
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRules.NameFound", eGDRaiseError_Raise, g.strAppPath
    
End Function

Public Sub LoadSystemRules(ByVal lSystemID As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim Rule As cRule
    
    m.lSystemID = lSystemID
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
            "WHERE [SystemNumber]=" & Str(lSystemID) & " " & _
            "ORDER BY [RuleID];", dbOpenDynaset)
    Do While Not rs.EOF
        Set Rule = New cRule
        If Rule.LoadWithSystemInfo(rs!RuleID) Then
            m.Rules.Add Rule, Str(rs!RuleID)
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Set rs = Nothing
    Set Rule = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    Set Rule = Nothing
    RaiseError "cRules.LoadSystemRules", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Sub SaveSystemRules(Optional ByVal lNewSystemID& = 0&)
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim rs As Recordset
    Dim lOldRuleID As Long
    Dim astrLinkedRules() As String
    Dim lRuleID As Long
    Dim lIndex2 As Long
    Dim bFixLink As Boolean

    If lNewSystemID <> 0& Then m.lSystemID = lNewSystemID
    
    ' If there are any exits linked to entries that have not been saved yet(the
    ' Rule ID is negative), save the position in the collection of that entry so
    ' that we can go through and fix them after the Rule is assigned a real ID
    bFixLink = False
    For lIndex = 1 To Count
        astrLinkedRules = Split(Item(lIndex).LinkedRules, ",")
        For lIndex2 = 0 To UBound(astrLinkedRules)
            lRuleID = CLng(ValOfText(astrLinkedRules(lIndex2)))
            If lRuleID < 0 Then
                astrLinkedRules(lIndex2) = "*" & Str(m.Rules.Index(Str(lRuleID)))
            End If
        Next lIndex2
        If UBound(astrLinkedRules) >= 0 Then
            Item(lIndex).LinkedRules = Join(astrLinkedRules, ",")
            bFixLink = True
        End If
    Next lIndex

    ' Save the Rules in the collection...
    For lIndex = 1 To Count
        Item(lIndex).SaveWithSystemInfo lNewSystemID
        If m.Rules.Key(lIndex) <> Str(Item(lIndex).RuleID) Then
            m.Rules.Key(lIndex) = Str(Item(lIndex).RuleID)
        End If
    Next lIndex
    
    ' If we have to fix any linked rules with the real Rule ID, then fix it
    ' and resave it
    If bFixLink Then
        For lIndex = 1 To Count
            astrLinkedRules = Split(Item(lIndex).LinkedRules, ",")
            For lIndex2 = 0 To UBound(astrLinkedRules)
                If Left(astrLinkedRules(lIndex2), 1) = "*" Then
                    astrLinkedRules(lIndex2) = Str(Item(CLng(Mid(astrLinkedRules(lIndex2), 2))).RuleID)
                End If
            Next lIndex2
            If UBound(astrLinkedRules) >= 0 Then
                Item(lIndex).LinkedRules = Join(astrLinkedRules, ",")
                Item(lIndex).SaveWithSystemInfo
            End If
        Next lIndex
    End If
    
    ' Remove any rules in the database that are no longer in the collection...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
            "WHERE [SystemNumber]=" & Str(m.lSystemID) & " " & _
            "ORDER BY [RuleID];", dbOpenDynaset)
    Do While Not rs.EOF
        If Not Found(Str(rs!RuleID)) Then
            rs.Delete
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cRules.SaveSystemRules", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Friend Sub SetPrivateData(mCopyFrom As mPrivate, Optional ByVal lNewSystemID As Long = 0&)
    
    Dim lIndex As Long
    Dim lNewRuleID As Long
    Dim lOldRuleID As Long
    Dim astrLinkedRules() As String
    Dim lRuleID As Long
    Dim lIndex2 As Long
    Dim bFixLink As Boolean
    Dim astrRules As New cGdArray
    
    ' get copy of all private data
    m = mCopyFrom
    m.lSystemID = lNewSystemID
    
    ' If there are any exits linked to entries, save the position in the collection
    ' of that entry so that we can go through and fix them after the Rule is assigned
    ' a new ID
    astrRules.Create eGDARRAY_Strings
    bFixLink = False
    For lIndex = 1 To Count
        astrLinkedRules = Split(mCopyFrom.Rules.Item(lIndex).LinkedRules, ",")
        For lIndex2 = 0 To UBound(astrLinkedRules)
            lRuleID = CLng(ValOfText(astrLinkedRules(lIndex2)))
            If lRuleID <> 0& Then astrLinkedRules(lIndex2) = "*" & Str(mCopyFrom.Rules.Index(Str(lRuleID)))
        Next lIndex2
        If UBound(astrLinkedRules) >= 0 Then
            astrRules(lIndex) = Join(astrLinkedRules, ",")
            bFixLink = True
        End If
    Next lIndex

    Set m.Rules = New cGdTree
    For lIndex = 1 To mCopyFrom.Rules.Count
        If lNewSystemID < 0 Then
            lNewRuleID = (lNewSystemID * 100) - lIndex
        Else
            lNewRuleID = (lNewSystemID * 100) + lIndex
        End If
        m.Rules.Add mCopyFrom.Rules.Item(lIndex).MakeCopy(lNewRuleID, lNewSystemID), Str(lNewRuleID)
    Next lIndex

    ' If we have to fix any linked rules with the real Rule ID, then fix it
    ' and resave it
    If bFixLink Then
        For lIndex = 1 To Count
            astrLinkedRules = Split(astrRules(lIndex), ",")
            For lIndex2 = 0 To UBound(astrLinkedRules)
                If Left(astrLinkedRules(lIndex2), 1) = "*" Then
                    astrLinkedRules(lIndex2) = Str(Item(CLng(Mid(astrLinkedRules(lIndex2), 2))).RuleID)
                End If
            Next lIndex2
            If UBound(astrLinkedRules) >= 0 Then
                Item(lIndex).LinkedRules = Join(astrLinkedRules, ",")
            End If
        Next lIndex
    End If
    
    Set astrRules = Nothing

End Sub

Public Function MakeCopy(Optional ByVal lNewSystemID As Long = 0&) As cRules
    
    Dim aCopy As New cRules
    
    aCopy.SetPrivateData m, lNewSystemID
    Set MakeCopy = aCopy

End Function

Public Function NextRuleID() As Long
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim lMinRuleID As Long
    
    lMinRuleID = 0
    For lIndex = 1 To m.Rules.Count
        If m.Rules(lIndex).RuleID < lMinRuleID Then lMinRuleID = m.Rules(lIndex).RuleID
    Next lIndex
    
    If lMinRuleID = 0 Then
        NextRuleID = Abs(m.lSystemID * 100 + 1) * -1
    Else
        NextRuleID = lMinRuleID - 1
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRules.NextRuleID", eGDRaiseError_Raise, g.strAppPath
    
End Function

