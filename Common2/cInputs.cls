VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private Type mPrivate
    Inputs As cGdTree
    lRuleID As Long
End Type
Private m As mPrivate

'Adds a Input to the collection
Public Sub Add(pRuleName As Variant, pParmSeq As Variant, _
        pParmName As Variant, pParmDesc As Variant, pParmID As Variant, _
        pValue As Variant, pIfOptimize As Variant, _
        pOptFromValue As Variant, _
        pOptToValue As Variant, pOptStepValue As Variant, _
        pOptListID As Variant, _
        pRuleID As Variant, pParmTypeID As Variant, _
        pDefaultValue As Variant, pRequired As Variant, pFromValue As Variant, pToValue As Variant, _
        pListID As Variant, pFillPre As Variant, pFillPost As Variant)
On Error GoTo ErrSection:
    
    Dim TmpInput As cInput
    
    Set TmpInput = New cInput
    With TmpInput
        .RuleName = pRuleName
        .ParmSeq = pParmSeq
        .ParmName = pParmName
        .ParmID = pParmID
        .Value = pValue
        .IfOptimize = pIfOptimize
        .OptFromValue = pOptFromValue
        .OptToValue = pOptToValue
        .OptStepValue = pOptStepValue
        .OptListID = pOptListID
        .RuleID = pRuleID
        .ParmDesc = NullChk(pParmDesc)
        .ParmTypeID = pParmTypeID
        .DefaultValue = NullChk(pDefaultValue)
        .Required = pRequired
        .ListID = pListID
        .FillPre = NullChk(pFillPre)
    
        'MT 6/2001: TradeSense compatibility =======================
        'Borrow out-of-date fields to allow calls from TradeSense to store
        'Usage, ParmTypeName, and list name in inputs collection
        '.FromValue = pFromValue
        '.ToValue = pToValue
        'If IsNull(pFillPost) Then .FillPost = "" Else .FillPost = pFillPost
        If Not IsDate(pFromValue) And Val(pFromValue) >= 0 And Val(pFromValue) <= 256 Then
            .ParmTypeUsage = Val(pFromValue)
        End If
        If Not IsNumeric(pToValue) Then
            .ParmTypeName = Str(pToValue)
        End If
        If IsNull(pFillPost) Then .ListName = "" Else .ListName = pFillPost
        
    End With
    m.Inputs.Add TmpInput, Str(pParmID)
    
ErrExit:
    Set TmpInput = Nothing
    Exit Sub
    
ErrSection:
    Set TmpInput = Nothing
    RaiseError "cInputs.Add", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Property Get Found(ByVal vKey As Variant) As Boolean
    Found = m.Inputs.Exists(vKey)
End Property

'Returns true if input1,2 are Compatible inputs.
'Sets pNewInputTypeID to new type based on the inputs.  For example, series
'of numbers and constant number defaults to constant.
Public Function Compare(pInputTypeID1 As Long, pInputTypeID2 As Long, _
    pNewInputTypeID As Long) As Boolean
On Error GoTo ErrSection:

    Select Case pInputTypeID1
        Case C_RetNumericConstant
        
            Select Case pInputTypeID2
                Case C_RetNumericConstant
                    Compare = True: pNewInputTypeID = C_RetNumericConstant
                Case C_RetTrueFalseConstant
                    Compare = True: pNewInputTypeID = C_RetNumericConstant
                Case C_RetText
                    Compare = False: pNewInputTypeID = 0
                Case C_RetNumeric
                    Compare = True: pNewInputTypeID = C_RetNumericConstant
                Case C_RetTrueFalse
                    Compare = True: pNewInputTypeID = C_RetNumericConstant
                Case Else   'Series of text
                    Compare = False: pNewInputTypeID = 0
            End Select
            
        Case C_RetTrueFalseConstant
        
            Select Case pInputTypeID2
                Case C_RetNumericConstant
                    Compare = True: pNewInputTypeID = C_RetNumericConstant
                Case C_RetTrueFalseConstant
                    Compare = True: pNewInputTypeID = C_RetTrueFalseConstant
                Case C_RetText
                    Compare = False: pNewInputTypeID = 0
                Case C_RetNumeric
                    Compare = True: pNewInputTypeID = C_RetTrueFalseConstant
                Case C_RetTrueFalse
                    Compare = True: pNewInputTypeID = C_RetTrueFalseConstant
                Case Else   'series of text
                    Compare = False: pNewInputTypeID = 0
            End Select
            
        Case C_RetText
        
            Select Case pInputTypeID2
                Case C_RetNumericConstant
                    Compare = False: pNewInputTypeID = 0
                Case C_RetTrueFalseConstant
                    Compare = False: pNewInputTypeID = 0
                Case C_RetText
                    Compare = True: pNewInputTypeID = C_RetText
                Case C_RetNumeric
                    Compare = False: pNewInputTypeID = 0
                Case C_RetTrueFalse
                    Compare = False: pNewInputTypeID = 0
                Case Else   'Series of text
                    Compare = True: pNewInputTypeID = C_RetText
            End Select
            
        Case C_RetNumeric
        
            Select Case pInputTypeID2
                Case C_RetNumericConstant
                    Compare = True: pNewInputTypeID = C_RetNumericConstant
                Case C_RetTrueFalseConstant
                    Compare = True: pNewInputTypeID = C_RetTrueFalseConstant
                Case C_RetText
                    Compare = False: pNewInputTypeID = 0
                Case C_RetNumeric
                    Compare = True: pNewInputTypeID = C_RetNumeric
                Case C_RetTrueFalse
                    Compare = True: pNewInputTypeID = C_RetNumeric
                Case Else   'Series of text
                    Compare = False: pNewInputTypeID = 0
            End Select
            
        Case C_RetTrueFalse
        
            Select Case pInputTypeID2
                Case C_RetNumericConstant
                    Compare = True: pNewInputTypeID = C_RetNumericConstant
                Case C_RetTrueFalseConstant
                    Compare = True: pNewInputTypeID = C_RetTrueFalseConstant
                Case C_RetText
                    Compare = False: pNewInputTypeID = 0
                Case C_RetNumeric
                    Compare = True: pNewInputTypeID = C_RetNumeric
                Case C_RetTrueFalse
                    Compare = True: pNewInputTypeID = C_RetNumeric
                Case Else       'Series of text
                    Compare = False: pNewInputTypeID = 0
            End Select
            
       Case Else   'Series of text
       
            Select Case pInputTypeID2
                Case C_RetNumericConstant
                    Compare = False: pNewInputTypeID = 0
                Case C_RetTrueFalseConstant
                    Compare = False: pNewInputTypeID = 0
                Case C_RetText
                    Compare = True: pNewInputTypeID = C_RetText
                Case C_RetNumeric
                    Compare = False: pNewInputTypeID = 0
                Case C_RetTrueFalse
                    Compare = False: pNewInputTypeID = 0
                Case Else   'Series of text
                    Compare = True: pNewInputTypeID = C_TEXT
            End Select
            
    End Select
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cInputs.Compare", eGDRaiseError_Raise, g.strAppPath

End Function

Property Get Item(ByVal vKey As Variant) As cInput
    Set Item = m.Inputs(vKey)
End Property

Property Let Item(ByVal vKey As Variant, ByVal Parm As cInput)
    Set m.Inputs(vKey) = Parm
    If m.Inputs.Key(vKey) <> Str(Parm.ParmID) Then
        m.Inputs.Key(vKey) = Str(Parm.ParmID)
    End If
End Property

Property Get Count() As Long
    Count = m.Inputs.Count
End Property

Public Sub Remove(ByVal vKey As Variant)
    m.Inputs.Remove vKey
End Sub

Private Sub Class_Initialize()
    Set m.Inputs = New cGdTree
End Sub

Private Sub Class_Terminate()
    Set m.Inputs = Nothing
End Sub

Public Sub LoadRuleParms(ByVal lRuleID As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim Parm As cInput
    
    m.lRuleID = lRuleID
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleParms] " & _
                "WHERE [RuleID]=" & Str(lRuleID) & " " & _
                "ORDER BY [ParmSeq];", dbOpenDynaset)
    Do While Not rs.EOF
        Set Parm = New cInput
        If Parm.LoadRuleParm(rs!ParmID) Then m.Inputs.Add Parm, Str(rs!ParmID)
        
        rs.MoveNext
    Loop

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cInputs.LoadRuleParms", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Sub SaveRuleParms(Optional ByVal lNewRuleID& = 0&, Optional ByVal lNewSystemID& = 0&)
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim rs As Recordset
    Dim lOldParmID As Long
    
    If lNewRuleID <> 0& Then m.lRuleID = lNewRuleID
    
    ' Save the existing parms...
    For lIndex = 1 To Count
        Item(lIndex).SaveRuleParm lNewRuleID, lNewSystemID
        If Str(Item(lIndex).ParmID) <> m.Inputs.Key(lIndex) Then
            m.Inputs.Key(lIndex) = Str(Item(lIndex).ParmID)
        End If
    Next lIndex
    
    ' Delete any parms in the database that no longer exist...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleParms] " & _
                "WHERE [RuleID]=" & Str(m.lRuleID) & " " & _
                "ORDER BY [ParmID];", dbOpenDynaset)
    Do While Not rs.EOF
        If Not Found(Str(rs!ParmID)) Then
            rs.Delete
        End If

        rs.MoveNext
    Loop

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cInputs.SaveRuleParms", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Sub AddRuleParm(Parm As cInput, ByVal lParmID As Long)
    m.Inputs.Add Parm, Str(lParmID)
End Sub

Friend Sub SetPrivateData(mCopyFrom As mPrivate, Optional ByVal lNewRuleID As Long = 0&, _
            Optional ByVal lNewSystemID As Long = 0&)
    
    Dim lIndex As Long
    Dim lNewParmID As Long
    
    ' get copy of all private data
    m = mCopyFrom
    m.lRuleID = lNewRuleID
    
    Set m.Inputs = New cGdTree
    For lIndex = 1 To mCopyFrom.Inputs.Count
        If lNewRuleID < 0 Then
            lNewParmID = (lNewRuleID * 100) - lIndex
        Else
            lNewParmID = (lNewRuleID * 100) + lIndex
        End If
        m.Inputs.Add mCopyFrom.Inputs.Item(lIndex).MakeCopy(lNewParmID, lNewRuleID, lNewSystemID), Str(lNewParmID)
    Next lIndex

End Sub

Public Function MakeCopy(Optional ByVal lNewRuleID As Long = 0&, _
            Optional ByVal lNewSystemID As Long = 0&) As cInputs
    
    Dim aCopy As New cInputs
    
    aCopy.SetPrivateData m, lNewRuleID, lNewSystemID
    Set MakeCopy = aCopy

End Function

Public Function NextParmID() As Long
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim lMinInputID As Long
    
    lMinInputID = 0
    For lIndex = 1 To m.Inputs.Count
        If m.Inputs(lIndex).ParmID < lMinInputID Then lMinInputID = m.Inputs(lIndex).ParmID
    Next lIndex
    
    If lMinInputID = 0 Then
        NextParmID = Abs(m.lRuleID * 100 + 1) * -1
    Else
        NextParmID = lMinInputID - 1
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cInputs.NextParmID", eGDRaiseError_Raise, g.strAppPath
    
End Function
