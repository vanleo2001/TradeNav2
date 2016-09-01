VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cInput.cls
'' Description: Wrapper class for an input
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/23/2010   DAJ         Make sure to do a Str around default value and value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    'The following properties are used by the system and money management objects
    lSystemID As Long
    lRuleID As Long
    strRuleName As String
    vValue As Variant
    bIfOptimize As Boolean
    dOptFromValue As Double
    dOptToValue As Double
    dOptStepValue As Double
    lOptListID As Long
    
    'The following are only used by the functionmgr
    strFillPre As String
    strFillPost As String
    
    ' The follwing are only for Securities
    strPath As String
    strSymbol As String
    strMarketSymbol As String
    strPeriodicity As String
    strFormat As String
    strSecurityType As String
    strSecurityName As String
    strGroupID As String
    lSymbolID As Long
    
    'The following are used by all callers
    lParmSeq As Long
    strParmName As String
    strParmDesc As String
    lParmID As Long
    bRequired As Boolean
    lParmTypeID As Long
    vDefaultValue As Variant
    dFromValue As Double
    dToValue As Double
    lListID As Long
    
    'MT 6/2001: TradeSense compatibility =======================
    bParmTypeUsage As Byte
    strListName As String
    strParmTypeName As String
    
    'Variables used in error processing...
    bError As Byte
End Type
Private m As mPrivate

'MT 6/2001: TradeSense compatibility =======================
Property Get ParmTypeUsage() As Byte
    ParmTypeUsage = m.bParmTypeUsage
End Property
Property Get ParmTypeName() As String
    ParmTypeName = m.strParmTypeName
End Property
Property Get ListName() As String
    ListName = m.strListName
End Property
Property Let ParmTypeUsage(pData As Byte)
    m.bParmTypeUsage = pData
End Property
Property Let ParmTypeName(pData As String)
    m.strParmTypeName = pData
End Property
Property Let ListName(pData As String)
    m.strListName = pData
End Property
'MT 6/2001: TradeSense compatibility =======================


Property Get ErrNbr() As Byte
    ErrNbr = m.bError
End Property

Property Get RuleName() As String
    RuleName = m.strRuleName
End Property
Property Get RuleID() As Long
    RuleID = m.lRuleID
End Property
Property Get ParmSeq() As Long
    ParmSeq = m.lParmSeq
End Property
Property Get ParmName() As String
    ParmName = m.strParmName
End Property
Property Get ParmDesc() As String
    ParmDesc = m.strParmDesc
End Property
Property Get ParmID() As Long
    ParmID = m.lParmID
End Property
Property Get Value() As Variant
    Value = m.vValue
End Property
Property Get Required() As Boolean
    Required = m.bRequired
End Property
Property Get IfOptimize() As Boolean
    IfOptimize = m.bIfOptimize
End Property
Property Get OptFromValue() As Double
    OptFromValue = m.dOptFromValue
End Property
Property Get OptToValue() As Double
    OptToValue = m.dOptToValue
End Property
Property Get OptStepValue() As Double
    OptStepValue = m.dOptStepValue
End Property
Property Get OptListID() As Long
    OptListID = m.lOptListID
End Property
Property Get ParmTypeID() As Long
    ParmTypeID = m.lParmTypeID
End Property
Property Get DefaultValue() As Variant
    DefaultValue = m.vDefaultValue
End Property
Property Get FromValue() As Double
    FromValue = m.dFromValue
End Property
Property Get ToValue() As Double
    ToValue = m.dToValue
End Property
Property Get ListID() As Long
    ListID = m.lListID
End Property
Property Get FillPre() As String
    FillPre = m.strFillPre
End Property
Property Get FillPost() As String
    FillPost = m.strFillPost
End Property
Property Get Path() As String
    Path = m.strPath
End Property
Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Property Get MarketSymbol() As String
    MarketSymbol = m.strMarketSymbol
End Property
Property Get Periodicity() As String
    Periodicity = m.strPeriodicity
End Property
Property Get Format() As String
    Format = m.strFormat
End Property
Property Get SecurityType() As String
    SecurityType = m.strSecurityType
End Property
Property Get SecurityName() As String
    SecurityName = m.strSecurityName
End Property
Property Get GroupID() As String
    GroupID = m.strGroupID
End Property
Property Get SystemID() As Long
    SystemID = m.lSystemID
End Property
Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property

Property Let RuleName(pData As String)
    m.strRuleName = pData
End Property
Property Let RuleID(pData As Long)
    m.lRuleID = pData
End Property
Property Let ParmSeq(pData As Long)
    m.lParmSeq = pData
End Property
Property Let ParmName(pData As String)
    m.strParmName = pData
End Property
Property Let ParmDesc(pData As String)
    m.strParmDesc = pData
End Property
Property Let ParmID(pData As Long)
    m.lParmID = pData
End Property
Property Let Value(pData As Variant)
    m.vValue = pData
End Property
Property Let Required(pData As Boolean)
    m.bRequired = pData
End Property
Property Let IfOptimize(pData As Boolean)
    m.bIfOptimize = pData
End Property
Property Let OptFromValue(pData As Double)
    m.dOptFromValue = pData
End Property
Property Let OptToValue(pData As Double)
    m.dOptToValue = pData
End Property
Property Let OptStepValue(pData As Double)
    m.dOptStepValue = pData
End Property
Property Let OptListID(pData As Long)
    m.lOptListID = pData
End Property
Property Let ParmTypeID(pData As Long)
    m.lParmTypeID = pData
End Property
Property Let DefaultValue(pData As Variant)
    m.vDefaultValue = pData
End Property
Property Let FromValue(pData As Double)
    m.dFromValue = pData
End Property
Property Let ToValue(pData As Double)
    m.dToValue = pData
End Property
Property Let ListID(pData As Long)
    m.lListID = pData
End Property
Property Let FillPre(pData As String)
    m.strFillPre = pData
End Property
Property Let FillPost(pData As String)
    m.strFillPost = pData
End Property
Property Let Path(ByVal pData As String)
    m.strPath = pData
End Property
Property Let Symbol(ByVal pData As String)
    m.strSymbol = pData
End Property
Property Let MarketSymbol(ByVal pData As String)
    m.strMarketSymbol = pData
End Property
Property Let Periodicity(ByVal pData As String)
    m.strPeriodicity = pData
End Property
Property Let Format(ByVal pData As String)
    m.strFormat = pData
End Property
Property Let SecurityType(ByVal pData As String)
    m.strSecurityType = pData
End Property
Property Let SecurityName(ByVal pData As String)
    m.strSecurityName = pData
End Property
Property Let GroupID(ByVal strGroupID As String)
    m.strGroupID = strGroupID
End Property
Property Let SystemID(ByVal pData As Long)
    m.lSystemID = pData
End Property
Property Let SymbolID(ByVal lSymbolID As Long)
    m.lSymbolID = lSymbolID
End Property

Public Sub ValidateFunctionInput()
On Error GoTo ErrSection:

    'Input name...
    If Len(m.strParmName) = 0 Then
        m.bError = 1
        Err.Raise vbObjectError + 1000, , "An Input name must be entered"
    End If
    If Len(m.strParmName) > 50 Then
        m.bError = 1
        Err.Raise vbObjectError + 1000, , "Input name cannot exceed 50 characters"
    End If

    'Make sure delimiter for builtin in functions is not specified in
    'input name
    If InStr(1, m.strParmName, "^") > 0 Then
        m.bError = 1
        Err.Raise vbObjectError + 1000, , "Input name cannot contains the '^' character"
    End If
    
    'Verify that backets are around optional inputs
    'If Not m.bRequired Then
    '    If Left(m.strParmName, 1) <> "[" Then m.strParmName = "[" & m.strParmName
    '    If Right(m.strParmName, 1) <> "]" Then m.strParmName = m.strParmName & "]"
    'End If
    
    'Validate Input value and default value
    Select Case m.lParmTypeID
        Case C_RetNumericConstant, C_RetTrueFalseConstant, C_RetNumeric
            If m.dFromValue <> 0 And m.dToValue <> 0 Then
                If Val(m.vDefaultValue) < m.dFromValue Or Val(m.vDefaultValue) > m.dToValue Then
                    m.bError = 3
                    Err.Raise vbObjectError + 1000, , _
                        "The Default Value must be between " & VBA.Format(m.dFromValue, "general number") & _
                        " and " & VBA.Format(m.dToValue, "general number")
                End If
            End If

        Case C_RetTrades
            If m.strParmName = "Trades" Then m.vDefaultValue = "Trades"
            m.vDefaultValue = m.strParmName
        
        Case C_RetBars
            If m.strParmName = "Market1" Then m.vDefaultValue = "Market1"
            m.vDefaultValue = m.strParmName
        
        Case C_RetText
            'If m.lListID = 0 Then
            '    m.bError = 4
            '    Err.Raise vbObjectError + 1000, , "Please select a list"
            'End If
            
        Case 0
            m.bError = 2
            Err.Raise vbObjectError + 1000, , "Please select an Input type"
    End Select
    
    'Fill pre/pos words
    If Len(m.strFillPre) > 30 Then
        m.bError = 5
        Err.Raise vbObjectError + 1000, , _
            "Fill words before value must be 30 characters or less"
    End If
    If Len(m.strFillPost) > 30 Then
        m.bError = 6
        Err.Raise vbObjectError + 1000, , _
            "Fill words after value must be 30 characters or less"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cInput.ValidateFunctionInput", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Sub ValidateRuleInput()
On Error GoTo ErrSection:

    Select Case m.lParmTypeID
        Case C_RetNumericConstant, C_RetNumeric, C_RetTrueFalseConstant
            If m.bRequired Then
                If m.dFromValue <> 0 And m.dToValue <> 0 Then
                    If Val(m.vValue) < m.dFromValue Or Val(m.vValue) > m.dToValue Then
                        m.bError = 3
                        Err.Raise vbObjectError + 1000, , _
                            "The Input Value must be between " & _
                            VBA.Format(m.dFromValue, "general number") & _
                            " and " & VBA.Format(m.dToValue, "general number")
                    End If
                End If
            End If
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cInput.ValidateRuleInput", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub ValidateSystemInput()
On Error GoTo ErrSection:
    
    ValidateRuleInput

    'Optimization validations
    If m.bIfOptimize Then
        If (m.dOptFromValue < -100000000000# Or m.dOptFromValue > 100000000000#) Or _
           (m.dOptToValue < -100000000000# Or m.dOptToValue > 100000000000#) Then
            m.bError = 7
            Err.Raise vbObjectError + 1000, , _
                "Optimization From/To values are invalid.  They must be between -100,000,000,000 and 100,000,000,000"
        End If
        If m.dOptStepValue < 0 Or m.dOptStepValue > 100000000000# Then
            m.bError = 7
            Err.Raise vbObjectError + 1000, , _
                "STEP value is invalid.  Cannot be negative or larger then 10,000,000"
        End If
        If m.dOptFromValue > m.dOptToValue Then
            m.dOptToValue = m.dOptFromValue
        End If
        
        If Val(m.dFromValue) <> 0 Or Val(m.dToValue) <> 0 Then
            If m.dOptFromValue < m.dFromValue Then
                m.bError = 7
                Err.Raise vbObjectError + 1000, , _
                    "The FROM optimization value cannot be less than " & VBA.Format(m.dFromValue, "general number")
            End If
            If m.dOptToValue > m.dToValue Then
                m.bError = 7
                Err.Raise vbObjectError + 1000, , _
                    "The TO optimization value cannot be greater than " & VBA.Format(m.dToValue, "general number")
            End If
        End If
        
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cInput.ValidateSystemInput", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Function LoadRuleParm(ByVal lParmID As Long) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleParms] WHERE [ParmID]=" & Str(lParmID) & ";", dbOpenDynaset)
    If Not rs.EOF Then
        m.lRuleID = rs!RuleID
        m.lParmID = lParmID
        m.lParmSeq = rs!ParmSeq
        m.strParmName = rs!ParmName
        m.strParmDesc = rs!ParmDesc
        m.lParmTypeID = rs!ParmTypeID
        m.vDefaultValue = rs!DefaultValue
        m.dFromValue = rs!FromValue
        m.dToValue = rs!ToValue
        m.lListID = rs!ListID
        
        If m.lParmTypeID = C_RetBars Then
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemSecurities] WHERE [ParmID]=" & Str(lParmID) & ";", dbOpenDynaset)
            If Not rs.EOF Then
                m.lSystemID = rs!SystemNumber
                m.strPath = rs!Path
                m.strSymbol = rs!Symbol
                m.strMarketSymbol = rs!MarketSymbol
                m.strPeriodicity = rs!Periodicity
                m.strFormat = rs!Format
                m.strSecurityType = rs!SecurityType
                m.strSecurityName = rs!SecurityName
                m.strGroupID = NullChk(rs!GroupID)
                m.lSymbolID = NullChk(rs!SymbolID, 0&)
            Else
                m.lSystemID = 0&
                m.strPath = ""
                m.strSymbol = ""
                m.strMarketSymbol = ""
                m.strPeriodicity = ""
                m.strFormat = ""
                m.strSecurityType = ""
                m.strSecurityName = ""
                m.strGroupID = ""
                m.lSymbolID = 0&
            End If
        Else
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemParms] WHERE [ParmID]=" & Str(lParmID) & ";", dbOpenDynaset)
            If Not rs.EOF Then
                m.lSystemID = rs!SystemNumber
                m.vValue = rs!Value
                m.bIfOptimize = rs!IfOptimize
                m.dOptFromValue = rs!OptFromValue
                m.dOptToValue = rs!OptToValue
                m.dOptStepValue = rs!OptStepValue
                m.lOptListID = rs!OptListID
            Else
                m.lSystemID = 0&
                m.vValue = m.vDefaultValue
                m.bIfOptimize = False
                m.dOptFromValue = 0#
                m.dOptToValue = 0#
                m.dOptStepValue = 0#
                m.lOptListID = 0&
            End If
        End If
        
        LoadRuleParm = True
    End If

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "cInput.LoadRuleParm", eGDRaiseError_Raise, g.strAppPath
    
End Function

Public Sub SaveRuleParm(Optional ByVal lNewRuleID& = 0&, Optional ByVal lNewSystemID& = 0&)
On Error GoTo ErrSection:

    Dim rs As Recordset

    If lNewRuleID <> 0& Then m.lRuleID = lNewRuleID
    If lNewSystemID <> 0& Then m.lSystemID = lNewSystemID
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleParms] " & _
            "WHERE [RuleID]=" & Str(m.lRuleID) & " AND [ParmID]=" & Str(m.lParmID) & ";", dbOpenDynaset)
    If rs.EOF And rs.BOF Then
        rs.AddNew
        rs!RuleID = m.lRuleID
        m.lParmID = rs!ParmID
    Else
        rs.Edit
    End If

    rs!ParmSeq = m.lParmSeq
    rs!ParmName = m.strParmName
    rs!ParmDesc = m.strParmDesc
    rs!ParmTypeID = m.lParmTypeID
    rs!DefaultValue = Str(m.vDefaultValue)
    rs!FromValue = m.dFromValue
    rs!ToValue = m.dToValue
    rs!ListID = m.lListID
    rs.Update
    
    If m.lSystemID <> 0 Then
        If m.lParmTypeID = C_RetBars Then
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemSecurities] " & _
                        "WHERE [SystemNumber]=" & Str(m.lSystemID) & " AND [ParmID]=" & Str(m.lParmID) & ";", dbOpenDynaset)
            If rs.EOF And rs.BOF Then
                rs.AddNew
                rs!SystemNumber = m.lSystemID
                rs!ParmID = m.lParmID
            Else
                rs.Edit
            End If
                
            rs!Path = m.strPath
            rs!Symbol = m.strSymbol
            rs!MarketSymbol = m.strMarketSymbol
            rs!Periodicity = m.strPeriodicity
            rs!Format = m.strFormat
            rs!SecurityType = m.strSecurityType
            rs!SecurityName = m.strSecurityName
            rs!GroupID = m.strGroupID
            rs!SymbolID = m.lSymbolID
            rs.Update
        Else
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemParms] " & _
                        "WHERE [SystemNumber]=" & Str(m.lSystemID) & " AND [ParmID]=" & Str(m.lParmID) & ";", dbOpenDynaset)
            If rs.EOF And rs.BOF Then
                rs.AddNew
                rs!SystemNumber = m.lSystemID
                rs!ParmID = m.lParmID
            Else
                rs.Edit
            End If
                
            rs!Value = Str(m.vValue)
            rs!IfOptimize = m.bIfOptimize
            rs!OptFromValue = m.dOptFromValue
            rs!OptToValue = m.dOptToValue
            rs!OptStepValue = m.dOptStepValue
            rs!OptListID = m.lOptListID
            rs.Update
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cInput.SaveRuleParm", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Friend Sub SetPrivateData(mCopyFrom As mPrivate, Optional ByVal lNewID As Long = 0&, _
        Optional ByVal lNewRuleID As Long = 0&, Optional ByVal lNewSystemID As Long = 0&)
    
    ' get copy of all private data
    m = mCopyFrom
    m.lParmID = lNewID
    m.lRuleID = lNewRuleID
    m.lSystemID = lNewSystemID

End Sub

Public Function MakeCopy(Optional ByVal lNewID As Long = 0&, _
                Optional ByVal lNewRuleID As Long = 0&, Optional ByVal lNewSystemID As Long = 0&) As cInput
    
    Dim aCopy As New cInput
    
    aCopy.SetPrivateData m, lNewID, lNewRuleID, lNewSystemID
    Set MakeCopy = aCopy

End Function

Public Sub CopyRuleInfo(ByVal Parm As cInput)
On Error GoTo ErrSection:

    m.lParmSeq = Parm.ParmSeq
    m.strParmName = Parm.ParmName
    m.strParmDesc = Parm.ParmDesc
    m.lParmTypeID = Parm.ParmTypeID
    m.vDefaultValue = Parm.DefaultValue
    m.dFromValue = Parm.FromValue
    m.dToValue = Parm.ToValue
    m.lListID = Parm.ListID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cInput.CopyRuleInfo", eGDRaiseError_Raise, g.strAppPath
    
End Sub
