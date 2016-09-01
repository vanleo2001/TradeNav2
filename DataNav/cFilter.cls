VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' cFilter class:  defines a "Filter" (with one or more conditions)
Option Explicit

' type of expressions (for each condition)
Public Enum eCOND_ExprMode
    eCOND_Boolean = 1       ' true/false (Max = 1 or 0)
    eCOND_Percentiles = 2   ' use percentiles instead of values
    eCOND_LowOp = 4         ' if operator used for low number
    eCOND_HighOp = 8        ' if operator used for high number
    eCOND_Invert = 16       ' if invert operators for low/high numbers
End Enum

Private Type mPrivate
    strID As String     ' filename (without path)
    strName As String   ' displayed name
    strDesc As String   ' full description
    nCountMode As Long  ' display mode for counts
    
    bCustom As Boolean  ' true if custom Filter
    bIsActive As Boolean
    strRequired As String
    
    'arrays for Condition info
    aConditions As cGdArray 'string array: def. for each condition
    aCondFields As cGdArray 'to hold field# for each condition
    aCondCounts As cGdArray 'to hold counts for each condition
    
    aResults As New cGdArray
End Type
Private m As mPrivate

Private Sub Class_Initialize()
    
    Set m.aResults = New cGdArray
    m.aResults.Create eGDARRAY_TinyInts
    m.aResults.MakeConstantArray m.aResults.NullValue
    
    Set m.aConditions = New cGdArray
    m.aConditions.Create eGDARRAY_Strings
    Set m.aCondFields = New cGdArray
    m.aCondFields.Create eGDARRAY_Longs
    Set m.aCondCounts = New cGdArray
    m.aCondCounts.Create eGDARRAY_Longs
    
    m.bCustom = True 'default
    m.bIsActive = True

End Sub

Private Sub Class_Terminate()
    
    Set m.aResults = Nothing
    Set m.aConditions = Nothing
    Set m.aCondFields = Nothing
    Set m.aCondCounts = Nothing

End Sub

Public Property Get ID() As String
    ID = m.strID
End Property
Public Property Let ID(ByVal strID As String)
    m.strID = strID
End Property

Public Property Get Name() As String
    Name = m.strName
End Property
Public Property Let Name(ByVal strValue As String)
    m.strName = strValue
End Property

Public Property Get Desc() As String
    Desc = m.strDesc
End Property
Public Property Let Desc(ByVal strDesc As String)
    m.strDesc = strDesc
End Property

Public Property Get Custom() As Boolean
    Custom = m.bCustom
End Property
Public Property Let Custom(ByVal bCustom As Boolean)
    m.bCustom = bCustom
End Property

Public Property Get Conditions() As cGdArray
    Set Conditions = m.aConditions
End Property
Public Property Let Conditions(astrConditions As cGdArray)
    Set m.aConditions = astrConditions
End Property

Public Property Get CondFields() As cGdArray
    Set CondFields = m.aCondFields
End Property
Public Property Let CondFields(alCondFields As cGdArray)
    Set m.aCondFields = alCondFields
End Property

Public Property Get CondCounts() As cGdArray
    Set CondCounts = m.aCondCounts
End Property
Public Property Let CondCounts(alCondCounts As cGdArray)
    Set m.aCondCounts = alCondCounts
End Property

Public Property Get IsActive() As Boolean
    IsActive = m.bIsActive And HasModule(m.strRequired)
End Property
Public Property Let IsActive(ByVal bValue As Boolean)
    m.bIsActive = bValue
End Property

Public Property Get CountMode() As Long
    CountMode = m.nCountMode
End Property
Public Property Let CountMode(ByVal lCountMode As Long)
    m.nCountMode = lCountMode
End Property

Public Property Get Required() As String
    Required = m.strRequired
End Property

Public Sub AddToPool(Optional ByVal bReplaceArrays As Boolean = False)
On Error GoTo ErrSection:

    Dim strFieldID As String            ' ID of the Filter in the pool
    Dim lField As Long                  ' Field number of the Filter in the pool
    Dim lIndex As Long                  ' Index into the Filters collection

    ' Add object to collection...
    m.strID = UCase(m.strID)
    With g.SymbolPool.Filters
        lIndex = .Index(m.strID)
        If lIndex > 0 Then
            .Item(lIndex) = Me
        Else
            .Add Me, m.strID
        End If
    End With
    
    ' Add Results array to ArrayTable...
    With g.SymbolPool
        strFieldID = "FIL:" & m.strID
        lField = .FieldNumForID(strFieldID)
        If lField < 0 Or bReplaceArrays Then
            lField = .ArrayTable.AttachField(m.aResults, lField, m.strName)
            .FieldID(lField) = strFieldID
        ElseIf lField >= 0 Then
            .ArrayTable.FieldName(lField) = m.strName
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cFilter.AddToPool", eGDRaiseError_Raise

End Sub

' Load Filter from file
Public Function FromFile(ByVal strPath As String, ByVal strFile As String) As Boolean
On Error GoTo ErrSection:

    Dim i&, iPos&, strParm$, strText$, iCond&
    Dim aLines As New cGdArray

    ' determine if Provided or Custom
    If Len(strPath) = 0 Then
        strPath = App.Path & "\Provided\"
        If Not FileExist(strPath & strFile) Then
            strPath = App.Path & "\Custom\"
        End If
    End If

    m.strID = UCase(strFile)
    strPath = AddSlash(UCase(strPath))
    If Right(strPath, 8) = "\CUSTOM\" Then
        m.bCustom = True
    Else
        m.bCustom = False
    End If
    strFile = strPath & strFile

    ' to remove obsolete files
    If FileLength(strFile) < 5 Then KillFile strFile
    
    m.aConditions.Clear
    aLines.FromFile strFile
    For i = 0 To aLines.Size - 1
        strText = aLines(i)
        iPos = InStr(strText, "=")
        If iPos > 1 Then
            strParm = UCase(Trim(Left(strText, iPos - 1)))
            strText = Trim(Mid(strText, iPos + 1))
            Select Case strParm
                Case "NAME":
                    m.strName = strText
                Case "DESC":
                    m.strDesc = strText
                Case "COUNTMODE":
                    m.nCountMode = Val(strText)
                Case "ISACTIVE":
                    m.bIsActive = Val(strText)
                Case "REQUIRED":
                    m.strRequired = strText
                Case Else:
                    If Left(strParm, 9) = "CONDITION" And InStr(strText, "|") > 0 Then
                        ' TLB 3/4/2014: need this fix as a leftover from a transtional issue
                        strText = Replace(strText, "GRP:ALL U.S. STOCKS.GRP", "GRP:ALL STOCKS.GRP")
                        m.aConditions.Add Trim(strText)
                    End If
            End Select
            
            FromFile = True
        End If
    Next
    
ErrExit:
    Set aLines = Nothing
    Exit Function
    
ErrSection:
    Set aLines = Nothing
    RaiseError "cFilter.FromFile", eGDRaiseError_Raise

End Function

' Save Filter to file
Public Sub ToFile()
On Error GoTo ErrSection:

    Dim aLines As New cGdArray
    Dim strPath$, strFile$, i&
    
    strFile = m.strID
    If Len(strFile) = 0 Then
        strFile = NewCustomObjectName(".FIL")
        m.strID = UCase(strFile)
        m.bCustom = True
    End If
    
    If m.bCustom Then
        strPath = AddSlash(App.Path) & "Custom\"
    Else
        strPath = AddSlash(App.Path) & "Provided\"
    End If
    
    With aLines
        .Add "Name=" & m.strName
        .Add "Required=" & m.strRequired
        .Add "Desc=" & m.strDesc
        .Add "CountMode=" & Str(m.nCountMode)
        .Add "IsActive=" & Str(CLng(m.bIsActive))
        For i = 0 To m.aConditions.Size - 1
            .Add "Condition " & Str(i) & "=" & m.aConditions(i)
        Next
        .ToFile strPath & strFile
    End With

ErrExit:
    Set aLines = Nothing
    Exit Sub
    
ErrSection:
    Set aLines = Nothing
    RaiseError "cFilter.ToFile", eGDRaiseError_Raise

End Sub

#If 0 Then
' Edit Filter using frmFilter
Public Sub Edit(ByVal strPath$, ByVal strID$, Optional frm As frmFilter = Nothing)

    'Dim frm As frmFilter
    Dim bExistingForm As Boolean

    ' load from file (if existing)
    If Len(strID) > 0 Then
        If Not FromFile(strPath, strID) Then
            InfBox "i=[] ; h=ERROR ; Could not be loaded:|" & strID
            Exit Sub
        End If
    
        ' recalc results (so get counts, etc)
        CalcFilter
    Else
        m.strID = ""
    End If
        
    If frm Is Nothing Then
        Set frm = New frmFilter
    Else
        bExistingForm = True
    End If
    
    With frm
        If Not bExistingForm Then
            ' init form
            Screen.MousePointer = vbHourglass
            Set .mFilter = Me
            .txtName = m.strName
            .txtDesc = m.strDesc
            If m.nCountMode = 1 Then
                .optVisualCount = True
            End If
            .InitGrids
            SetEditorCaption frm, "Filter", .txtName
            If Trim(.txtName) = "" Then
                .cmdSaveAs.Visible = False
                .cmdRename.Visible = False
                .cmdCancel.Left = .cmdSaveAs.Left
            End If
                
            ' show form
            Screen.MousePointer = vbDefault
            .mSaved = False
            ShowForm frm, True
            Set .mFilter = Nothing
        Else
            .mSaved = True
        End If
        
        ' save
        If .mSaved Then
            ' get from form
            m.strName = Trim(.txtName)
            m.strDesc = Trim(.txtDesc)
            
            ' save to file
            ToFile
            
            ' add back into pool
            AddToPool True
            
            ' Refresh symbol grid dropdown and list
            frmSymbolGrid.RefreshGrid
        End If
    End With
    
    If Not bExistingForm Then
        Unload frm
        Set frm = Nothing
    End If

End Sub
#End If

Public Function CalcFilter() As Boolean
On Error GoTo ErrSection:

    Dim i&, nFld&, nSize&, bSuccess As Boolean
    Dim strFieldID$, dLow#, dHigh#
    Dim ExprMode As eCOND_ExprMode
    Dim aCond As New cGdArray
    Dim aLow As New cGdArray
    Dim aHigh As New cGdArray
    Dim aConst As New cGdArray
    Dim aArray As cGdArray

    ' init things
    bSuccess = True
    nSize = g.SymbolPool.NumRecords
    m.aCondFields.Clear
    m.aCondCounts.Clear
    aCond.Create eGDARRAY_TinyInts
    aLow.Create eGDARRAY_TinyInts
    aHigh.Create eGDARRAY_TinyInts
    aConst.Create eGDARRAY_Doubles
    aConst.MakeConstantArray 0, nSize
    
    ' init results array to a "copy" of the All Symbols group
    m.aResults.Clear False
    'nFld = g.SymbolPool.FieldNumForID("GRP:ALL SYMBOLS")
    'If nFld >= 0 Then
    '    Set m.aResults = g.SymbolPool.ArrayTable.FieldArray(nFld, True)
    'End If
    
    ' apply each condition
    For i = 0 To m.aConditions.Size - 1
        ' parse condition
        If Not ParseCondition(m.aConditions(i), strFieldID, ExprMode, dLow, dHigh) Then
            bSuccess = False
            Exit For
        End If
        
        ' get array for condition item
        If ExprMode = 0 Then
            bSuccess = False
            Exit For
        ElseIf (ExprMode And eCOND_Percentiles) <> 0 Then
            '(get percentiles array)
            strFieldID = "DSP:" & Mid(strFieldID, 5)
        End If
        nFld = g.SymbolPool.FieldNumForID(strFieldID)
        If nFld < 0 Then
            bSuccess = False
            Exit For
        End If
        m.aCondFields(i) = nFld
        
        ' do array calculations for this condition
        Set aArray = g.SymbolPool.ArrayTable.FieldArray(nFld)
        If (ExprMode And eCOND_Boolean) <> 0 Then
            ' Boolean
            If dHigh = 0 Then
                'if false
                aConst(0) = 0
                aCond.ArrayOperate aArray, "NOT", aConst 'Nothing
            Else
                'if true
                Set aCond = aArray.MakeCopy
            End If
        Else
            ' Low operator
            If (ExprMode And eCOND_LowOp) <> 0 Then
                aConst(0) = dLow
'aArray.ToFile (App.Path & "\chk\low1.chk")
                If (ExprMode And eCOND_Invert) <> 0 Then
                    aLow.ArrayOperate aArray, "<", aConst
                Else
                    aLow.ArrayOperate aArray, ">=", aConst
                End If
            End If
'aLow.ToFile (App.Path & "\chk\low.chk")
            ' High operator
            If (ExprMode And eCOND_HighOp) <> 0 Then
                aConst(0) = dHigh
                If (ExprMode And eCOND_Invert) <> 0 Then
                    aHigh.ArrayOperate aArray, ">", aConst
                Else
                    aHigh.ArrayOperate aArray, "<=", aConst
                End If
            End If
            ' Combine
            If (ExprMode And eCOND_HighOp) = 0 Then
                Set aCond = aLow.MakeCopy 'just using low
            ElseIf (ExprMode And eCOND_LowOp) = 0 Then
                Set aCond = aHigh.MakeCopy 'just using high
            ElseIf (ExprMode And eCOND_Invert) <> 0 Then
                aCond.ArrayOperate aLow, "OR", aHigh
            Else
                aCond.ArrayOperate aLow, "AND", aHigh
            End If
        End If
        
        ' combine with results
        If i = 0 Then
            ' if first condition, just make it a copy
            Set m.aResults = aCond.MakeCopy
        Else
            ' "AND" the condition to the results
            m.aResults.ArrayOperate m.aResults, "AND", aCond
        End If
        ' store count through this condition
        m.aCondCounts(i) = m.aResults.CountOf(1)
    Next

'BenchMark "Calc"

ErrExit:
    Set aArray = Nothing
    Set aCond = Nothing
    Set aLow = Nothing
    Set aHigh = Nothing
    Set aConst = Nothing
    Exit Function
    
ErrSection:
    Set aArray = Nothing
    Set aCond = Nothing
    Set aLow = Nothing
    Set aHigh = Nothing
    Set aConst = Nothing
    RaiseError "cFilter.CalcFilter", eGDRaiseError_Raise
    
End Function

Public Function CreateCondition(ByVal strFieldID$, _
    ByVal ExprMode As eCOND_ExprMode, ByVal dLow#, ByVal dHigh#, _
    Optional ByVal nCondition& = -1) As String
On Error GoTo ErrSection:

    Dim s$
    
    s = Trim(strFieldID) & "|" & Str(ExprMode) & "|" _
            & Str(dLow) & "|" & Str(dHigh)
    If nCondition >= 0 Then
        m.aConditions(nCondition) = s
    End If
    CreateCondition = s

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFilter.CreateCondition", eGDRaiseError_Raise

End Function

Public Function ParseCondition(ByVal strCond$, strFieldID$, _
        ExprMode As eCOND_ExprMode, dLow#, dHigh#) As Boolean
On Error GoTo ErrSection:

    Dim aSplit As New cGdArray
    
    aSplit.SplitFields strCond, "|"
    If aSplit.Size = 4 Then
        strFieldID = aSplit(0)
        ExprMode = Val(aSplit(1))
        dLow = Val(aSplit(2))
        dHigh = Val(aSplit(3))
        ParseCondition = True
        If (ExprMode And eCOND_Percentiles) <> 0 Then
            If dLow <= 0 Then
                dLow = 0
                ExprMode = (ExprMode And Not eCOND_LowOp)
            End If
            If dHigh >= 100 Then
                dHigh = 100
                ExprMode = (ExprMode And Not eCOND_HighOp)
            End If
        End If
    Else
        strFieldID = ""
        ExprMode = 0
        dLow = 0
        dHigh = 0
        ParseCondition = False
    End If
    
ErrExit:
    Set aSplit = Nothing
    Exit Function
    
ErrSection:
    Set aSplit = Nothing
    RaiseError "cFilter.ParseCondition", eGDRaiseError_Raise

End Function

Public Function InactiveCriteria() As cGdArray
On Error GoTo ErrSection:

    Dim Criteria As New cCriteria
    Dim lIndex As Long
    Dim astrInactive As New cGdArray
    Dim strID As String
    
    astrInactive.Create eGDARRAY_Strings
    
    For lIndex = 0 To m.aConditions.Size - 1
        strID = Parse(m.aConditions(lIndex), "|", 1)
        If UCase(Right(Trim(strID), 4)) = ".SCN" Then
            Set Criteria = g.SymbolPool.PoolObject(strID)
            If Criteria.IsActive = False Then
                astrInactive.Add Criteria.ID & "|" & Criteria.Name
            End If
        End If
    Next lIndex
    
    Set InactiveCriteria = astrInactive

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFilter.InactiveCriteria", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CriteriaInFilter
'' Description: Determine if the given criteria is in the filter
'' Inputs:      Criteria
'' Returns:     True if in Filter, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CriteriaInFilter(ByVal strCriteriaID As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    For lIndex = 0 To m.aConditions.Size - 1
        If UCase(Parse(Parse(m.aConditions(lIndex), "|", 1), ":", 2)) = UCase(strCriteriaID) Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    CriteriaInFilter = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFilter.CriteriaInFilter"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RenameCriteria
'' Description: Rename a criteria that is used by the filter
'' Inputs:      Old Criteria, New Criteria
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RenameCriteria(ByVal strOldCriteria As String, ByVal strNewCriteria As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrCondition As cGdArray       ' Condition split out into an array
    Dim bConditionChanged As Boolean    ' Did the condition change?
    Dim bFilterChanged As Boolean       ' Did the filter change?
    
    Set astrCondition = New cGdArray
    astrCondition.Create eGDARRAY_Strings
    
    bFilterChanged = False
    For lIndex = 0 To m.aConditions.Size - 1
        astrCondition.SplitFields m.aConditions(lIndex), "|"
        
        bConditionChanged = False
        If UCase(Parse(astrCondition(0), ":", 2)) = UCase(strOldCriteria) Then
            astrCondition(0) = Parse(astrCondition(0), ":", 1) & ":" & strNewCriteria
            bConditionChanged = True
        End If
        
        If bConditionChanged Then
            m.aConditions(lIndex) = astrCondition.JoinFields("|")
            bFilterChanged = True
        End If
    Next lIndex
    
    If bFilterChanged Then
        ToFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFilter.RenameCriteria"

End Sub

Friend Sub SetPrivateData(mCopyFrom As mPrivate)
    
    ' get copy of all private data
    m = mCopyFrom
    ' now make gdArrays a "copy" of the original
    Set m.aCondCounts = mCopyFrom.aCondCounts.MakeCopy
    Set m.aCondFields = mCopyFrom.aCondFields.MakeCopy
    Set m.aConditions = mCopyFrom.aConditions.MakeCopy
    Set m.aResults = mCopyFrom.aResults.MakeCopy

    m.strID = ""

End Sub

Public Function MakeCopy() As cFilter
    
    Dim aCopy As New cFilter
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

End Function
