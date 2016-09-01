VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cCriteria.cls
'' Description: Object to manipulate criteria
''
'' Author:      Genesis Financial Data Services
''              425 Wind Chime Pl
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Enum eCriteriaStorageType
    'Note: on 8/30/02 TLB swapped Float and Double so all
    'upgrades would auto-convert all criteria to floats
    '(in order to cut the memory pool to half the size)
    eCriteria_Float = 0      '4 bytes (single precision)
    eCriteria_Double = 1     '8 bytes (double precision)
    eCriteria_Long = 2       '4 bytes (+/- 2 billion)
    eCriteria_Short = 3      '2 bytes (+/- 32767)
    eCriteria_Byte = 4       '1 byte (+/- 127)
    eCriteria_Boolean = 5    '1 byte (0=false, +/-1=true, -128=null)
    eCriteria_Percentile = 6 '1 byte (only if stored as 0-100 percentiles)
    eCriteria_String = 7     '(probably won't be used!)
End Enum

Enum eCriteriaDisplayType
    eCriteria_Numeric = 0
    eCriteria_Text = 1
    eCriteria_Enumerated = 2 'if enumerated strings
End Enum

Enum eCriteriaPriceDisplay
    eCriteria_AutoRound = 0
    eCriteria_RoundToDecimal
    eCriteria_TradingUnits
End Enum

Enum eCriteriaUsageType
    eCriteria_FilterCriteria = 0
    eCriteria_QuoteBoardField = 1
End Enum

Private Type mPrivate
    strID As String     ' filename (without path)
    strName As String   ' displayed name
    strDesc As String
    strEnglishText As String
    strCodedText As String
    strFormattedText As String
    strGroupID As String
    StorageType As eCriteriaStorageType
    UsageType As eCriteriaUsageType
    lNumDays As Long
    lNumDaysOver As Long
    bIsWeekly As Boolean
    strRequired As String
    
    nPriceDisplay As eCriteriaPriceDisplay
    lDecimals As Long
    
    lCategoryID As Long                 'JM (05-21-2010): category number in SCN file assigned by Chad
    strCategoryName As String           'JM (05-21-2010): category name associated with category number in CriteriaCat.txt

    bCustom As Boolean
    bIsBoolean As Boolean
    bIsDirty As Boolean
    bIsActive As Boolean

    aValues As cGdArray
    aPercentiles As cGdArray
    aEnumeratedItems As cGdArray
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class memebers when the object is created
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:
    
    Set m.aValues = New cGdArray
    Set m.aPercentiles = New cGdArray
    Set m.aEnumeratedItems = New cGdArray
    m.aEnumeratedItems.Create eGDARRAY_Strings
    'defaults
    m.bCustom = True
    m.bIsBoolean = True
    m.bIsDirty = True
    m.UsageType = eCriteria_FilterCriteria
    'make result arrays const null's for now
    StorageType = eCriteria_Boolean
    m.lNumDays = 0&
    m.lNumDaysOver = 0&
    m.bIsActive = True
    
    m.nPriceDisplay = eCriteria_AutoRound
    m.lDecimals = 2&

    'default category ID and name to Miscellaneous
    m.lCategoryID = 999999
    m.strCategoryName = "Miscellaneous"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCriteria.Class.Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the object when the object is destroyed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:
    
    Set m.aEnumeratedItems = Nothing
    Set m.aPercentiles = Nothing
    Set m.aValues = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCriteria.Class.Terminate"
    
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
Public Property Let Name(ByVal strName As String)
    m.strName = strName
End Property

Public Property Get Desc() As String
    Desc = m.strDesc
End Property
Public Property Let Desc(ByVal strDesc As String)
    m.strDesc = strDesc
End Property

Public Property Get CodedText() As String
    CodedText = FixCodedText(m.strCodedText)
End Property

Public Property Let CodedText(ByVal strCodedText As String)
    m.strCodedText = strCodedText
End Property

Public Property Get FormattedText() As String
    FormattedText = m.strFormattedText
End Property
Public Property Let FormattedText(ByVal strFormattedText As String)
    m.strFormattedText = strFormattedText
End Property

Public Property Get EnglishText() As String
    EnglishText = m.strEnglishText
End Property
Public Property Let EnglishText(ByVal strEnglishText As String)
    m.strEnglishText = strEnglishText
End Property

Public Property Get Custom() As Boolean
    Custom = m.bCustom
End Property
Public Property Let Custom(ByVal bIsCustom As Boolean)
    m.bCustom = bIsCustom
End Property

Public Property Get ValuesArray() As cGdArray
    Set ValuesArray = m.aValues
End Property
Public Property Let ValuesArray(ByVal aValues As cGdArray)
    Set m.aValues = aValues
End Property

Public Property Get NumDays() As Long
    If m.lNumDaysOver > 0 Then
        NumDays = m.lNumDaysOver
    Else
        NumDays = m.lNumDays
    End If
End Property
Public Property Let NumDays(ByVal lValue As Long)
    m.lNumDays = lValue
End Property

Public Property Get NumDaysCalc() As Long
    NumDaysCalc = m.lNumDays
End Property
Public Property Let NumDaysCalc(ByVal lNumDaysCalc As Long)
    m.lNumDays = lNumDaysCalc
End Property

Public Property Get NumDaysOverride() As Long
    NumDaysOverride = m.lNumDaysOver
End Property
Public Property Let NumDaysOverride(ByVal lOverride As Long)
    m.lNumDaysOver = lOverride
End Property

Public Property Get IsWeekly() As Boolean
    IsWeekly = m.bIsWeekly
End Property
Public Property Let IsWeekly(ByVal bValue As Boolean)
    m.bIsWeekly = bValue
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = m.bIsDirty
End Property
Public Property Let IsDirty(ByVal bValue As Boolean)
    m.bIsDirty = bValue
End Property

Public Property Get IsActive() As Boolean
    IsActive = m.bIsActive And HasModule(m.strRequired)
End Property
Public Property Let IsActive(ByVal bValue As Boolean)
    m.bIsActive = bValue
End Property

Public Property Get IsBoolean() As Boolean
    IsBoolean = m.bIsBoolean
End Property
Public Property Let IsBoolean(ByVal bIsBoolean As Boolean)
    m.bIsBoolean = bIsBoolean
End Property

Public Property Get UsageType() As eCriteriaUsageType
    UsageType = m.UsageType
End Property
Public Property Let UsageType(ByVal vNewValue As eCriteriaUsageType)
    m.UsageType = vNewValue
End Property

Public Property Get StorageType() As eCriteriaStorageType
    StorageType = m.StorageType
End Property
Public Property Let StorageType(ByVal eStorageType As eCriteriaStorageType)
    
    ' reset data type for Values array
    m.StorageType = eStorageType
    Select Case m.StorageType
        Case eCriteria_Byte, _
             eCriteria_Boolean, _
             eCriteria_Percentile
                m.aValues.Create eGDARRAY_TinyInts
        Case eCriteria_Short
            m.aValues.Create eGDARRAY_Shorts
        Case eCriteria_Long
            m.aValues.Create eGDARRAY_Longs
        Case eCriteria_Float
            m.aValues.Create eGDARRAY_Floats
        Case eCriteria_Double
            m.aValues.Create eGDARRAY_Doubles
        Case eCriteria_String
            m.aValues.Create eGDARRAY_Strings
    End Select
    ' and make both arrays "null const" arrays
    m.aValues.MakeConstantArray m.aValues.NullValue
    m.aPercentiles.Create eGDARRAY_TinyInts
    m.aPercentiles.MakeConstantArray m.aPercentiles.NullValue
    
End Property

Public Property Get GroupID() As String
    If m.strGroupID = "" Then
        m.strGroupID = "GRP:ALL SYMBOLS.GRP"
    End If
    GroupID = m.strGroupID
End Property
Public Property Let GroupID(ByVal strGroupID As String)
    m.strGroupID = strGroupID
End Property
Public Property Get Required() As String
    Required = m.strRequired
End Property

Public Property Get PriceDisplay() As eCriteriaPriceDisplay
    PriceDisplay = m.nPriceDisplay
End Property
Public Property Let PriceDisplay(ByVal nPriceDisplay As eCriteriaPriceDisplay)
    m.nPriceDisplay = nPriceDisplay
End Property

Public Property Get DecimalPlaces() As Long
    DecimalPlaces = m.lDecimals
End Property
Public Property Let DecimalPlaces(ByVal lDecimals As Long)
    m.lDecimals = lDecimals
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddToPool
'' Description: Add the criteria to the symbol pool
'' Inputs:      Replace Arrays?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddToPool(Optional ByVal bReplaceArrays As Boolean = False)
On Error GoTo ErrSection:

    Dim strFldID$, nFld&, i&

    If m.UsageType <> eCriteria_FilterCriteria Then Exit Sub

    ' add object to collection
    m.strID = UCase(m.strID)
    With g.SymbolPool.Criterias
        i = .Index(m.strID)
        If i > 0 Then
            .Item(i) = Me
        Else
            .Add Me, m.strID
        End If
    End With
    
    ' add Values array to ArrayTable
    With g.SymbolPool
        strFldID = "DSV:" & m.strID
        nFld = .FieldNumForID(strFldID)
        If nFld < 0 Or bReplaceArrays Then
            bReplaceArrays = True
            nFld = .ArrayTable.AttachField(m.aValues, nFld, m.strName)
            .FieldID(nFld) = strFldID
        ElseIf nFld >= 0 Then
            .ArrayTable.FieldName(nFld) = m.strName
        End If
    
        ' and add Percentiles array to ArrayTable
        strFldID = "DSP:" & m.strID
        nFld = .FieldNumForID(strFldID)
        If m.bIsBoolean Then
            'if boolean, there is NOT a percentiles array
            If nFld >= 0 Then
                .ArrayTable.ClearField nFld
                .FieldID(nFld) = ""
            End If
        ElseIf nFld < 0 Or bReplaceArrays Then
            nFld = .ArrayTable.AttachField(m.aPercentiles, nFld, m.strName) ' & " (%)")
            .FieldID(nFld) = strFldID
        ElseIf nFld >= 0 Then
            .ArrayTable.FieldName(nFld) = m.strName
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCriteria.AddToPool", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromFile
'' Description: Load the criteria from a file
'' Inputs:      Path of the File, Name of the File
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FromFile(ByVal strPath As String, ByVal strFile As String) As Boolean
On Error GoTo ErrSection:

    Dim i&, iPos&, strParm$, strText$
    Dim aLines As New cGdArray

#If 0 Then
'Public Function FromFile(ByVal strFile As String, Optional ByVal strPath As String = "") As Boolean
    ' check if path is passed as 1st arg (which is the way it used to be)
    If InStr(strFile, "\") > 0 And Len(strPath) > 0 Then
        ' if so, swap the args
        strText = strFile
        strFile = strPath
        strPath = strText
    Else
        strText = ""
    End If
#End If

    m.strID = UCase(strFile)

    ' determine if Provided or Custom
    If Len(strPath) = 0 Then
        strPath = App.Path & "\Provided\"
        If Not FileExist(strPath & strFile) Then
            strPath = App.Path & "\Custom\"
        End If
    End If
    
    strPath = AddSlash(UCase(strPath))
   
    If Right(strPath, 8) = "\CUSTOM\" Then
        m.bCustom = True
    Else
        m.bCustom = False
    End If
    strFile = strPath & strFile

    ' to remove obsolete files
    If FileLength(strFile) < 5 Then KillFile strFile
    
    ' initialize to -1 so will know if it has been stored yet or not
    ' (if not stored yet, we want to auto-detect in the editor if it had been set manually)
    m.lNumDaysOver = -1

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
                Case "ENGLISH":
                    ' TLB 9/20/2012: rename "LW Sentiment" functions
                    strText = Replace(strText, "LW Sentiment", "TN Consensus")
                    m.strEnglishText = strText
                Case "CODED":
                    ' TLB 9/20/2012: rename "LW Sentiment" functions
                    strText = Replace(strText, "LWSentiment", "TNConsensus")
                    m.strCodedText = FixCodedText(strText)
                Case "FORMATTED":
                    ' TLB 9/20/2012: rename "LW Sentiment" functions
                    strText = Replace(strText, "LW Sentiment", "TN Consensus")
                    m.strFormattedText = strText
                Case "USAGETYPE":
                    m.UsageType = Val(strText)
                Case "STORAGETYPE":
                    'sets local results arrays to null const
                    StorageType = Val(strText)
                Case "NUMDAYS":
                    m.lNumDays = Val(strText)
                Case "NUMDAYSOVER":
                    m.lNumDaysOver = Val(strText)
                Case "ISBOOLEAN":
                    m.bIsBoolean = Val(strText)
                Case "ISDIRTY":
                    m.bIsDirty = Val(strText)
                Case "ISACTIVE":
                    m.bIsActive = Val(strText)
                Case "ISWEEKLY":
                    m.bIsWeekly = Val(strText)
                Case "GROUPID":
                    m.strGroupID = strText
                Case "REQUIRED":
                    m.strRequired = strText
                Case "PRICEDISPLAY":
                    m.nPriceDisplay = Val(strText)
                Case "DECIMALPLACES":
                    m.lDecimals = Val(strText)
            End Select
            
            FromFile = True
        End If
    Next
    
    'assume miscellaneous
    m.lCategoryID = 999999
    m.strCategoryName = "Miscellaneous"
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCriteria.FromFile", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToFile
'' Description: Save the criteria to the file
'' Inputs:      Keep same file date?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ToFile(Optional ByVal bKeepSameFileDate As Boolean = False)
On Error GoTo ErrSection:

    Dim aLines As New cGdArray
    Dim strPath$, strFile$, dFileDate#
    
    strFile = m.strID
    If Len(strFile) = 0 Then
        If m.UsageType = eCriteria_QuoteBoardField Then
            strFile = NewCustomObjectName(".QBF")
        Else
            strFile = NewCustomObjectName(".SCN")
        End If
        m.strID = UCase(strFile)
        m.bCustom = True
    End If
    strPath = App.Path & "\"
    If m.bCustom Then
        strPath = strPath & "Custom\"
    Else
        strPath = strPath & "Provided\"
    End If
    
    With aLines
        ' create lines of file
        .Add "Name=" & m.strName
        .Add "Required=" & m.strRequired
        .Add "Desc=" & m.strDesc
        .Add "StorageType=" & Str(CLng(m.StorageType))
        .Add "English=" & m.strEnglishText
        .Add "Coded=" & m.strCodedText
        .Add "Formatted=" & m.strFormattedText
        .Add "IsBoolean=" & Str(CLng(m.bIsBoolean))
        .Add "UsageType=" & Str(CLng(m.UsageType))
        .Add "IsDirty=" & Str(CLng(m.bIsDirty))
        .Add "NumDays=" & Str(m.lNumDays)
        .Add "NumDaysOver=" & Str(m.lNumDaysOver)
        .Add "IsActive=" & Str(CLng(m.bIsActive))
        .Add "IsWeekly=" & Str(CLng(m.bIsWeekly))
        If m.strGroupID = "" Then m.strGroupID = "GRP:ALL SYMBOLS.GRP"
        .Add "GroupID=" & m.strGroupID
        .Add "PriceDisplay=" & Str(m.nPriceDisplay)
        .Add "DecimalPlaces=" & Str(m.lDecimals)
        
        If m.lCategoryID > 0 And m.strCategoryName <> "Miscellaneous" Then
            .Add "Grouping=" & Str(m.lCategoryID)
        End If
        
        ' save to file
        strFile = strPath & strFile
        dFileDate = FileDate(strFile)
        .ToFile strFile
        If bKeepSameFileDate = True And dFileDate > 0 Then
            gdSetFileDate strFile, dFileDate
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCriteria.ToFile", eGDRaiseError_Raise

End Sub

#If 0 Then
' Edit using frmCriteria
Public Function Edit(ByVal strPath$, ByVal strID$, Optional frm As frmCriteria = Nothing) As Boolean

    Dim strNewCodedText$, bReplaceArrays As Boolean
    Dim bExistingForm As Boolean
    
    ' load from file (if existing)
    If Len(strID) > 0 Then
        If Not FromFile(strPath, strID) Then
            InfBox "i=[] ; h=ERROR ; Could not be loaded:|" & strID
            Exit Function
        End If
    Else
        m.strID = ""
    End If
        
    If frm Is Nothing Then
        Set frm = New frmCriteria
    Else
        bExistingForm = True
    End If
    
    With frm
        If Not bExistingForm Then
            ' init form
            Screen.MousePointer = vbHourglass
            .txtName = m.strName
            .txtDesc = m.strDesc
            .Editor1.Text = m.strEnglishText
            .CodedText = m.strCodedText
            .IsBoolean = m.bIsBoolean
            If m.lNumDays < 0 Then m.lNumDays = 0
            .txtNumDays.Text = Str(m.lNumDays)
            If m.bIsWeekly Then .optWeekly.Value = True Else .optDaily.Value = True
            If m.strName = "" Then
                .cmdSaveAs.Visible = False
                .cmdRename.Visible = False
                .cmdVerify.Top = .cmdSaveAs.Top
                .cmdCancel.Top = .cmdRename.Top
            Else
                .cmdSaveAs.Visible = True
                .cmdRename.Visible = True
            End If
            .cmdVerify.Enabled = False
            
            If m.strGroupID <> "" Then
                .cboSymbolGroups.ComboItems(m.strGroupID).Selected = True
            Else
                .cboSymbolGroups.ComboItems("GRP:ALL SYMBOLS.GRP").Selected = True
            End If
            
            If Me.UsageType = eCriteria_FilterCriteria Then
                SetEditorCaption frm, "Criteria", .txtName
                .cmdSave.Enabled = False
            Else
                SetEditorCaption frm, "Quote Board Field", .txtName
                .cmdSave.Enabled = True
                .cmdSaveAs.Visible = False
                .cmdRename.Top = .cmdSaveAs.Top
            End If
            
            ' show form
            Screen.MousePointer = vbDefault
            .mbSaved = False
            Set .Criteria = Me
            ShowForm frm, True
            Set .Criteria = Nothing
        Else
            .mbSaved = True
        End If
        
        ' save
        If .mbSaved Then
            Edit = True
            ' see if "dirty"
            If m.bIsWeekly <> .optWeekly.Value Then m.bIsDirty = True
            If m.lNumDays <> CLng(Trim(.txtNumDays)) Then m.bIsDirty = True
            If UCase(Trim(m.strCodedText)) <> UCase(Trim(.CodedText)) Then m.bIsDirty = True
            ' get from form
            m.strName = Trim(.txtName)
            m.strDesc = Trim(.txtDesc)
            m.lNumDays = CLng(Trim(.txtNumDays))
            m.strEnglishText = Trim(.Editor1.Text)
            m.strCodedText = Trim(.CodedText)
            m.bIsBoolean = .IsBoolean
            m.bIsWeekly = .optWeekly.Value
            m.strGroupID = .cboSymbolGroups.SelectedItem.Key
            If m.bIsDirty Then
                'and put null const arrays into table
                '(will need to recalc later, when do Criterias)
                bReplaceArrays = True
                If .IsBoolean Then
                    StorageType = eCriteria_Boolean
                Else
                    StorageType = eCriteria_Float 'Double
                End If
            End If
            
            ' save to file
            ToFile
            
            ' add back into pool
            AddToPool bReplaceArrays
        
            ' Refresh symbol grid dropdown and list
            frmSymbolGrid.RefreshGrid
        End If
    End With
    If Not bExistingForm Then
        Unload frm
        Set frm = Nothing
    End If
    
End Function
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcPercentiles
'' Description: Calculate the percentiles for the criteria
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CalcPercentiles()
On Error GoTo ErrSection:
    
    Dim nSize&, i&, p&, nMin&, nMax&
    Dim hValues&, hPercentiles&, hSortedIdx&
    
    If m.UsageType <> eCriteria_FilterCriteria Then Exit Sub
    
    ' see if calculating percentiles is valid
    If m.bIsBoolean Or m.aValues.IsConstantValue Then
        m.aPercentiles.MakeConstantArray m.aPercentiles.NullValue
        Exit Sub
    End If
    
    ' work with gdArray handles directly for speed
    nSize = m.aValues.Size
    m.aPercentiles.Clear '(to clear constant value)
    m.aPercentiles.Size = nSize
    hValues = m.aValues.ArrayHandle
    hPercentiles = m.aPercentiles.ArrayHandle
    If nSize > 0 And hPercentiles <> 0 Then
        ' created a sorted index of the values
        hSortedIdx = gdCreateArray(eGDARRAY_Longs)
        ''If gdIndexList(hSortedIdx, 0, hValues, eGdSort_DeleteNullValues) Then
        If gdSortAsIndex(hSortedIdx, hValues, True, eGdSort_DeleteNullValues, 0, -1) Then
            ' get new size (after nulls deleted)
            nSize = gdGetSize(hSortedIdx)
            If nSize > 0 Then
                ' for each percentile (1% segments of sorted values)
                nMax = nSize - 1
                For p = 99 To 0 Step -1
                    ' find min and max item# for that segment
                    nMin = Int(nSize / 100# * p + 0.5)
                    If nMax < nMin Then nMax = nMin
                    ' set the corresponding items to this percentile
                    For i = nMin To nMax
                        gdSetNum hPercentiles, gdGetNum(hSortedIdx, i), p
                    Next
                    ' set nMax at next item below
                    nMax = nMin - 1
                Next
            End If
        End If
        gdDestroyArray hSortedIdx
    End If
    
    If 0 Then
        'debugging
        Dim fh%
        fh = FreeFile
        Open App.Path & "\perc.chk" For Output As #fh
        For i = 0 To gdGetSize(hValues) - 1
            Print #fh, Format(gdGetNum(hPercentiles, i), "0000") _
                & " " & Format(gdGetNum(hValues, i), "000000.0000")
        Next
        Close #fh
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCriteria.CalcPercentiles", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private properties for a copy of the criteria
'' Inputs:      Copy From
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    ' get copy of all private data
    m = mCopyFrom
    ' now make gdArrays a "copy" of the original
    Set m.aValues = mCopyFrom.aValues.MakeCopy
    Set m.aPercentiles = mCopyFrom.aPercentiles.MakeCopy
    Set m.aEnumeratedItems = mCopyFrom.aEnumeratedItems.MakeCopy
    m.strID = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCriteria.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the criteria
'' Inputs:      None
'' Returns:     New Copy
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cCriteria
On Error GoTo ErrSection:

    Dim aCopy As New cCriteria
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCriteria.MakeCopy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the criteria to the appropriate places
'' Inputs:      Replace Arrays?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save(ByVal bReplaceArrays As Boolean)
On Error GoTo ErrSection:

    If bReplaceArrays Then
        If m.bIsBoolean Then
            StorageType = eCriteria_Boolean
        Else
            StorageType = eCriteria_Float 'Double
        End If
    End If
    
    ' Save to file
    ToFile
    
    ' Add back into pool
    AddToPool bReplaceArrays

    ' Refresh symbol grid dropdown and list
    frmSymbolGrid.RefreshGrid
    
    ' Refresh the quote board if it uses this criteria as a Quote Board Field...
    frmQuotes.RefreshQBF m.strID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCriteria.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixCodedText
'' Description: Fix the coded text of the criteria
'' Inputs:      Old Coded Text
'' Returns:     New Coded Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FixCodedText(ByVal strCodedText As String) As String
On Error GoTo ErrSection:

    Dim iPos&, iPos2&
    
    ' Look for older RSI that needs to be converted (first arg changed from Bars to Array) ...
    ' from: ~01003RSI ~16001( ~07007Market1 ~22001, ~130017
    ' to:   ~01003RSI ~16001( ~01005Close ~16001( ~07007Market1 ~17001) ~22001, ~130017
    Do While True
        iPos = InStr(strCodedText, "~01003RSI ~16001( ~07")
        If iPos = 0 Then Exit Do
        iPos = iPos + 18 'point to the ~07
        iPos2 = InStr(iPos, strCodedText, "~22001,") 'point to after the market
        If iPos2 = 0 Then Exit Do
        strCodedText = Left(strCodedText, iPos - 1) & "~01005Close ~16001( " _
            & Mid(strCodedText, iPos, iPos2 - iPos) & "~17001) " & Mid(strCodedText, iPos2)
    Loop

    FixCodedText = strCodedText
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCriteria.FixCodedText"
    
End Function

Public Property Get CategoryID() As Long
On Error GoTo ErrSection:

    CategoryID = m.lCategoryID

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCriteria.Get.CategoryID"
    
End Property

Public Property Let CategoryID(ByVal nID&)
On Error GoTo ErrSection:
    
    m.lCategoryID = nID

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCriteria.Let.CategoryID"
    
End Property

Public Property Get CategoryName() As String
On Error GoTo ErrSection:

    CategoryName = m.strCategoryName

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCriteria.Get.CategoryName"
    
End Property

Public Property Let CategoryName(ByVal strCatName$)
On Error GoTo ErrSection:

    m.strCategoryName = strCatName

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCriteria.Let.CategoryName"
    
End Property

