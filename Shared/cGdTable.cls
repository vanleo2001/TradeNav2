VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGdTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: cGdTable -- manages a table of data (fields and records)
'   - each "Field" is either a string or numeric array
'   - each field is "NumRecords" in size (size of each array)
'   - both fields and records are 0-based (from 0 to Num-1)
'   - fields can be defined by either "CreateField" (to create
'       a new array of the specified type) or by "AttachField"
'       (to attach to an existing gdArray)
Option Explicit

Private Type mPrivate
    hTable As Long
    bAutoDestroy As Boolean '(normally TRUE)
End Type
Private m As mPrivate

Private Sub Class_Initialize()
    m.hTable = gdCreateTable(0)
    m.bAutoDestroy = True
End Sub

Private Sub Class_Terminate()
    If m.bAutoDestroy Then
        gdDestroyTable m.hTable
    End If
End Sub

'Returns the "handle" to the gdTable (mostly used in DLL calls)
'- if bDetach = True, then the gdTable is detached from the
'   cGdTable object (up to caller to destroy the handle later)
Public Property Get TableHandle(Optional ByVal bDetach As Boolean = False) As Long
    Dim h&
    h = m.hTable
    If bDetach Then m.hTable = 0
    TableHandle = h
End Property

'This is allowed for some flexibility -- BUT SHOULD BE USED
'WITH GREAT CAUTION (know what you're doing and be careful)!
Public Sub SetTableHandle(ByVal NewTableHandle As Long, _
        ByVal bAutoDestroy As Boolean)
    
    'destroy current table
    If m.bAutoDestroy Then gdDestroyTable m.hTable
    
    m.hTable = NewTableHandle
    ' set flag to determine if will handle auto-destroying
    ' when this object gets done with it
    m.bAutoDestroy = bAutoDestroy
    
End Sub

Public Sub ClearField(ByVal nField&)
    If nField >= 0 Then
        gdClearField m.hTable, nField
    End If
End Sub

Public Sub Clear()
    gdClearField m.hTable, -1
End Sub

Public Property Get NumRecords() As Long
    NumRecords = gdNumRecords(m.hTable)
End Property

Public Property Let NumRecords(ByVal nNumRecords As Long)
    gdSetNumRecords m.hTable, nNumRecords
End Property

Public Property Get NumFields() As Long
    NumFields = gdNumFields(m.hTable)
End Property

Public Function FieldArrayHandle(ByVal nField&) As Long
    FieldArrayHandle = gdFieldArrayHandle(m.hTable, nField)
End Function

Public Function FieldArray(ByVal nField&, Optional ByVal bGetCopy As Boolean = True) As cGdArray
    Dim hArray As Long
    Dim aField As cGdArray
    hArray = gdFieldArrayHandle(m.hTable, nField)
    If hArray Then
        Set aField = New cGdArray
        aField.SetArrayHandle hArray, False
        If bGetCopy Then
            Set FieldArray = aField.MakeCopy
        Else
            Set FieldArray = aField
        End If
    Else
        Set FieldArray = Nothing
    End If
End Function

Public Property Get FieldName(ByVal nField&) As String
    Dim hString&
    hString = gdFieldName(m.hTable, nField)
    FieldName = gdGetStr(hString)
    gdDestroyString hString
End Property

Public Property Let FieldName(ByVal nField&, ByVal strFieldName As String)
    gdSetFieldName m.hTable, nField, strFieldName
End Property

Public Function FieldType(ByVal nField&) As eGdArray_Type
    FieldType = gdFieldType(m.hTable, nField)
End Function

Public Function FieldNum(ByVal strFieldName As String) As Long
    FieldNum = gdFieldNum(m.hTable, strFieldName)
End Function

Public Function CreateField(ByVal eArrayType As eGdArray_Type, Optional ByVal nField& = -1, Optional ByVal strFieldName$, _
        Optional ByVal dNullValue# = USE_DEFAULT_NULL) As Long
    
    If dNullValue = USE_DEFAULT_NULL Then
        CreateField = gdCreateField(m.hTable, eArrayType, nField, strFieldName)
    Else
        Dim a As New cGdArray
        a.Create eArrayType, 0, dNullValue
        CreateField = AttachField(a, nField, strFieldName)
        Set a = Nothing
    End If
    
End Function

Public Function AttachField(gdArray As cGdArray, Optional ByVal nField& = -1, Optional ByVal strFieldName$) As Long
    AttachField = gdAttachField(m.hTable, gdArray.ArrayHandle, nField, strFieldName)
End Function

' Using "Num" is a little more efficient than using "Item"
' for numeric arrays (since avoids variant conversions).
Public Property Get Num(ByVal nField&, ByVal nRecord&) As Double
    Num = gdGetTableNum(m.hTable, nField, nRecord)
End Property
Public Property Let Num(ByVal nField&, ByVal nRecord&, ByVal dNumber As Double)
    gdSetTableNum m.hTable, nField, nRecord, dNumber
End Property

Public Property Get Item(ByVal nField&, ByVal nRecord&) As Variant
Attribute Item.VB_UserMemId = 0
    If gdFieldType(m.hTable, nField) = eGDARRAY_Strings Then
        Item = gdGetTableString(m.hTable, nField, nRecord)
    Else
        Item = gdGetTableNum(m.hTable, nField, nRecord)
    End If
End Property

Public Property Let Item(ByVal nField&, ByVal nRecord&, ByVal vNewValue As Variant)
    If gdFieldType(m.hTable, nField) = eGDARRAY_Strings Then
        gdSetTableStr m.hTable, nField, nRecord, CStr(vNewValue)
    Else
        gdSetTableNum m.hTable, nField, nRecord, CDbl(vNewValue)
    End If
End Property

' To create an index for this table (to then be used with SortIndex)
Public Function CreateIndex(Optional ByVal nFilterByField& = -1) As cGdArray
    Dim hIndex As Long, aIndex As New cGdArray
    hIndex = gdCreateTableIndex(m.hTable, nFilterByField)
    aIndex.SetArrayHandle hIndex, True
    Set CreateIndex = aIndex
End Function
' To sort an index (created by CreateIndex) based on a field of this table
Public Function SortIndex(ByVal aIndex As cGdArray, ByVal nSortByField As Long, _
        Optional ByVal eSortFlags As eGdSort_Flags = eGdSort_Default) As Long

    SortIndex = gdSortTableIndex(m.hTable, aIndex.ArrayHandle, nSortByField, eSortFlags)
End Function

' To create a sorted index for this table
' - this routine does NOT sort the table itself, rather it just returns an array of
'       record numbers into this table that provides a sorted order
' - can create a compound index by passing multiple fields (pass primary field first)
' - more than one index on the same table can be created at a time
' - this is easier than calling both CreateIndex and SortIndex (though a little less flexible)
Public Function CreateSortedIndex(ByVal nSortByField1 As Long, Optional ByVal eSortFlags1 As eGdSort_Flags = eGdSort_Default, _
                Optional ByVal nSortByField2 As Long = -1, Optional ByVal eSortFlags2 As eGdSort_Flags = eGdSort_Default, _
                Optional ByVal nSortByField3 As Long = -1, Optional ByVal eSortFlags3 As eGdSort_Flags = eGdSort_Default, _
                Optional ByVal nSortByField4 As Long = -1, Optional ByVal eSortFlags4 As eGdSort_Flags = eGdSort_Default, _
                Optional ByVal nSortByField5 As Long = -1, Optional ByVal eSortFlags5 As eGdSort_Flags = eGdSort_Default) As cGdArray
    
    Dim aIndex As cGdArray
    Set aIndex = CreateIndex
    
    If nSortByField5 >= 0 And nSortByField5 < NumFields Then
        SortIndex aIndex, nSortByField5, eSortFlags5 Or eGdSort_Stable
    End If
    If nSortByField4 >= 0 And nSortByField4 < NumFields Then
        SortIndex aIndex, nSortByField4, eSortFlags4 Or eGdSort_Stable
    End If
    If nSortByField3 >= 0 And nSortByField3 < NumFields Then
        SortIndex aIndex, nSortByField3, eSortFlags3 Or eGdSort_Stable
    End If
    If nSortByField2 >= 0 And nSortByField2 < NumFields Then
        SortIndex aIndex, nSortByField2, eSortFlags2 Or eGdSort_Stable
    End If
    If nSortByField1 >= 0 And nSortByField1 < NumFields Then
        SortIndex aIndex, nSortByField1, eSortFlags1 Or eGdSort_Stable
    End If
    
    Set CreateSortedIndex = aIndex
End Function

Public Function IsNullAt(ByVal nField&, ByVal nRecord&) As Boolean
    IsNullAt = gdTableIsNullAt(m.hTable, nField, nRecord)
End Function

Public Function Serialize(ByVal vFile As Variant, ByVal bPut As Boolean) As Boolean
    Dim bSuccess As Boolean, hFile As Long
    If VarType(vFile) <> vbString Then
        'file handle is being passed (file already open)
        hFile = CLng(vFile)
    ElseIf bPut Then
        'open file for binary output
        hFile = gdFileOpen(CStr(vFile), "wb")
    Else
        'open file for binary input
        hFile = gdFileOpen(CStr(vFile), "rb")
    End If
    If hFile <> 0 Then
        'serialize the table
        bSuccess = gdSerializeTable(m.hTable, hFile, bPut)
        'close file (if opened here)
        If VarType(vFile) = vbString Then
            gdFileClose hFile
        End If
    End If
    'clear if unsuccessful read
    If Not bSuccess And Not bPut Then Clear
    Serialize = bSuccess
End Function

'Returns a "copy" of this table.
'(Note: array data not actually copied yet -- both
' tables are pointing to same array data until one of them
' is edited, then that one gets its own copy of the data)
Public Function MakeCopy() As cGdTable
    'first create a new table
    Dim aCopy As cGdTable
    Set aCopy = New cGdTable
    'then make it a "copy" of this one
    gdCopy aCopy.TableHandle, m.hTable
    Set MakeCopy = aCopy
End Function

'Makes current object a copy of the handle being passed.
Public Function CopyFromHandle(ByVal CopyFromTableHandle As Long) As Boolean
    
    'destroy current table
    If m.bAutoDestroy Then gdDestroyTable m.hTable
    
    'make copy of table
    m.hTable = gdCreateTable(0) '(must start with a valid object)
    If gdCopy(m.hTable, CopyFromTableHandle) <> 0 Then
        CopyFromHandle = True
    End If
    m.bAutoDestroy = True
    
End Function

'Sets data for entire table from a double-delimited string
Public Sub FromString(ByVal strTableData$, Optional ByVal strRecordDelim$ = vbLf, _
                Optional ByVal strFieldDelim$ = vbTab)

    If m.hTable = 0 Then m.hTable = gdCreateTable(0)
    gdTableFromString m.hTable, strTableData, strRecordDelim, strFieldDelim

End Sub

'Puts data from entire table into a double-delimited string
Public Function ToString(Optional ByVal strRecordDelim$ = vbLf, Optional ByVal strFieldDelim$ = vbTab, _
                    Optional ByVal bIncludeFieldNames As Boolean = False) As String

    Dim hString&, strFields$, i&
    
    If bIncludeFieldNames Then
        For i = 0 To NumFields - 1
            If i = NumFields - 1 Then
                strFields = strFields & FieldName(i) & strRecordDelim
            Else
                strFields = strFields & FieldName(i) & strFieldDelim
            End If
        Next
    End If
    
    hString = gdTableToString(m.hTable, strRecordDelim, strFieldDelim)
    If hString <> 0 Then
        ToString = strFields & gdGetStr(hString)
        gdDestroyArray hString
    Else
        ToString = strFields
    End If

End Function

'Gets a record of the table as a delimited string
Public Function GetRecord(ByVal nRecord&, Optional ByVal strFieldDelim$ = vbTab) As String

    Dim hString&
    hString = gdGetTableRecord(m.hTable, nRecord, strFieldDelim)
    If hString <> 0 Then
        GetRecord = gdGetStr(hString)
        gdDestroyArray hString
    End If

End Function

'Sets a record of the table from a delimited string
Public Sub SetRecord(ByVal strRecord$, Optional ByVal nRecord& = -1, _
                Optional ByVal strFieldDelim$ = vbTab)

    If m.hTable = 0 Then m.hTable = gdCreateTable(0)
    If nRecord < 0 Then nRecord = gdNumRecords(m.hTable)
    gdSetTableRecord m.hTable, strRecord, nRecord, strFieldDelim

End Sub

'Adds a record of the table from a delimited string (size increases by 1)
' - adds to end of the table by default
' - use nAtPosition to "insert" a record into the table at specified position
Public Sub AddRecord(ByVal strRecord$, Optional ByVal nInsertAtRecord& = -1, _
                Optional ByVal strFieldDelim$ = vbTab)

    If m.hTable = 0 Then m.hTable = gdCreateTable(0)
    gdAddTableRecord m.hTable, strRecord, nInsertAtRecord, strFieldDelim

End Sub

' removes one or more records from the table (records above are moved down)
Public Sub RemoveRecords(ByVal iFromItem As Long, Optional ByVal nNumToDelete As Long = 1)
    gdDeleteTableRecords m.hTable, iFromItem, nNumToDelete
End Sub

Public Function SearchAsIndex(ByVal aIndex As cGdArray, ByVal lSortByField&, _
        ByVal vSearchFor As Variant, Optional lMatchPos As Long, _
        Optional ByVal eSortFlags As eGdSort_Flags = eGdSort_Default, _
        Optional ByVal lFromItem& = 0, Optional ByVal lToItem& = 999999999) As Boolean

    Dim lPos As Long                    ' Matching or Insert At location in the array
    Dim bFound As Boolean               ' Whether or not the item was found
    Dim hString As Long                 ' Handle to a gdString

    If gdFieldType(m.hTable, lSortByField) <> eGDARRAY_Strings Then
        bFound = gdBinarySearchAsIndex(aIndex.ArrayHandle, _
                        gdFieldArrayHandle(m.hTable, lSortByField), CDbl(vSearchFor), _
                        lPos, eSortFlags, lFromItem, lToItem)
    Else
        hString = gdCreateArray(eGDARRAY_gdString)
        If hString <> 0 Then
            gdSetStr hString, 0, CStr(vSearchFor)
            bFound = gdBinarySearchAsIndex(aIndex.ArrayHandle, _
                            gdFieldArrayHandle(m.hTable, lSortByField), hString, _
                            lPos, eSortFlags, lFromItem, lToItem)
            gdDestroyArray hString
        End If
    End If
    
    If Not IsMissing(lMatchPos) Then
        lMatchPos = lPos
    End If
    
    SearchAsIndex = bFound

End Function

' Appends records from another cGdTable (data types of each field must match).
' - returns true if one or more records were appended
Public Function AppendFromTable(FromTable As cGdTable, Optional ByVal iFromRecord& = 0, Optional ByVal iToRecord& = -1) As Boolean

    AppendFromTable = gdAppendFrom(m.hTable, FromTable.TableHandle, iFromRecord, iToRecord)

End Function

