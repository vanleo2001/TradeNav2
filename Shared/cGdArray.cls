VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGdArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Class:  cGdArray  (requires mGdDll.BAS module and G32_GD.DLL)
'Can manage an array of numbers, strings, or even objects,
'   with all the gdArray features (auto-growing, null-values, etc.).
'To use:
' 1) Dim as an instance of this class.
'       e.g. Dim aArray As New cGdArray
' 2) Optional: declare what type of array it is.
'       e.g. aArray.Create(eGDARRAY_Longs) for array of longs
'       (if skip this step, the array will be created and the
'       array type will be automatically chosen when the first
'       item is set: either as an array of strings or doubles)
' 3) Set items of the array (array will auto-grow as needed).
'       e.g. aArray(3) = 43.5
'        or  aArray(3) = "A string"
'        or  Set aArray(3) = AnObject
' 4) Retrieve items of the array.
'       e.g. dNum = aArray(3)
'        or  strTemp = aArray(3)
'        or  Set AnObject = aArray(3)
Option Explicit

Private Type mPrivate
    hArray As Long
    eType As eGdArray_Type
    bAutoDestroy As Boolean '(normally TRUE)
End Type
Private m As mPrivate

Private Sub Class_Initialize()
    m.hArray = 0
    m.eType = eGDARRAY_NoArray
    m.bAutoDestroy = True 'by default
    'gAllocChk = gAllocChk + 1
End Sub

Private Sub Class_Terminate()
    Destroy
    'gAllocChk = gAllocChk - 1
End Sub

Public Sub Destroy()
    If m.hArray <> 0 And m.bAutoDestroy Then
        gdDestroyArray m.hArray
        If m.hArray <> 0 Then
            'HOUSTON, WE HAVE A PROBLEM!  Somehow, we still have
            'a pointer to something that used to be a valid gdArray
            'but it had already been destroyed.
            '-- THIS INDICATES A PROGRAMMING ERROR (most likely
            'the "SetArrayHandle" was used in an invalid fashion).
            m.hArray = 0
        End If
    End If
    m.hArray = 0
    m.eType = eGDARRAY_NoArray
    m.bAutoDestroy = True 'reset flag to default
End Sub

Private Sub CreateArrayOfType(vNewItem As Variant)
        
    If VarType(vNewItem) = vbString Then
        ' assume a string array
        Create eGDARRAY_Strings
    ElseIf VarType(vNewItem) = vbObject Then
        ' assume an object array
        Create eGDARRAY_Objects
    Else
        ' assume an array of doubles
        Create eGDARRAY_Doubles
    End If

End Sub

' creates the array (defines type and allocates memory)
Public Function Create(ByVal eArrayType As eGdArray_Type, _
        Optional ByVal nSize As Long = 0, _
        Optional ByVal dNullValue# = USE_DEFAULT_NULL) As Boolean
    
    ' if existing array, destroy it first
    Destroy
    
    ' then create new array
    m.hArray = gdCreateArray(eArrayType, nSize, dNullValue)
    If m.hArray <> 0 Then
        m.eType = eArrayType
        Create = True
    Else
        m.eType = eGDARRAY_NoArray
        Create = False
    End If
    
End Function

' To retrieve an item from the array.
Public Property Get Item(ByVal nOffset&) As Variant
Attribute Item.VB_UserMemId = 0
    
    Dim vObject As Variant
    Select Case m.eType
        Case eGDARRAY_NoArray:
            Item = "" ' not created yet!
        Case eGDARRAY_Strings:
            Item = gdGetStr(m.hArray, nOffset)
        Case eGDARRAY_gdString:
            Item = gdGetStr(m.hArray, 0)
        Case eGDARRAY_Objects:
            ' get the object as a variant
            Set Item = Nothing
            If gdGetVariant(m.hArray, nOffset, vObject) Then
                If VarType(vObject) <> vbEmpty Then
                    Set Item = vObject
                End If
            End If
            Set vObject = Nothing
        Case Else:
            Item = gdGetNum(m.hArray, nOffset)
    End Select

End Property

' Non-object arrays use "Let" version to set an item.
Public Property Let Item(ByVal nOffset&, ByVal vNewValue As Variant)

    ' see if need to create array first
    If m.hArray = 0 Then
        CreateArrayOfType vNewValue
    End If
    
    ' set item
    Select Case m.eType
        Case eGDARRAY_Strings:
            gdSetStr m.hArray, nOffset, Str(vNewValue)
        Case eGDARRAY_gdString:
            gdSetStr m.hArray, 0, Str(vNewValue)
        Case eGDARRAY_Objects:
            ' error
            Err.Raise vbObjectError + 1, "cGdArray", _
                "Must use 'SET' to set an item of an array of OBJECTS."
        Case Else:
            gdSetNum m.hArray, nOffset, Val(vNewValue)
    End Select
    
End Property

' Object arrays should use the "Set" version to set an item
' (e.g. Set aObjects(i) = AnObject)
Public Property Set Item(ByVal nOffset&, ByVal vNewValue As Variant)
    
    ' see if need to create array first
    If m.hArray = 0 Then
        CreateArrayOfType vNewValue
    End If
    
    ' set item
    If m.eType = eGDARRAY_Objects Then
        gdSetVariant m.hArray, nOffset, vNewValue
    Else
        ' error
        Err.Raise vbObjectError + 2, "cGdArray", _
            "Can only 'SET' an item for an array of OBJECTS."
    End If

End Property

' Using "Num" is a little more efficient than using "Item"
' for numeric arrays (since avoids variant conversions).
Public Property Get Num(ByVal nOffset&) As Double
    Num = gdGetNum(m.hArray, nOffset)
End Property
Public Property Let Num(ByVal nOffset&, ByVal vNewValue As Double)
    ' see if need to create array first
    If m.hArray = 0 Then
        CreateArrayOfType vNewValue
    End If
    gdSetNum m.hArray, nOffset, vNewValue
End Property

' Adds an item to the array (size increases by 1)
' - adds to end of the array by default
' - use nAtPosition to "insert" an item into the array at specified position
Public Sub Add(ByVal vNewValue As Variant, Optional ByVal nAtPosition& = -1)
    
    ' see if need to create array first
    If m.hArray = 0 Then
        CreateArrayOfType vNewValue
    End If
    
    If nAtPosition < 0 Then nAtPosition = gdGetSize(m.hArray)
    
    ' insert item
    Select Case m.eType
        Case eGDARRAY_Objects:
            gdInsertVariant m.hArray, vNewValue, nAtPosition
        Case eGDARRAY_Strings:
            gdInsertStr m.hArray, Str(vNewValue), nAtPosition
        Case eGDARRAY_gdString:
            'concatenate to a gdString
            gdSetStr m.hArray, 0, gdGetStr(m.hArray, 0) & Str(vNewValue)
        Case Else:
            gdInsertNum m.hArray, Val(vNewValue), nAtPosition
    End Select
    
End Sub

' returns size of the array (has items from 0 thru size-1)
Public Property Get Size() As Long
    Size = gdGetSize(m.hArray)
End Property

' sets size of the array
Public Property Let Size(ByVal nSize As Long)
    If m.hArray <> 0 Then gdSetSize m.hArray, nSize, True
End Property

'Returns the "handle" to the gdArray (mostly used in DLL calls)
'- if bDetach = True, then the gdArray is detached from the
'   cGdArray object (up to caller to destroy the handle later)
Public Property Get ArrayHandle(Optional ByVal bDetach As Boolean = False) As Long
    Dim h&, bAutoDestroy As Boolean
    h = m.hArray
    If bDetach Then
        bAutoDestroy = m.bAutoDestroy
        m.bAutoDestroy = False
        Destroy
        m.bAutoDestroy = bAutoDestroy
    End If
    ArrayHandle = h
End Property

'This is allowed for some flexibility -- BUT SHOULD BE USED
'WITH GREAT CAUTION (know what you're doing and be careful)!
Public Sub SetArrayHandle(ByVal NewArrayHandle As Long, _
        ByVal bAutoDestroy As Boolean)
    
    Destroy 'destroy current array if exists
    m.hArray = NewArrayHandle
    m.eType = gdGetType(m.hArray)
    ' set flag to determine if will handle auto-destroying
    ' the gdArray when this object gets done with it
    m.bAutoDestroy = bAutoDestroy
    
End Sub

' performs a quick-sort on the array
' (bIgnoreCase and iAtStringPosition only valid for string arrays)
Public Sub Sort(Optional ByVal eSortFlags As eGdSort_Flags = eGdSort_Default, _
        Optional ByVal iFromItem As Long = 0, _
        Optional ByVal iToItem As Long = 999999999)
    
    gdSort m.hArray, eSortFlags, iFromItem, iToItem
    
End Sub

' reinitializes array (sets size to 0)
Public Sub Clear(Optional ByVal bFreeExtraMemory As Boolean = True)
    gdClear m.hArray, bFreeExtraMemory
End Sub

' removes one or more items from the array (items above are moved down)
Public Sub Remove(ByVal iFromItem As Long, Optional ByVal nNumToDelete As Long = 1)
    gdDeleteItems m.hArray, iFromItem, nNumToDelete
End Sub

' returns the value considered "null" (empty) for an item in the array
Public Property Get NullValue() As Variant
    NullValue = gdNullValue(m.hArray)
End Property

' Pre-allocates the requested amount of memory
' (can preserve data, but does NOT set the "size")
Public Sub Reserve(ByVal nForNumItems As Long, Optional ByVal bPreserveData = True)
    gdReserve m.hArray, nForNumItems, bPreserveData
End Sub

' performs efficient math and comparison operations for entire arrays
' (i.e. ResultArray = Array1 "op" Array2)
Public Function ArrayOperate(gdArray1 As cGdArray, _
    ByVal strOperation$, gdArray2 As cGdArray)

    Dim hArray2 As Long
    If gdArray2 Is Nothing Then
        hArray2 = 0 '(Array2 ignored for unary operators)
    Else
        hArray2 = gdArray2.ArrayHandle
    End If
    ArrayOperate = gdArrayOperate(m.hArray, gdArray1.ArrayHandle, _
            strOperation, hArray2)
    
End Function

' To parse fields from a string into an array (each field = item in array)
Public Function SplitFields(ByVal strToParse$, _
        Optional ByVal strDelimiters$ = " ," & vbTab & vbLf, _
        Optional ByVal nNumFields& = 0) As Long
        
    ' if not created yet, make it a string array
    If m.hArray = 0 Then Create eGDARRAY_Strings
                
    SplitFields = gdSplitFields(m.hArray, strToParse, strDelimiters, nNumFields)
End Function

' returns lowest value in array (or piece of it)
Public Function MinValue(Optional ByVal nFromItem& = 0, Optional ByVal nToItem& = -1) As Double
    MinValue = gdMinValue(m.hArray, nFromItem, nToItem)
End Function

' returns highest value in array (or piece of it)
Public Function MaxValue(Optional ByVal nFromItem& = 0, Optional ByVal nToItem& = -1) As Double
    MaxValue = gdMaxValue(m.hArray, nFromItem, nToItem)
End Function

' returns difference between highest and lowest values in array (or piece of it)
Public Function Range(Optional ByVal nFromItem& = 0, Optional ByVal nToItem& = -1) As Double
    Range = gdRange(m.hArray, nFromItem, nToItem)
End Function

' returns # of non-null items in array (or piece of it)
Public Function CountOfNonNull(Optional ByVal nFromItem& = 0, Optional ByVal nToItem& = -1) As Long
    CountOfNonNull = gdCount(m.hArray, nFromItem, nToItem)
End Function

' returns # of items that match a specified value
Public Function CountOf(ByVal dMatchValue#, Optional ByVal nFromItem& = 0, Optional ByVal nToItem& = -1) As Long
    CountOf = gdCountOf(m.hArray, dMatchValue, nFromItem, nToItem)
End Function

' writes array to a text file
Public Function ToFile(ByVal strFileName$, _
        Optional ByVal bAppendToFile As Boolean = False, _
        Optional ByVal bReverseOrder As Boolean = False, _
        Optional ByVal bDeleteFileIfEmpty As Boolean = True) As Boolean
    
    Dim fh%, i&

    On Error GoTo ToFileError
    If m.hArray = 0 Then GoTo ToFileError
    
    If gdGetSize(m.hArray) > 0 Then
        fh = FreeFile
        If bAppendToFile Then
            Open strFileName For Append As #fh
        Else
            Open strFileName For Output As #fh
        End If
        If bReverseOrder Then
            For i = Size - 1 To 0 Step -1
                Print #fh, Item(i)
            Next
        Else
            For i = 0 To Size - 1
                Print #fh, Item(i)
            Next
        End If
        Close #fh
    ElseIf Not bAppendToFile Then
        If bDeleteFileIfEmpty Then
            On Error Resume Next
            Kill strFileName
        Else
            fh = FreeFile
            Open strFileName For Output As #fh
            Close #fh
        End If
    End If
    ToFile = True
    Exit Function

ToFileError:
    ToFile = False
    Exit Function
End Function

' Loads array from a text file
' - bAppendToArray: will append file to existing array
' - vUntilLine: if string, will read until it finds the string;
'       if numeric, will read up to that many lines
Public Function FromFile(ByVal strFileName$, _
        Optional ByVal bAppendToArray As Boolean = False, _
        Optional ByVal vUntilLine As Variant = "", _
        Optional ByVal strIgnoreIfLineStartsWith = "") As Boolean

    Dim fh%, iLine&, strTemp$, nLines&, nMaxLines&, strLastLine$
    Dim aLines As New cGdArray

    On Error GoTo FromFileError
    If m.hArray = 0 Then
        ' assume a string array
        Create eGDARRAY_Strings
    End If

    If Not bAppendToArray Then Size = 0
    
    If VarType(vUntilLine) = vbString Then
        strLastLine = vUntilLine
        nMaxLines = 999999999
    Else
        nMaxLines = Int(vUntilLine)
    End If
    
    If FileExist(strFileName) Then
        fh = FreeFile
        Open strFileName For Input As #fh
        Do While Not EOF(fh)
            Line Input #fh, strTemp
            ' TLB 6/8/2010: handle files with only a Chr(10) line delimiter (e.g. from Unix)
            If InStr(strTemp, Chr(10)) > 0 Then
                aLines.SplitFields strTemp, Chr(10)
            Else
                aLines.Size = 0
            End If
            For iLine = 0 To aLines.Size
                If aLines.Size > 0 Then
                    If iLine = aLines.Size Then Exit For
                    strTemp = aLines(iLine)
                End If
                If Len(strIgnoreIfLineStartsWith) = 0 Then
                    Add strTemp
                ElseIf (Len(strTemp) < Len(strIgnoreIfLineStartsWith)) Then
                    Add strTemp
                ElseIf (Left(strTemp, Len(strIgnoreIfLineStartsWith)) <> strIgnoreIfLineStartsWith) Then
                    Add strTemp
                End If
                nLines = nLines + 1
                If nLines >= nMaxLines Then Exit Do
                If Len(strLastLine) > 0 Then
                    If strLastLine = strTemp Then Exit Do
                End If
            Next
        Loop
        Close #fh
        FromFile = True
    End If
    Exit Function

FromFileError:
    FromFile = False
    Exit Function
End Function

' Performs a binary search on a SORTED array.
' - if FOUND: returns TRUE, iPos is position of match
' - if NOT FOUND: returns FALSE, iPos is position to insert at
'(ByVal hArray&, ByVal search_for#, iPos&, ByVal Bottom&, ByVal Top&)
Public Function BinarySearch(ByVal vSearchFor As Variant, Optional iMatchPos As Long, _
        Optional ByVal eSortFlags As eGdSort_Flags = eGdSort_Default, _
        Optional ByVal iFromItem& = 0, Optional ByVal iToItem& = 999999999) As Boolean
            
    Dim iPos As Long, bFound As Boolean, hString&
   
    If m.hArray <> 0 Then
        If m.eType <> eGDARRAY_Strings Then
            'For numeric arrays
            bFound = gdBinarySearch(m.hArray, CDbl(vSearchFor), iPos, eSortFlags, iFromItem, iToItem)
        Else
            'For a string array, convert string to search
            'for to a gdString and pass it's handle.
            hString = gdCreateArray(eGDARRAY_gdString)
            If hString <> 0 Then
                gdSetStr hString, 0, Str(vSearchFor)
                bFound = gdBinarySearch(m.hArray, hString, iPos, eSortFlags, iFromItem, iToItem)
                gdDestroyArray hString
            End If
        End If
    End If
    
    If Not IsMissing(iMatchPos) Then
        iMatchPos = iPos
    End If
    BinarySearch = bFound
End Function

Public Property Get ArrayType() As eGdArray_Type
    ArrayType = m.eType
End Property

'Returns a "copy" of this array.
'(Note: array data not actually copied yet -- both
' arrays are pointing to same data until one of them is
' edited, then that one gets its own copy of the data)
Public Function MakeCopy() As cGdArray
    'first create one of same type
    Dim aCopy As cGdArray
    Set aCopy = New cGdArray
    aCopy.Create m.eType
    'then make it be a "copy" of this one
    gdCopy aCopy.ArrayHandle, m.hArray
    Set MakeCopy = aCopy
End Function

'Makes current object a copy of the handle being passed.
Public Function CopyFromHandle(ByVal CopyFromArrayHandle As Long) As Boolean
    
    Destroy 'destroy current array if exists
    
    'make copy of array
    m.eType = gdGetType(CopyFromArrayHandle)
    m.hArray = gdCreateArray(m.eType) '(must start with a valid object)
    If gdCopy(m.hArray, CopyFromArrayHandle) <> 0 Then
        CopyFromHandle = True
    End If
    m.bAutoDestroy = True
    
End Function

Public Function IsNullAt(ByVal nOffset&) As Boolean
    Select Case m.eType
        Case eGDARRAY_NoArray:
            IsNullAt = True
        Case eGDARRAY_Strings:
            If nOffset >= gdGetSize(m.hArray) Or nOffset < 0 Then
                IsNullAt = True
            End If
        Case eGDARRAY_Objects:
            If Item(nOffset) Is Nothing Then
                IsNullAt = True
            End If
        Case Else:
            If gdGetNum(m.hArray, nOffset) = gdNullValue(m.hArray) Then
                IsNullAt = True
            End If
    End Select
End Function

'Reads/Writes array to a binary file
Public Function Serialize(ByVal vFile As Variant, ByVal bPut As Boolean)
    Dim bSuccess As Boolean, hFile As Long, TypeByte As Byte
    If bPut And (m.hArray = 0) Then Exit Function
    If VarType(vFile) <> vbString Then
        'file handle is being passed (file already open)
        hFile = CLng(vFile)
    ElseIf bPut Then
        'open file for binary output
        hFile = gdFileOpen(vFile, "wb")
    Else
        'open file for binary input
        hFile = gdFileOpen(vFile, "rb")
    End If
    If hFile <> 0 Then
        'read/write type of array
        TypeByte = m.eType
        If gdFileBinaryIO(hFile, TypeByte, 1, bPut) = 1 Then
            If Not bPut Then
                'when reading, make sure correct type
                If m.eType <> TypeByte Then
                    Create TypeByte
                End If
            End If
            'serialize the array
            bSuccess = gdSerializeArray(m.hArray, hFile, bPut)
        End If
        'close file (if opened here)
        If VarType(vFile) = vbString Then
            gdFileClose hFile
        End If
    End If
    'clear if unsuccessful read
    If Not bSuccess And Not bPut Then Destroy
    Serialize = bSuccess
End Function

' To join array items into a single delimited string
Public Function JoinFields(ByVal strDelim$) As String
    Dim hString&
    hString = gdJoinFields(m.hArray, strDelim)
    JoinFields = gdGetStr(hString)
    gdDestroyString hString
End Function

Public Sub MakeConstantArray(ByVal dValue#, Optional ByVal nSize& = -1)
    If m.hArray Then
        If nSize < 0 Then
            nSize = gdGetSize(m.hArray)
            'nSize = 0 clears Constant status, so make at least 1
            If nSize = 0 Then nSize = 1
        End If
        gdMakeConstantValue m.hArray, dValue, nSize
    End If
End Sub

Public Property Get IsConstantValue() As Boolean
    IsConstantValue = gdIsConstantValue(m.hArray)
End Property

' Moves one or more items up or down in the array.
' - nBlockStart: starting offset of block to move
' - nBlockSize: number of items to move (1 to move a single item)
' - nShiftAmount: number of positions to shift block up/down
'      (if > 0, moves toward end; if < 0, moves toward beginning)
' - returns true if performed a move
Public Function MoveItems(ByVal nBlockStart&, ByVal nBlockSize&, ByVal nShiftAmount&) As Boolean
    If m.hArray Then
        MoveItems = gdMoveItems(m.hArray, nBlockStart, nBlockSize, nShiftAmount)
    End If
End Function

' Calculates various statistics for the array
' - eStatType: sum, average, sum of squares, standard deviation, variance, etc.
' - for whole array: pass nFromItem = 0, nToItem = -1
Public Function CalcStatistic(ByVal eStatType As eGdStatisticType, _
        Optional ByVal nFromItem& = 0, Optional ByVal nToItem& = -1) As Double
    CalcStatistic = gdCalcStatistic(m.hArray, eStatType, nFromItem, nToItem)
End Function

' Calculates various moving statistics for the array (for a rolling "window")
' - eStatType: sum, average, sum of squares, standard deviation, variance, etc.
' - nPeriods: size of the rolling window, or 0 = from beginning
Public Function CalcMovingStatistic(ByVal eStatType As eGdStatisticType, Optional ByVal nPeriods# = -1) As cGdArray
    Dim aResults As New cGdArray
    aResults.Create eGDARRAY_Doubles
    gdCalcMovingStatistic aResults.ArrayHandle, m.hArray, eStatType, nPeriods
    Set CalcMovingStatistic = aResults
End Function

' Fills string array with all matching files
' - strFileSpec: can include wildcards and special search options
'   (e.g. /s for subdirs, /n=-3 for newer than 3 days ago, see
'    help for "Files32" program for description of search options)
' - bIncludeFullPath: true to include the full path in the filenames
' - bIncludeFolderNames: true to include folder names in the array
' - bIncludeExtraFields: true to append tab-delimited fields (name, size, date.time, attribs)
' - dTotalBytes: optional byref to pass back total bytes of matching files
Public Function GetMatchingFiles(ByVal strFileSpec$, _
        Optional ByVal bIncludeFullPath As Boolean = True, _
        Optional ByVal bIncludeFolderNames As Boolean = False, _
        Optional ByVal bIncludeExtraFields As Boolean = False, _
        Optional dTotalBytes As Double = 0) As Long

    Dim MatchedTotals As gdFileMatchingTotals
    Dim dwFlags As Long

    If bIncludeFullPath Then dwFlags = dwFlags Or &H1
    If bIncludeFolderNames Then dwFlags = dwFlags Or &H2
    If bIncludeExtraFields Then dwFlags = dwFlags Or &H4

    If m.hArray = 0 Or m.eType <> eGDARRAY_Strings Then
        Create eGDARRAY_Strings
    Else
        Clear
    End If
   
    GetMatchingFiles = gdGetMatchingFiles2(strFileSpec, dwFlags, m.hArray, MatchedTotals)
    dTotalBytes = MatchedTotals.dMatchedBytes

End Function

' Appends items from another cGdArray (data types must match).
' - returns true if one or more items were appended
Public Function AppendFromArray(FromArray As cGdArray, Optional ByVal iFromItem& = 0, Optional ByVal iToItem& = -1) As Boolean

    AppendFromArray = gdAppendFrom(m.hArray, FromArray.ArrayHandle, iFromItem, iToItem)

End Function

' To create a sorted index on this array.
' - this routine does NOT sort this array, rather it returns an array of index values
'       which will reference this array in a sorted order -- you can then use these
'       index values to walk through this array in a sorted order
' - for example, here's how to walk through this array in a descending fashion:
'       Set aIndex = ThisArray.CreateSortedIndex(eGdSort_Descending)
'       For i = 0 to aIndex.Size - 1
'           vNextItem = ThisArray(aIndex(i))
'       Next
Public Function CreateSortedIndex(Optional ByVal eSortFlags As eGdSort_Flags = eGdSort_Default, _
            Optional ByVal nFirstItem As Long = 0, Optional ByVal nLastItem As Long = -1) As cGdArray

    Dim aIndex As New cGdArray
    
    aIndex.Create eGDARRAY_Longs, Me.Size
    gdSortAsIndex aIndex.ArrayHandle, m.hArray, 1, eSortFlags, nFirstItem, nLastItem

    Set CreateSortedIndex = aIndex
End Function

Public Property Get SharedDataMode() As Boolean
    SharedDataMode = gdGetShared(m.hArray)
End Property

Public Property Let SharedDataMode(ByVal bSharedData As Boolean)
    gdSetShared m.hArray, bSharedData
End Property
