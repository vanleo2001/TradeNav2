VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMemBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class allows for managing a memory buffer (memory
'which can be preallocated and appended to through direct
'memory copies) using low-level API memory functions:
'- MUCH faster than constantly appending to a VB string
'- Auto-allocates memory whenever necessary and can
'   use auto-chunk feature for greater efficiency
'- Get/Put functions make it easy to work with a block
'   of memory containing strings and binary numbers
Option Explicit

Private Type mPrivate
    nPtr As Long        'ptr to allocated memory
    nAllocated As Long  '# bytes allocated
    nLength As Long     '# bytes currently "used"
    nNextGetOffset As Long 'AtOffset default used by Get functions
    bAutoChunk As Boolean 'true if want to double allocations
    hHeap As Long       'handle to process heap
End Type
Private m As mPrivate

Private Sub Class_Initialize()
    m.hHeap = GetProcessHeap '(used by API memory functions)
    m.bAutoChunk = True 'default
End Sub

Private Sub Class_Terminate()
    Clear
End Sub

'AutoChunk: true if want to double allocations
' whenever need to allocate more memory for buffer
Public Property Get AutoChunk() As Boolean
    AutoChunk = m.bAutoChunk
End Property
Public Property Let AutoChunk(ByVal bAutoChunk As Boolean)
    m.bAutoChunk = bAutoChunk
End Property

'Free memory and init members
Public Sub Clear()
    If m.nPtr <> 0 Then
        HeapFree m.hHeap, 0, ByVal m.nPtr
        m.nPtr = 0
    End If
    m.nAllocated = 0
    m.nLength = 0
    m.nNextGetOffset = 0
End Sub

'(verify length is not bigger than what is currently allocated)
Private Sub CheckLength()
    If m.nLength > m.nAllocated Then m.nLength = m.nAllocated
    If m.nNextGetOffset >= m.nLength Then m.nNextGetOffset = 0
End Sub

'Returns the memory address of buffer
Public Property Get MemPtr() As Long
    MemPtr = m.nPtr
End Property

'Returns # bytes currently allocated
Public Property Get Allocated() As Long
    Allocated = m.nAllocated
End Property

'Allocates # bytes for memory buffer
Public Property Let Allocated(ByVal nNewAllocated As Long)
    If nNewAllocated <= 0 Then
        Clear '(free memory)
    ElseIf nNewAllocated <> m.nAllocated Then
        If m.nPtr = 0 Then
            'need a new allocation
            m.nPtr = HeapAlloc(m.hHeap, 0, nNewAllocated)
        Else
            'resize (will preserve existing data)
            m.nPtr = HeapReAlloc(m.hHeap, 0, ByVal m.nPtr, nNewAllocated)
        End If
        If m.nPtr = 0 Then
            m.nAllocated = 0
        Else
            'should request size since # bytes allocated may be
            'slightly different (rounded up) than # requested
            m.nAllocated = HeapSize(m.hHeap, 0, ByVal m.nPtr)
        End If
        CheckLength
    End If
End Property

'Returns length of buffer (# bytes currently "used")
Public Property Get Length() As Long
    Length = m.nLength
End Property

'Sets length of buffer (will allocate more if necessary)
Public Property Let Length(ByVal nNewLength As Long)
    If nNewLength < 0 Then
        nNewLength = 0
    ElseIf nNewLength > m.nAllocated Then
        'need to allocate more
        Allocated = nNewLength
    End If
    m.nLength = nNewLength
    CheckLength
End Property

'Frees extra memory (if allocated > length)
Public Sub FreeExtra()
    Allocated = m.nLength
End Sub

'Returns memory buffer (as array of bytes)
Public Property Get Bytes() As Byte()
    Dim aBytes() As Byte
    If m.nPtr <> 0 And m.nLength > 0 Then
        ReDim aBytes(0 To m.nLength - 1) As Byte
        CopyMemory aBytes(0), ByVal m.nPtr, m.nLength
    Else
        ' send back an "empty" array (set UBound < 0)
        ReDim aBytes(-1 To -1) As Byte
        aBytes(-1) = 0
    End If
    Bytes = aBytes
End Property

'Sets memory buffer
Public Property Let Bytes(aBytes() As Byte)
    Dim nLength&
    ' check for an empty array (if UBound < 0, or if UBound fails)
    On Error GoTo EmptyArray
    If UBound(aBytes) < 0 Then
        Clear
    Else
        nLength = UBound(aBytes) - LBound(aBytes) + 1
        Length = nLength
        If m.nLength = nLength And m.nLength > 0 Then
            CopyMemory ByVal m.nPtr, aBytes(LBound(aBytes)), m.nLength
        End If
    End If
    Exit Property
EmptyArray:
    Clear
    Exit Property
End Property

'Returns memory buffer (as VB string)
Public Property Get Buffer() As String
    Dim s As String
    If m.nPtr <> 0 And m.nLength > 0 Then
        s = Space(m.nLength)
        If IsDBCS Then
            Dim aBytes() As Byte, i&
            ReDim aBytes(1 To m.nLength) As Byte
            CopyMemory aBytes(1), ByVal m.nPtr, m.nLength
            For i = 1 To m.nLength
                Mid(s, i, 1) = ChrB(aBytes(i))
            Next
        Else
            '(VB will automatically convert the bytes to a Unicode string)
            CopyMemory ByVal s, ByVal m.nPtr, m.nLength
        End If
    End If
    Buffer = s
End Property

'Sets memory buffer
Public Property Let Buffer(ByVal strNewBuffer As String)
    Length = Len(strNewBuffer)
    If m.nLength = Len(strNewBuffer) And m.nLength > 0 Then
        If IsDBCS Then
            Dim aBytes() As Byte, i&
            ReDim aBytes(1 To m.nLength) As Byte
            For i = 1 To m.nLength
                aBytes(i) = AscB(Mid(strNewBuffer, i, 1))
            Next
            CopyMemory ByVal m.nPtr, aBytes(1), m.nLength
        Else
            '(VB will automatically convert from a Unicode string to bytes)
            CopyMemory ByVal m.nPtr, ByVal strNewBuffer, m.nLength
        End If
    End If
End Property

'Prepares buffer for insert/append and
'returns address of where to copy to.
Private Function DestPtr(ByVal nCopyLength As Long, _
        ByVal nAtOffset As Long, _
        Optional ByVal bInsert As Boolean = False) As Long

    Dim nDest&, nMoveTo&, nMoveLen&, nNewLength&, nAlloc&
    'see if appending
    If nAtOffset = -1 Or nAtOffset > m.nLength Then
        nAtOffset = m.nLength
    End If
    'make sure things are valid
    If nCopyLength > 0 And nAtOffset >= 0 Then
        'determine what new length will be
        If bInsert Then
            nNewLength = m.nLength + nCopyLength
        Else
            nNewLength = nAtOffset + nCopyLength
            '(but can't be less than current)
            If nNewLength < m.nLength Then nNewLength = m.nLength
        End If
        'allocate more memory if need to
        nAlloc = m.nAllocated
        Do While nAlloc < nNewLength
            If Not m.bAutoChunk Then
                'allocate just enough
                nAlloc = nNewLength
            ElseIf nAlloc < 256 Then
                'minimum alloc when chunking
                nAlloc = 512
            Else
                'chunk allocations up by multiples of 2
                nAlloc = nAlloc * 2
            End If
        Loop
        Allocated = nAlloc
        'make sure allocation worked
        If nNewLength <= m.nAllocated Then
            If nAtOffset >= m.nLength Then
                'just appending to end
                nDest = m.nPtr + m.nLength
            ElseIf bInsert Then
                'if inserting into middle of buffer,
                'need to first move some existing memory
                'before overwriting at that location
                nMoveLen = m.nLength - nAtOffset
                nDest = m.nPtr + nAtOffset
                nMoveTo = nDest + nCopyLength
                MoveMemory ByVal nMoveTo, ByVal nDest, nMoveLen
            Else
                'overwrite at offset
                nDest = m.nPtr + nAtOffset
            End If
            m.nLength = nNewLength
        End If
    End If
    
    DestPtr = nDest
End Function


'PUT/INSERT FUNCTIONS ...
'- used to write or append into the memory buffer
'   at a specified offset (zero-based),
'   appends if AtOffset = -1
'- calls "DestPtr" to prepare buffer for copy
'- performs a direct memory copy
'- returns true if successful

'Inserts/appends a string into the memory buffer
'(if nAtOffset = -1, then appends)
Public Function Insert(ByVal strInsertToBuffer As String, _
        Optional ByVal nAtOffset As Long = -1) As Boolean

    Insert = PutStr(strInsertToBuffer, nAtOffset)

End Function

'Writes/appends a string into the memory buffer
Public Function PutStr(strToBuffer As String, _
        Optional ByVal nAtOffset As Long = -1) As Boolean

    Dim nCopyTo&, nCopyLength&
    nCopyLength = Len(strToBuffer)
    nCopyTo = DestPtr(nCopyLength, nAtOffset)
    If nCopyTo <> 0 Then
        If IsDBCS Then
            Dim aBytes() As Byte, i&
            ReDim aBytes(1 To nCopyLength) As Byte
            For i = 1 To nCopyLength
                aBytes(i) = AscB(Mid(strToBuffer, i, 1))
            Next
            CopyMemory ByVal nCopyTo, aBytes(1), nCopyLength
        Else
            CopyMemory ByVal nCopyTo, ByVal strToBuffer, nCopyLength
        End If
        PutStr = True
    End If
    
End Function

'Writes/appends bytes into the memory buffer
Public Function PutBytes(aBytes() As Byte, _
        Optional ByVal nAtOffset As Long = -1) As Boolean

    Dim nCopyTo&, nCopyLength&
    ' check for an empty array (if UBound fails or is < 0)
    On Error Resume Next
    nCopyLength = -1
    nCopyLength = UBound(aBytes) ' - LBound(aBytes) + 1
    If nCopyLength >= 0 Then
        nCopyLength = UBound(aBytes) - LBound(aBytes) + 1
        nCopyTo = DestPtr(nCopyLength, nAtOffset)
        If nCopyTo <> 0 And nCopyLength > 0 Then
            CopyMemory ByVal nCopyTo, aBytes(LBound(aBytes)), nCopyLength
            PutBytes = True
        End If
    End If
    
End Function

'Writes/appends from memory into the memory buffer
Public Function PutFromMemory(ByVal nCopyFrom&, ByVal nCopyLength&, _
        Optional ByVal nAtOffset As Long = -1) As Boolean

    Dim nCopyTo&
    nCopyTo = DestPtr(nCopyLength, nAtOffset)
    If nCopyTo <> 0 Then
        CopyMemory ByVal nCopyTo, ByVal nCopyFrom, nCopyLength
        PutFromMemory = True
    End If
    
End Function

'Writes/appends a Byte or Boolean into the memory buffer
Public Function PutByte(ByVal Number As Byte, _
        Optional ByVal nAtOffset As Long = -1) As Boolean
    PutByte = PutFromMemory(GetAddress(Number), 1, nAtOffset)
End Function

'Writes/appends an Int into the memory buffer
Public Function PutInt(ByVal Number As Integer, _
        Optional ByVal nAtOffset As Long = -1) As Boolean
    PutInt = PutFromMemory(GetAddress(Number), 2, nAtOffset)
End Function

'Writes/appends a Long into the memory buffer
Public Function PutLong(ByVal Number As Long, _
        Optional ByVal nAtOffset As Long = -1) As Boolean
    PutLong = PutFromMemory(GetAddress(Number), 4, nAtOffset)
End Function

'Writes/appends a Single into the memory buffer
Public Function PutSingle(ByVal Number As Single, _
        Optional ByVal nAtOffset As Long = -1) As Boolean
    PutSingle = PutFromMemory(GetAddress(Number), 4, nAtOffset)
End Function

'Writes/appends a Double into the memory buffer
Public Function PutDouble(ByVal Number As Double, _
        Optional ByVal nAtOffset As Long = -1) As Boolean
    PutDouble = PutFromMemory(GetAddress(Number), 8, nAtOffset)
End Function


'GET functions ...
'- default for AtOffset is -1, which means
'   get from position after the last Get call
'- calls "SourcePtr" to verify ability to get from memory

'(called by other Get functions)
Private Function SourcePtr(ByVal nAtOffset&, ByVal nBytes&) As Long
    If nAtOffset = -1 Then nAtOffset = m.nNextGetOffset
    If m.nPtr <> 0 And nAtOffset >= 0 And nBytes > 0 _
            And nAtOffset + nBytes <= m.nLength Then
        'update m.nNextGetOffset for next Get call
        m.nNextGetOffset = nAtOffset + nBytes
        'return good address
        SourcePtr = m.nPtr + nAtOffset
    Else
        SourcePtr = 0 '(unavailable address)
    End If
End Function

'Returns part of the memory buffer
'(if nLengthToGet = -1, returns to end of buffer)
Public Function GetStr(Optional ByVal nLengthToGet& = -1, _
        Optional ByVal nAtOffset& = -1) As String
    
    Dim strGet As String, nPtr As Long
    
    'fix length to get
    If nAtOffset = -1 Then nAtOffset = m.nNextGetOffset
    If nLengthToGet = -1 Or nLengthToGet > m.nLength - nAtOffset Then
        nLengthToGet = m.nLength - nAtOffset
    End If
    'get address of source
    nPtr = SourcePtr(nAtOffset, nLengthToGet)
    If nPtr <> 0 Then
        strGet = Space(nLengthToGet)
        If IsDBCS Then
            Dim aBytes() As Byte, i&
            ReDim aBytes(1 To nLengthToGet) As Byte
            CopyMemory aBytes(1), ByVal nPtr, nLengthToGet
            For i = 1 To nLengthToGet
                Mid(strGet, i, 1) = ChrB(aBytes(i))
            Next
        Else
            CopyMemory ByVal strGet, ByVal nPtr, nLengthToGet
        End If
    End If
    GetStr = strGet
    
End Function

'Returns part of the memory buffer
'(if nLengthToGet = -1, returns to end of the buffer)
Public Function GetBytes(Optional ByVal nLengthToGet& = -1, _
        Optional ByVal nAtOffset& = -1) As Byte()
    
    Dim aBytes() As Byte, nPtr As Long
    
    'fix length to get
    If nAtOffset = -1 Then nAtOffset = m.nNextGetOffset
    If nLengthToGet = -1 Or nLengthToGet > m.nLength - nAtOffset Then
        nLengthToGet = m.nLength - nAtOffset
    End If
    'get address of source
    nPtr = SourcePtr(nAtOffset, nLengthToGet)
    If nPtr <> 0 And nLengthToGet > 0 Then
        ReDim aBytes(0 To nLengthToGet - 1) As Byte
        CopyMemory aBytes(0), ByVal nPtr, nLengthToGet
    Else
        ' send back an "empty" array (set UBound < 0)
        ReDim aBytes(-1 To -1) As Byte
        aBytes(-1) = 0
    End If
    GetBytes = aBytes
    
End Function

'Retrieves a Byte or Boolean from memory buffer at specified offset
Public Function GetByte(Optional ByVal nAtOffset& = -1) As Byte
    Dim v As Byte, nPtr&
    nPtr = SourcePtr(nAtOffset, 1)
    If nPtr <> 0 Then CopyMemory v, ByVal nPtr, 1
    GetByte = v
End Function

'Retrieves an Int from memory buffer at specified offset
Public Function GetInt(Optional ByVal nAtOffset& = -1) As Integer
    Dim v As Integer, nPtr&
    nPtr = SourcePtr(nAtOffset, 2)
    If nPtr <> 0 Then CopyMemory v, ByVal nPtr, 2
    GetInt = v
End Function

'Retrieves a Long from memory buffer at specified offset
Public Function GetLong(Optional ByVal nAtOffset& = -1) As Long
    Dim v As Long, nPtr&
    nPtr = SourcePtr(nAtOffset, 4)
    If nPtr <> 0 Then CopyMemory v, ByVal nPtr, 4
    GetLong = v
End Function

'Retrieves a Single from memory buffer at specified offset
Public Function GetSingle(Optional ByVal nAtOffset& = -1) As Single
    Dim v As Single, nPtr&
    nPtr = SourcePtr(nAtOffset, 4)
    If nPtr <> 0 Then CopyMemory v, ByVal nPtr, 4
    GetSingle = v
End Function

'Retrieves a Double from memory buffer at specified offset
Public Function GetDouble(Optional ByVal nAtOffset& = -1) As Double
    Dim v As Double, nPtr&
    nPtr = SourcePtr(nAtOffset, 8)
    If nPtr <> 0 Then CopyMemory v, ByVal nPtr, 8
    GetDouble = v
End Function

Public Property Get NextGetOffset() As Long
    NextGetOffset = m.nNextGetOffset
End Property

Public Property Let NextGetOffset(ByVal nNextGetOffset As Long)
    m.nNextGetOffset = nNextGetOffset
    If m.nNextGetOffset < 0 Then m.nNextGetOffset = 0
    CheckLength
End Property

' Reads a binary file into the memory buffer (can append to the existing buffer)
Public Function FromFile(ByVal strFileName As String, Optional ByVal bAppend As Boolean = False) As Boolean

    Dim hFile&, nFileLen&, nStartAt&
    
    If Not bAppend Then Clear
    nFileLen = FileLength(strFileName)
    If nFileLen > 0 Then
        hFile = FileOpen(strFileName, "rb")
        If hFile Then
            nStartAt = Length
            Length = Length + nFileLen
            nStartAt = MemPtr + nStartAt
            If FileBinaryIO(hFile, ByVal nStartAt, nFileLen, False) <> 0 Then
                FromFile = True
            End If
            FileClose hFile
        End If
    End If

End Function

' Writes the entire memory buffer to a binary file (can append to an existing file)
Public Function ToFile(ByVal strFileName As String, Optional ByVal bAppend As Boolean = False) As Boolean

    Dim hFile&
    
    If bAppend Then
        hFile = FileOpen(strFileName, "a+b")
    Else
        hFile = FileOpen(strFileName, "w+b")
    End If
    If hFile Then
        If FileBinaryIO(hFile, ByVal MemPtr, Length, True) <> 0 Then
            ToFile = True
        End If
        FileClose hFile
    End If

End Function

