VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSymbolPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"gdArray"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' This class loads and manages all relevant symbols
' and their associated Filters, Criterias, SymbolGroups, etc.
Option Explicit

Enum eSYM_SecType           'in database:
    eSYMType_Unknown = 0    'U
    eSYMType_Index = 1      'I
    eSYMType_Stock = 2      'S
    eSYMType_Future = 3     'F
    eSYMType_MutualFund = 4 'M
    eSYMType_Forex = 5      'X
End Enum

' Field #'s for each symbol info array
Const kSymIndex = 0 '(array of record #'s sorted by SymbolID)
Const kDbRecNum = 1
Const kSymbol = 2
Const kSymbolID = 3
Const kDesc = 4
Const kFirstDate = 5
Const kLastDate = 6
Const kSecType = 7
Const kAllSymbols = 8
Const kAllFutures = 9
Const kAllForex = 10
Const kAllIndices = 11
Const kAllStocks = 12
Const kAllIntlStocks = 13
Const kAllMutFunds = 14
Const kOtherFields = 15

' Private data
Private Type mPrivate
    'Collections to hold all the objects in the pool
    SymbolGroups As cGdTree
    Criterias As cGdTree
    Filters As cGdTree
    
    'Table to hold all the data arrays (one for each
    ' symbol info, two for each numeric Criteria, one
    ' for all the other objects)
    ArrayTable As cGdTable

    'Array to associate an array in the table with an object
    '- format is TYPE:ObjectID
    '- types are: INF, GRP, FIL, DSV (Criteria values), DSP (Criteria percentiles)
    aFieldIDs As cGdArray

    'other members
    'bDirtyCriteria As Boolean

End Type
Private m As mPrivate

Private Sub Class_Initialize()
    
    Set m.SymbolGroups = New cGdTree
    Set m.Criterias = New cGdTree
    Set m.Filters = New cGdTree
    Set m.ArrayTable = New cGdTable
    Set m.aFieldIDs = New cGdArray
    m.aFieldIDs.Create eGDARRAY_Strings
    
    Clear

End Sub

Private Sub Class_Terminate()

    Dim i&
    ' write arrays to disk
    Serialize True
    Clear
    
    Set m.aFieldIDs = Nothing
    Set m.ArrayTable = Nothing
    Set m.SymbolGroups = Nothing
    Set m.Criterias = Nothing
    Set m.Filters = Nothing

End Sub

Private Sub Clear()
On Error GoTo ErrSection:

    m.aFieldIDs.Clear
    m.ArrayTable.Clear
    m.SymbolGroups.Clear
    m.Criterias.Clear
    m.Filters.Clear
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolPool.Clear", eGDRaiseError_Raise
    
End Sub

Public Property Get DirtyCriteria() As Boolean

    Dim lIndex As Long, bIsDirty As Boolean
    
    bIsDirty = False
    For lIndex = 1 To Criterias.Count
        If Criterias(lIndex).IsDirty And Criterias(lIndex).IsActive Then
            bIsDirty = True
            Exit For
        End If
    Next lIndex
    
    DirtyCriteria = bIsDirty
End Property

Public Property Let DirtyCriteria(ByVal bDirty As Boolean)

    Dim lIndex As Long
    
    For lIndex = 1 To Criterias.Count
        Criterias(lIndex).IsDirty = bDirty
        Criterias(lIndex).ToFile True
    Next lIndex

End Property

Public Property Get SymbolGroups() As cGdTree
    Set SymbolGroups = m.SymbolGroups
End Property

Public Property Get Criterias() As cGdTree
    Set Criterias = m.Criterias
End Property

Public Property Get Filters() As cGdTree
    Set Filters = m.Filters
End Property

' Returns the object from the pool's collections for the
' FieldID (e.g. "GRP:Dow30") -- returns Nothing if not found
Public Function PoolObject(ByVal strFieldID$) As Object
    Dim strObjKey$
    strObjKey = UCase(Mid(strFieldID, 5))
    Set PoolObject = Nothing
    Select Case UCase(Left(strFieldID, 3))
        Case "GRP":
            Set PoolObject = m.SymbolGroups(strObjKey)
        Case "FIL":
            Set PoolObject = m.Filters(strObjKey)
        Case "DSP", "DSV", "SCN":
            Set PoolObject = m.Criterias(strObjKey)
    End Select
End Function

Public Property Get ArrayTable() As cGdTable
    Set ArrayTable = m.ArrayTable
End Property

Public Property Get FieldID(ByVal nField&) As String
    FieldID = m.aFieldIDs(nField)
End Property

Public Property Let FieldID(ByVal nField&, ByVal strFieldID$)
    m.aFieldIDs(nField) = UCase(Trim(strFieldID))
End Property

Public Property Get FieldNumForID(ByVal strFieldID$) As Long
    Dim nFld&, i&
    strFieldID = Trim(UCase(strFieldID))
    nFld = -1
    For i = 0 To m.aFieldIDs.Size - 1
        If UCase(m.aFieldIDs(i)) = strFieldID Then
            nFld = i
            Exit For
        End If
    Next
    FieldNumForID = nFld
End Property

Public Property Get NumRecords() As Long
    NumRecords = m.ArrayTable.NumRecords
End Property

'Public Property Get NumFields() As Long
'    NumFields = m.ArrayTable.NumFields
'End Property

Public Function PoolRecForSymbolID(ByVal nSymbolID&) As Long
On Error GoTo ErrSection:

    Dim iLo&, iHi&, iMid&, nFoundRec&, nCheck&, nRec&
    
    nFoundRec = -1
    iLo = 0
    iHi = m.ArrayTable.NumRecords - 1
    Do While iLo <= iHi
        iMid = iLo + (iHi - iLo) \ 2  ' so sum does not overflow
        nRec = m.ArrayTable.Num(kSymIndex, iMid)
        nCheck = m.ArrayTable.Num(kSymbolID, nRec)
        If nSymbolID < nCheck Then
            iHi = iMid - 1  ' in lower half
        ElseIf nSymbolID > nCheck Then
            iLo = iMid + 1  ' in upper half
            iMid = iLo  ' position to insert
        Else
            nFoundRec = nRec
            Exit Do
        End If
    Loop
    PoolRecForSymbolID = nFoundRec

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.PoolRecForSymbolID", eGDRaiseError_Raise
    
End Function

Public Function PoolRecForSymbol(ByVal strSymbol$, _
        Optional ByVal bExactMatch As Boolean = True) As Long
On Error GoTo ErrSection:

    Dim iLo&, iHi&, iMid&, nFoundRec&, strCheck$
    
    nFoundRec = -1
    strSymbol = UCase(Trim(strSymbol))
    iLo = 0
    iHi = m.ArrayTable.NumRecords - 1
    Do While iLo <= iHi
        iMid = iLo + (iHi - iLo) \ 2  ' so sum does not overflow
        strCheck = UCase(Trim(Symbol(iMid)))
        If strSymbol < strCheck Then
            iHi = iMid - 1  ' in lower half
        ElseIf strSymbol > strCheck Then
            iLo = iMid + 1  ' in upper half
            iMid = iLo  ' position to insert
        Else
            nFoundRec = iMid
            Exit Do
        End If
    Loop
    If Not bExactMatch Then nFoundRec = iMid
    PoolRecForSymbol = nFoundRec
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.PoolRecForSymbol", eGDRaiseError_Raise
    
End Function

Public Function PoolRecForDesc(ByVal strDesc$, _
        Optional ByVal bExactMatch As Boolean = True) As Long
On Error GoTo ErrSection:

    Dim iLo&, iHi&, iMid&, nFoundRec&, strCheck$, nRec&
    Dim aSorted As New cGdArray
    
    ' first build a sorted index on the description array
    aSorted.Create eGDARRAY_Longs
    gdSortAsIndex aSorted.ArrayHandle, m.ArrayTable.FieldArrayHandle(kDesc), True, eGdSort_IgnoreCase, 0, -1
    
    ' now do binary search on sorted index
    nFoundRec = -1
    strDesc = UCase(Trim(strDesc))
    iLo = 0
    iHi = aSorted.Size - 1
    Do While iLo <= iHi
        iMid = iLo + (iHi - iLo) \ 2  ' so sum does not overflow
        nRec = aSorted.Num(iMid)
        strCheck = UCase(Trim(Desc(nRec)))
        If strDesc < strCheck Then
            iHi = iMid - 1  ' in lower half
        ElseIf strDesc > strCheck Then
            iLo = iMid + 1  ' in upper half
            iMid = iLo  ' position to insert
        Else
            nFoundRec = nRec
            Exit Do
        End If
    Loop
    If Not bExactMatch Then nFoundRec = nRec
    
    PoolRecForDesc = nFoundRec
    
ErrExit:
    Set aSorted = Nothing
    Exit Function
    
ErrSection:
    Set aSorted = Nothing
    RaiseError "cSymbolPool.PoolRecForDesc", eGDRaiseError_Raise
    
End Function

' Returns an item of the db as displayed.
Public Property Get DataItem(ByVal nField&, ByVal nRecord&, _
        Optional ByVal strFormat$ = "") As Variant
    
    Dim i&, strType$, dNum#
    Dim Criteria As cCriteria
    
    Select Case nField
        Case kSymbol:
            DataItem = Symbol(nRecord)
        Case kDesc:
            DataItem = Desc(nRecord)
        Case kFirstDate, kLastDate:
            If nField = kFirstDate Then
                i = FirstDate(nRecord)
            Else
                i = LastDate(nRecord)
            End If
            If i > 0 Then
                DataItem = CDate(i)
            Else
                DataItem = ""
            End If
        Case kSecType:
            Select Case SecType(nRecord)
                Case eSYMType_Stock:
                    DataItem = "Stk"
                Case eSYMType_Future:
                    DataItem = "Fut"
                Case eSYMType_Index:
                    DataItem = "Idx"
                Case eSYMType_MutualFund:
                    DataItem = "MF"
                Case eSYMType_Forex:
                    DataItem = "Frx"
                Case Else:
                    DataItem = ""
            End Select
        Case Else:
            dNum = m.ArrayTable.Num(nField, nRecord)
            If dNum < -1 Then
                ' check for null value or no array
                If m.ArrayTable.IsNullAt(nField, nRecord) Then
                    DataItem = ""
                    Exit Property
                End If
            End If
            If m.ArrayTable.FieldType(nField) = eGDARRAY_TinyInts Then
                If Left(m.aFieldIDs(nField), 4) <> "DSP:" Then
                    If dNum = 0 Then
                        DataItem = "False"
                    Else
                        DataItem = "TRUE"
                    End If
                Else
                    DataItem = CStr(dNum)
                End If
            ElseIf Left(m.aFieldIDs(nField), 4) = "DSV:" Then
                Set Criteria = Criterias(Right(m.aFieldIDs(nField), Len(m.aFieldIDs(nField)) - 4))
                If InStr(UCase(Criteria.Name), "DATE") > 0 Then
                    DataItem = DateFormat(dNum)
                Else
                    Select Case Criteria.PriceDisplay
                        Case eCriteria_AutoRound
                            If dNum = 0 Then
                                DataItem = "0"
                            ElseIf Abs(dNum) > 10000 Or Int(dNum) = dNum Then
                                DataItem = Format(dNum, "0")
                            ElseIf Abs(dNum) < 10 Then
                                DataItem = Format(dNum, "0.####")
                            Else
                                DataItem = Format(dNum, "0.##")
                            End If
    
                        Case eCriteria_RoundToDecimal
                            DataItem = Trim(NumStr(dNum, 0, Criteria.DecimalPlaces))
    
                        Case eCriteria_TradingUnits
                            DataItem = PriceDisplay(dNum, SymbolID(nRecord))
                    End Select
                End If
            Else
                If dNum = 0 Then
                    DataItem = "0"
                ElseIf Abs(dNum) > 10000 Or Int(dNum) = dNum Then
                    DataItem = Format(dNum, "0")
                ElseIf Abs(dNum) < 10 Then
                    DataItem = Format(dNum, "0.####")
                Else
                    DataItem = Format(dNum, "0.##")
                End If
            End If
    End Select
    
    Set Criteria = Nothing
    
End Property

Public Property Get Symbol(ByVal nRecord&) As String
    Dim nRec&
    If nRecord >= 0 Then
        nRec = m.ArrayTable.Num(kDbRecNum, nRecord)
        If GotoRec(g.Universe.tblSymbols, nRec) Then
            Symbol = Trim(f4str(g.Universe.fldSymbol))
        End If
    End If
End Property

Public Property Get Desc(ByVal nRecord&) As String
    
    Dim nRec&
    Dim Info As vbSymbolInfo
    
    If nRecord >= 0 Then
If 0 Then
        nRec = m.ArrayTable.Num(kDbRecNum, nRecord)
        If GotoRec(g.Universe.tblSymbols, nRec) Then
            Desc = Trim(f4str(g.Universe.fldDesc))
        End If
Else
        If SU_GetSymbolInf(SymbolID(nRecord), Info) Then
            If Len(Info.Description) > 0 Then
                If Info.Exchange = "IDX" Then
                    Desc = Info.Description
                Else
                    Desc = Info.Description & " @ " & Info.Exchange '& ">"
                End If
            End If
        End If
End If
    End If
    
End Property

Public Property Get SymbolID(ByVal nRecord&) As Long
    SymbolID = m.ArrayTable.Num(kSymbolID, nRecord)
End Property

Public Function SymbolIDforSymbol(ByVal strSymbol$) As Long
    Dim nRecord&
    nRecord = PoolRecForSymbol(strSymbol, True)
    If nRecord < 0 Then
        SymbolIDforSymbol = 0
    Else
        SymbolIDforSymbol = m.ArrayTable.Num(kSymbolID, nRecord)
    End If
End Function

Public Function SymbolForID(ByVal nSymbolID&) As String
    Dim nRecord&
    nRecord = PoolRecForSymbolID(nSymbolID)
    If nRecord < 0 Then
        SymbolForID = ""
    Else
        SymbolForID = Symbol(nRecord)
    End If
End Function

Public Property Get DbRecNum(ByVal nRecord&) As Long
    DbRecNum = m.ArrayTable.Num(kDbRecNum, nRecord)
End Property

Public Property Get FirstDate(ByVal nRecord&) As Long
    FirstDate = m.ArrayTable.Num(kFirstDate, nRecord)
End Property

Public Property Get LastDate(ByVal nRecord&) As Long
    LastDate = m.ArrayTable.Num(kLastDate, nRecord)
End Property

Public Property Get SecType(ByVal nRecord&) As eSYM_SecType
    SecType = m.ArrayTable.Num(kSecType, nRecord)
End Property

Public Sub Load(Optional ByVal bTrySerialize = True)
On Error GoTo ErrSection:

    Dim rc&, i&, nCount&
    Dim bDontReloadIfSameCount As Boolean
    Dim bRecalc As Boolean
    
    ' If the pool already exists, save the Flag Group first
    If NumRecords > 0 Then SaveFlagGroup
    
    i = GetIniFileProperty("SymPoolOtherFields", 0, "", g.strIniFile)
    If i <> kOtherFields Then
        bTrySerialize = False
        SetIniFileProperty "SymPoolOtherFields", kOtherFields, "", g.strIniFile
    End If
    
    ' try to read serialized ArrayTable
    If bTrySerialize Then
        BenchMark
        If Serialize(False) Then
            If HasModule("F") = False Then
                bDontReloadIfSameCount = True
            ElseIf (SymbolIDforSymbol("AP-067") = 0) And (SfeAllowed = True) Then
                bDontReloadIfSameCount = False
            Else
                bDontReloadIfSameCount = True
            End If
            bTrySerialize = False
        End If
        'BenchMark "DataTable read: " & Str(bDontReloadIfSameCount)
    Else
        KillFile AddSlash(App.Path) & "SYMPOOL.MEM"
    End If
   
    ' reset last date of daily downloaded data
    LastDailyDownload True
   
    ' clear roll symbol cache
    RollSymbolForDate "", -999
   
    ' load symbols (unless serialized and same count)
    If LoadSymbols(bDontReloadIfSameCount) Then
        bRecalc = True
    End If
    i = NumRecords
    'BenchMark "LoadSymbols done"
    
    ' load all pool objects
    LoadSymbolGroups bRecalc
    LoadCriterias bRecalc
    LoadFilters
    
    ' clear any fields which no longer have valid objects
    RemoveOrphanedArraysFromTable
    
    MakeDebugFile
    
ErrExit:
    Exit Sub
    
ErrSection:
    KillFile AddSlash(App.Path) & "SYMPOOL.MEM"
    RaiseError "cSymbolPool.Load", eGDRaiseError_Raise
    
End Sub

Public Sub RemoveOrphanedArraysFromTable()
On Error GoTo ErrSection:

    Dim i&, bExists As Boolean
    
    For i = m.aFieldIDs.Size - 1 To 0 Step -1
        Select Case UCase(Parse(m.aFieldIDs(i), ":", 1))
            Case "GRP":
                bExists = m.SymbolGroups.Exists(Parse(m.aFieldIDs(i), ":", 2))
            Case "DSV", "DSP":
                bExists = m.Criterias.Exists(Parse(m.aFieldIDs(i), ":", 2))
            Case "FIL":
                bExists = m.Filters.Exists(Parse(m.aFieldIDs(i), ":", 2))
            Case "INF":
                bExists = True
            Case Else:
                bExists = False
        End Select
        If Not bExists Then
            m.ArrayTable.ClearField i
            m.aFieldIDs(i) = ""
        End If
    Next
    m.aFieldIDs.Size = m.ArrayTable.NumFields

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolPool.RemoveOrphanedArraysFromTable", eGDRaiseError_Raise
    
End Sub

Private Function Serialize(ByVal bPut As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim strFileName$, hFile&, bSuccess As Boolean
    Dim cSerialVersion As Byte, cCheck As Byte
    
    'increment cSerialVersion whenever format of
    'the Table or serialized file gets changed
    cSerialVersion = 22
    strFileName = AddSlash(App.Path) & "SYMPOOL.MEM"
    If bPut Then
        If NumRecords = 0 Then
            ' TLB 4/15/2005: if no records, don't leave file so will force rebuild at next startup
            KillFile strFileName
        Else
            ' Write FieldID's and Table to file
            hFile = gdFileOpen(strFileName, "wb")
            gdFileBinaryIO hFile, cSerialVersion, 1, bPut
            If m.aFieldIDs.Serialize(hFile, bPut) Then
                If m.ArrayTable.Serialize(hFile, bPut) Then
                    bSuccess = True
                End If
            End If
        End If
    ElseIf FileExist(strFileName) Then
        ' Read FieldID's and Table from file
        hFile = gdFileOpen(strFileName, "rb")
        
        ' First make sure cSerialVersion matches
        gdFileBinaryIO hFile, cCheck, 1, bPut
        If cCheck = cSerialVersion Then
            If m.aFieldIDs.Serialize(hFile, bPut) Then
                If m.ArrayTable.Serialize(hFile, bPut) Then
                    bSuccess = True
                End If
            End If
        End If
    End If
    gdFileClose hFile
    
    Serialize = bSuccess

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.Serialize", eGDRaiseError_Raise
    
End Function

'To load all the symbols and corresponding info
Private Function LoadSymbols(ByVal bDontReloadIfSameCount As Boolean) As Boolean
On Error GoTo ErrSection:

    'In order to minimize the "down time" when needing to reload
    'symbol info (e.g. after a distribution), we want to first gather
    'all the information into local arrays, then when we have it all
    'we'll do a quick "attach" to the DataTable.  This way, charts
    'and other stuff will still be operational with the old Table
    'while we're in the process of reloading the symbols.

    Dim rc&, i&, nCount&, tbl&, strTemp$, hArray&
    Dim nSymbolID&, strSymbol$, strSecType$, strDesc$
    Dim fh%, strOptNavFile$
    Dim aRecTrack() As Long
    Dim hSymbols As Long
    Dim b2ndTagMethod As Boolean
    Dim eSecType As eSYM_SecType
    
    Dim hSymbol&, hDbRecNum&, hSymbolID&, hDesc&
    Dim hSecType&, hFirstDate&, hLastDate&, hSymIndex&
    Dim hAllFutures&, hAllIndices&, hAllStocks&, hAllSymbols&
    Dim hAllIntlStocks&, hAllMutFunds&, hAllForex&
    
    ' get table handle and count of symbols
    tbl = g.Universe.tblSymbols
    If tbl = 0 Then Exit Function
    rc = TagSelect(tbl, g.Universe.tagAccessBySymbol)
    nCount = TagRecCount(tbl)
    If bDontReloadIfSameCount Then
        If nCount = m.ArrayTable.NumRecords Then
            LoadSymbols = False
            Exit Function
        End If
    End If

    ' first reset all the access flags
    SetAccessFlags
    
    rc = TagSelect(tbl, g.Universe.tagAccessBySymbol)
    nCount = TagRecCount(tbl)

    ' create local arrays
    hSymbol = gdCreateArray(eGDARRAY_Longs, nCount)
    hDbRecNum = gdCreateArray(eGDARRAY_Longs, nCount)
    hSymbolID = gdCreateArray(eGDARRAY_Longs, nCount)
    hDesc = gdCreateArray(eGDARRAY_Longs, nCount)
    hFirstDate = gdCreateArray(eGDARRAY_Longs, nCount)
    hLastDate = gdCreateArray(eGDARRAY_Longs, nCount)
    hSymIndex = gdCreateArray(eGDARRAY_Longs, nCount)
    hSecType = gdCreateArray(eGDARRAY_TinyInts, nCount)
    hAllFutures = gdCreateArray(eGDARRAY_TinyInts, nCount, 0)
    hAllIndices = gdCreateArray(eGDARRAY_TinyInts, nCount, 0)
    hAllStocks = gdCreateArray(eGDARRAY_TinyInts, nCount, 0)
    hAllIntlStocks = gdCreateArray(eGDARRAY_TinyInts, nCount, 0)
    hAllMutFunds = gdCreateArray(eGDARRAY_TinyInts, nCount, 0)
    hAllForex = gdCreateArray(eGDARRAY_TinyInts, nCount, 0)
    hAllSymbols = gdCreateArray(eGDARRAY_TinyInts, nCount, 0)

    BenchMark
    b2ndTagMethod = True ' more efficient (takes less than half the time)
    If b2ndTagMethod Then
        rc = d4recCount(tbl)
        ReDim aRecTrack(rc) As Long
        For i = 0 To rc
            aRecTrack(i) = -1
        Next
    Else
        hSymbols = gdCreateArray(eGDARRAY_Strings, nCount)
    End If
    
    ' Dump file for Option Nav?
    fh = 0
    'fh = FreeFile
    If fh <> 0 Then
        strOptNavFile = App.Path & "\Symbols.txt"
        Open strOptNavFile For Output As #fh
    End If
    
    'get info for each symbol
    nCount = 0
    rc = d4top(tbl)
    Do While rc = r4success
        ' store DbRecNum and SymbolID
        nSymbolID = f4long(g.Universe.fldSymbolID)
        strSecType = UCase(Trim(f4str(g.Universe.fldSecType)))
        strSymbol = UCase(Trim(f4str(g.Universe.fldSymbol)))
        If fh <> 0 Then
            strDesc = Trim(f4str(g.Universe.fldDesc))
            Print #fh, Str(nSymbolID) & vbTab & strSymbol & vbTab & strSecType & vbTab & strDesc
        End If
        
        ' ordered by Symbol, so just store index for sorting purposes
        gdSetNum hSymbol, nCount, nCount
        gdSetNum hDbRecNum, nCount, d4recNo(tbl)
        gdSetNum hSymbolID, nCount, nSymbolID
        gdSetNum hFirstDate, nCount, f4long(g.Universe.fldEodFirst)
        gdSetNum hLastDate, nCount, f4long(g.Universe.fldEodLast)
        
        Select Case strSecType
            Case "I", "X":
                If IsForex(strSymbol) Then
                    eSecType = eSYMType_Forex
                    gdSetNum hAllForex, nCount, 1
                Else
                    eSecType = eSYMType_Index
                    gdSetNum hAllIndices, nCount, 1
                End If
            Case "S":
                ' 9/20/2013: we'll now separate the U.S. and Int'l stocks into 2 different groups
                eSecType = eSYMType_Stock
                If InStr(strSymbol, "@") > 0 Then
                    gdSetNum hAllIntlStocks, nCount, 1
                Else
                    gdSetNum hAllStocks, nCount, 1
                End If
            Case "F":
                eSecType = eSYMType_Future
                gdSetNum hAllFutures, nCount, 1
            Case "M":
                eSecType = eSYMType_MutualFund
                gdSetNum hAllMutFunds, nCount, 1
            Case Else:
                eSecType = eSYMType_Unknown
        End Select
        gdSetNum hSecType, nCount, eSecType
        gdSetNum hAllSymbols, nCount, 1
        
        If b2ndTagMethod Then
            ' store index under record number for future ref.
            i = d4recNo(tbl)
            aRecTrack(i) = nCount
        Else
            ' store symbol w/index in temp. string array
            strTemp = UCase(Trim(f4str(g.Universe.fldSymbol))) _
                & Chr(9) & Trim(Str(nCount))
            gdSetStr hSymbols, nCount, strTemp
        End If
        
        nCount = nCount + 1
        rc = d4skip(tbl, 1)
    Loop
    
    If fh <> 0 Then Close #fh
   
    If b2ndTagMethod Then
        ' now order by desc
        rc = TagSelect(tbl, g.Universe.tagAccessByDesc)
        nCount = 0
        rc = d4top(tbl)
        Do While rc = r4success
            ' look up index by rec num
            i = d4recNo(tbl)
            i = aRecTrack(i)
            ' store count in array
            gdSetNum hDesc, i, nCount
            nCount = nCount + 1
            rc = d4skip(tbl, 1)
        Loop
    Else
        ' sort string array
        gdSort hSymbols, eGdSort_Default, 0, nCount - 1
        ' put index into symbol array
        For i = 0 To nCount - 1
            strTemp = gdGetStr(hSymbols, i)
            rc = InStr(strTemp, Chr(9))
            If rc > 0 Then
                rc = Val(Mid(strTemp, rc + 1))
                If rc >= 0 Then
                    gdSetNum hSymbol, rc, i
                End If
            End If
        Next
    End If
    gdDestroyArray hSymbols
    ReDim aRecTrack(0) As Long
            
    ' build index of SymbolID's (for quick searches)
    ''gdIndexList hSymIndex, 0, hSymbolID, 0
    i = gdSortAsIndex(hSymIndex, hSymbolID, True, eGdSort_Default, 0, -1)
   
    ' now put arrays into the DataTable
    With m.ArrayTable
        'first make all other arrays null const's
        '(until each object gets recalculated)
        For i = kOtherFields To m.ArrayTable.NumFields - 1
            hArray = .FieldArrayHandle(i)
            If hArray <> 0 Then
                gdMakeConstantValue hArray, gdNullValue(hArray), gdGetSize(hArray)
            End If
        Next
        .NumRecords = gdGetSize(hSymbolID)
        gdAttachField .TableHandle, hSymIndex, kSymIndex, "SymIndex"
        gdAttachField .TableHandle, hSymbol, kSymbol, "Symbol"
        gdAttachField .TableHandle, hSymbolID, kSymbolID, "Symbol ID"
        gdAttachField .TableHandle, hDesc, kDesc, "Description"
        gdAttachField .TableHandle, hDbRecNum, kDbRecNum, "DbRecNum"
        gdAttachField .TableHandle, hFirstDate, kFirstDate, "FirstDate"
        gdAttachField .TableHandle, hLastDate, kLastDate, "LastDate"
        gdAttachField .TableHandle, hSecType, kSecType, "Type"
        gdAttachField .TableHandle, hAllSymbols, kAllSymbols, "ALL Symbols"
        gdAttachField .TableHandle, hAllFutures, kAllFutures, "All FUTURES"
        gdAttachField .TableHandle, hAllIndices, kAllIndices, "All INDICES"
        gdAttachField .TableHandle, hAllForex, kAllForex, "All FOREX"
        ' 9/20/2013: for backwards-compatibility, should keep the group name "All STOCKS" the same,
        ' even though it's now only doing U.S. stocks (but now the descriptive label gets changed elsewhere)
        gdAttachField .TableHandle, hAllStocks, kAllStocks, "All STOCKS"
        gdAttachField .TableHandle, hAllIntlStocks, kAllIntlStocks, "All INT'L STOCKS"
        gdAttachField .TableHandle, hAllMutFunds, kAllMutFunds, "All MUTUAL FUNDS"
        
        'create FieldID's for these fields
        For i = 0 To kOtherFields - 1
            If i < kAllSymbols Then
                m.aFieldIDs(i) = "INF:" & UCase(.FieldName(i))
            End If
        Next
    End With
    
    'destroy local copies
    gdDestroyArray hSymbol
    gdDestroyArray hSymbolID
    gdDestroyArray hDesc
    gdDestroyArray hDbRecNum
    gdDestroyArray hSecType
    gdDestroyArray hFirstDate
    gdDestroyArray hLastDate
    gdDestroyArray hSymIndex
    gdDestroyArray hAllFutures
    gdDestroyArray hAllIndices
    gdDestroyArray hAllStocks
    gdDestroyArray hAllIntlStocks
    gdDestroyArray hAllMutFunds
    gdDestroyArray hAllForex
    gdDestroyArray hAllSymbols
    
If IsIDE Then
    'BenchMark Str(nCount) & " symbols loaded"
End If

    LoadSymbols = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.LoadSymbols", eGDRaiseError_Raise
    
End Function

Private Sub LoadSpecialSymbolGroup(ByVal nField&)
On Error GoTo ErrSection:
    
    Dim SymbolGroup As New cSymbolGroup
    Dim strName$, strID$
    
    strName = m.ArrayTable.FieldName(nField)
    ' 3/3/2014: needed to move this code for checking the "ALL STOCKS" up to here
    ' 9/20/2013: for backwards-compatibility, we need to keep the "All STOCKS" Group ID the same,
    ' even though it's now only doing U.S. stocks -- but we can change the descriptive label here ...
    If UCase(strName) = "ALL STOCKS" Or UCase(strName) = "ALL U.S. STOCKS" Then
        strName = "All U.S. STOCKS"
        strID = "ALL STOCKS.GRP"
    Else
        strID = UCase(strName) & ".GRP"
    End If
    m.aFieldIDs(nField) = "GRP:" & strID
    SymbolGroup.MakeSpecialType strID, strName, eGROUP_Builtin
    SymbolGroup.AddToPool
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolPool.LoadSpecialSymbolGroup", eGDRaiseError_Raise
    
End Sub

Private Sub LoadSymbolGroups(ByVal bRecalc As Boolean)
On Error GoTo ErrSection:
    
    Dim i&, strPath$, strFldID$, bLoadSymbols As Boolean
    Dim aFiles As New cGdArray
    Dim SymbolGroup As cSymbolGroup

    'Load predefined SymbolGroups
    LoadSpecialSymbolGroup kAllSymbols
    LoadSpecialSymbolGroup kAllFutures
    LoadSpecialSymbolGroup kAllIndices
    LoadSpecialSymbolGroup kAllStocks
    LoadSpecialSymbolGroup kAllIntlStocks
    LoadSpecialSymbolGroup kAllMutFunds
    LoadSpecialSymbolGroup kAllForex

    'Load standard SymbolGroups
    strPath = App.Path & "\Provided\"
    aFiles.GetMatchingFiles strPath & "*.GRP", False
    For i = 0 To aFiles.Size - 1
        Set SymbolGroup = New cSymbolGroup
        ' reload symbols if a new group (that didn't exist before)
        strFldID = "GRP:" & UCase(aFiles(i))
        If g.SymbolPool.FieldNumForID(strFldID) < 0 Then
            bLoadSymbols = True
        Else
            ' else reload symbols if doing a recalc
            bLoadSymbols = bRecalc
        End If
        ' load the group from the file
        If SymbolGroup.FromFile(strPath, aFiles(i), bLoadSymbols) Then
            SymbolGroup.AddToPool bLoadSymbols
        End If
    Next

    'Load custom SymbolGroups
    strPath = App.Path & "\Custom\"
    aFiles.GetMatchingFiles strPath & "*.GRP", False
    For i = 0 To aFiles.Size - 1
        Set SymbolGroup = New cSymbolGroup
        With SymbolGroup
            ' reload symbols if a new group (that didn't exist before)
            strFldID = "GRP:" & UCase(aFiles(i))
            If g.SymbolPool.FieldNumForID(strFldID) < 0 Then
                bLoadSymbols = True
            Else
                ' else reload symbols if doing a recalc
                bLoadSymbols = bRecalc
            End If
            If .FromFile(strPath, aFiles(i), bLoadSymbols) Then
                .AddToPool bLoadSymbols
                ' see if need to create the custom index from a new file that just "showed up"
                If .IsIndex Then
                    .SymbolID = g.SymbolPool.SymbolIDforSymbol(UCase("#" & .Name))
                    If .SymbolID = 0 Then
                        If SU_SetCompositeInf(.SymbolID, UCase("#" & .Name), .Desc, .PriceDivisor, _
                                .SymbolIDs, .PriceWeights, .Flags, _
                                .VolDivisor, .VolWeights) Then
                            If g.SymbolPool.AddCustomIndex(.SymbolID, UCase("#" & .Name)) = False Then
                                'InfBox "Problems saving #" & .Name, "!", , "Error"
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next
    
ErrExit:
    Set SymbolGroup = Nothing
    Exit Sub
    
ErrSection:
    Set SymbolGroup = Nothing
    RaiseError "cSymbolPool.LoadSymbolGroups", eGDRaiseError_Raise
    
End Sub

Private Sub LoadCriterias(ByVal bRecalc As Boolean)
On Error GoTo ErrSection:
    
    Dim i&, j&, strPath$, strText$
    Dim iCategoryID&, strCategoryName$
    
    Dim aFiles As New cGdArray
    
    Dim tbCriteriaGroups As cGdTable
    Dim aIdx As cGdArray
    
    Dim Criteria As cCriteria
    Dim Expr As cExpression

    'Load criteria category name
    strPath = App.Path & "\Provided\CriteriaGroups.txt"
    If FileExist(strPath) Then
        Set tbCriteriaGroups = New cGdTable
        tbCriteriaGroups.CreateField eGDARRAY_Longs, 0, "CategoryID", -1        'numeric value criteriagroups.txt
        tbCriteriaGroups.CreateField eGDARRAY_Strings, 1, "CategoryName"        'category name in criteriagroups.txt
        tbCriteriaGroups.CreateField eGDARRAY_Strings, 2, "CriteriaID"          'criteria ID, which is the criteria .SCN file name
        
        aFiles.FromFile strPath
        For i = 0 To aFiles.Size - 1
            strText = UCase(Parse(aFiles(i), vbTab, 1))
            If Left(strText, 1) <> "'" Then          'skip comment lines
                Select Case strText
                    Case "GROUP"
                        iCategoryID = 0
                        strCategoryName = ""
                        
                        iCategoryID = ValOfText(Parse(aFiles(i), vbTab, 2))
                        If iCategoryID > 0 Then
                            strCategoryName = Parse(aFiles(i), vbTab, 3)
                        End If
                        
                    Case "ITEM"
                        If iCategoryID > 0 Then
                            tbCriteriaGroups.AddRecord ""
                            j = tbCriteriaGroups.NumRecords - 1
                            tbCriteriaGroups(0, j) = iCategoryID
                            tbCriteriaGroups(1, j) = strCategoryName
                            
                            strText = UCase(Parse(aFiles(i), vbTab, 2))         'criteria ID
                            tbCriteriaGroups(2, j) = strText
                        End If
                End Select
            End If
        Next
        Set aIdx = tbCriteriaGroups.CreateSortedIndex(2, eGdSort_Default)       'sort by criteria ID
    End If

    'Load standard Criterias
    strPath = App.Path & "\Provided\"
    aFiles.GetMatchingFiles strPath & "*.SCN", False
    For i = 0 To aFiles.Size - 1
        strText = aFiles(i)
        Set Criteria = New cCriteria
        If Criteria.FromFile(strPath, strText) Then
            Criteria.AddToPool
            If Not aIdx Is Nothing Then
                strText = UCase(strText)
                If tbCriteriaGroups.SearchAsIndex(aIdx, 2, strText, j) Then
                    j = aIdx(j)
                    If tbCriteriaGroups(2, j) = strText Then
                        Criteria.CategoryID = tbCriteriaGroups(0, j)
                        Criteria.CategoryName = tbCriteriaGroups(1, j)
                    End If
                End If
            End If
        End If
    Next

    'Load custom Criterias
    strPath = App.Path & "\Custom\"
    aFiles.GetMatchingFiles strPath & "*.SCN", False
    For i = 0 To aFiles.Size - 1
        strText = aFiles(i)
        Set Criteria = New cCriteria
        If Criteria.FromFile(strPath, aFiles(i)) Then
            ' If the Formatted Text has not been generated yet, then generate it...
            If Criteria.FormattedText = "" And Criteria.EnglishText <> "" Then
                If Not g.Functions Is Nothing Then
                    Set Expr = New cExpression
                    With Expr
                        .PortfolioNavigator = False
                        .Functions = g.Functions
                        On Error Resume Next
                        .ValidateFunctionRule Criteria.EnglishText
                        On Error GoTo ErrSection:
                        If Len(Trim(Expr.CodedText)) > 0 Then
                            Criteria.FormattedText = Expr.EditText
                            Criteria.ToFile True
                        End If
                    End With
                End If
            End If
            
            Criteria.AddToPool
        
            If Not aIdx Is Nothing Then
                strText = UCase(strText)
                If tbCriteriaGroups.SearchAsIndex(aIdx, 2, strText, j) Then
                    j = aIdx(j)
                    If tbCriteriaGroups(2, j) = strText Then
                        Criteria.CategoryID = tbCriteriaGroups(0, j)
                        Criteria.CategoryName = tbCriteriaGroups(1, j)
                    End If
                End If
            End If
        End If
    Next
    
    If bRecalc Then DirtyCriteria = True
    
ErrExit:
    Set Criteria = Nothing
    Exit Sub
    
ErrSection:
    Set Criteria = Nothing
    RaiseError "cSymbolPool.LoadCriterias", eGDRaiseError_Raise
    
End Sub

Private Sub LoadFilters()
On Error GoTo ErrSection:
    
    Dim i&, strPath$
    Dim aFiles As New cGdArray
    Dim Filter As cFilter

    'Load standard Filters
    strPath = App.Path & "\Provided\"
    aFiles.GetMatchingFiles strPath & "*.FIL", False
    For i = 0 To aFiles.Size - 1
        Set Filter = New cFilter
        If Filter.FromFile(strPath, aFiles(i)) Then
            Filter.AddToPool
        End If
    Next

    'Load custom Criterias
    strPath = App.Path & "\Custom\"
    aFiles.GetMatchingFiles strPath & "*.FIL", False
    For i = 0 To aFiles.Size - 1
        Set Filter = New cFilter
        If Filter.FromFile(strPath, aFiles(i)) Then
            Filter.AddToPool
        End If
    Next
    
ErrExit:
    Set Filter = Nothing
    Exit Sub
    
ErrSection:
    Set Filter = Nothing
    RaiseError "cSymbolPool.LoadFilters", eGDRaiseError_Raise
    
End Sub

#If 0 Then
Public Function AddToPool(obj As Object, _
        ByVal strID$, _
        Optional ByVal strName$ = "", _
        Optional ByVal nFieldNum& = 0, _
        Optional ByVal bSpecial As Boolean = False)
    
    Dim nFld&
    
    If nFieldNum = 0 Then
        ' find first empty spot
        For nFld = 1 To mObjectPool.Count
            If mObjectPool(nFld) Is Nothing Then
                nFieldNum = nFld
                Exit For
            End If
        Next
        ' if no empty spots, append to end
        If nFieldNum = 0 Then
            nFieldNum = mObjectPool.Count + 1
        End If
    End If
    
    If strName = "" Then strName = strID
    Select Case TypeName(obj)
        Case "cSymbolInfo":
            mDataTable.CreateField obj.ArrayType, nFieldNum, strName
        Case "cGdArray":
            'If bSpecial Then
            '    Set obj = mDB.FieldArray(nFieldNum)
            'Else
            '    mDB.SetField obj, strName, nFieldNum
            'End If
            mDataTable.AttachField obj, nFieldNum, strName
        Case "cCriteria":
            'mDB.SetField obj.aValues, strName, nFieldNum
            mDataTable.AttachField obj.aValues, nFieldNum, strName
        Case "cFilter":
            'mDB.SetField obj.aValues, strName, nFieldNum
            mDataTable.AttachField obj.aValues, nFieldNum, strName
        Case "cSymbolGroup":
            'mDB.SetField obj.aFilter, strName, nFieldNum
            mDataTable.AttachField obj.aFilter, nFieldNum, strName
    End Select
    
    Set mObjectPool(nFieldNum) = obj
    mObjectPool.Key(nFieldNum) = strID
    mObjectPool.NodeType(nFieldNum) = TypeName(obj)
    
    AddToPool = nFieldNum
End Function
#End If


Public Function RecalcDirtyCriteria(Optional ByVal nForDate As Long = 0, _
                                    Optional ByVal strSymbolToCalc As String = "") As Boolean
On Error GoTo ErrSection:

    Dim iPercentDone&, bFewSymbols As Boolean
    Dim bOnlySelectedSymbols As Boolean
    Dim bIncludeSnapshotData As Integer

    Dim i&, rc&, d#, hArray&, lLastDateOfData&, j&, fh%, s$
    Dim dStartTime#, dPrevSleepTime#, dPrevFullYieldTime#
    Dim nPoolRec&, nSymbolID&, nCount&, nFld&
    Dim nNumDays&, dLastDate#, nFromDate&
    Dim strCodedText$, strSymbol$, strFile$
    Dim Criteria As cCriteria
    Dim Filter As cFilter
    Dim Bars As New cGdBars, Weekly As New cGdBars, Monthly As New cGdBars
    Dim Bars67 As New cGdBars, Bars67Weekly As New cGdBars, Bars67Monthly As New cGdBars
    'Dim GC As New cGdBars, GCWeekly As New cGdBars
    Dim aSymbolRecs As New cGdArray
    Dim aTemp As New cGdArray
    
    Dim astrParms As New cGdArray, astrBarNames As New cGdArray
    Dim aScanExpr As New cGdArray, aArrayOfResults As New cGdArray
    Dim aArrayOfBars As New cGdArray
    Dim aScanArrays As New cGdArray
    Dim aScanNames As New cGdArray

    Dim astrParmsW As New cGdArray ', astrBarNamesW As New cGdArray
    Dim aScanExprW As New cGdArray, aArrayOfResultsW As New cGdArray
    Dim aScanArraysW As New cGdArray
    Dim aScanNamesW As New cGdArray
        
    Dim lIndex As Long
    Dim lLastCompletedWeek As Long
    Dim lKindID As Long
    Dim dValue#
    Dim lActiveDate&
    Dim lBarDate&
    Dim iLifetime As Integer
    
    Dim bContinue As Boolean
        
    Dim aGroupFld As New cGdArray, aNumDays As New cGdArray
    Dim aGroupFldW As New cGdArray, aNumDaysW As New cGdArray
    Dim barsDow As New cGdBars
    
    Dim alFundIDs As New cGdArray
    Dim alFundLifetime As New cGdArray
    Dim aScanArraysF As New cGdArray
    Dim aGroupFldF As New cGdArray
    
    Dim i67&, i67Weekly&, i67Monthly&
       
    ' Get config values relevant for scans
    If Not ScansEnabled Then
        Exit Function
    End If
    
    If Len(strSymbolToCalc) = 0 Then
        If Not frmStatus.Visible Then
            frmStatus.Status = eStatus_Initialized
            frmStatus.SetTitle "Calculating Criteria"
            DoEvents
        End If
    End If
   
    dStartTime = gdTickCount
gdResetProfiles 700, 799
gdStartProfile 701

    ' TLB
    If nForDate >= 19000000 Then
        nForDate = JulFromLong(nForDate)
    End If
    If nForDate = 0 Then
        ' default to last daily download
        nForDate = LastDailyDownload
    ElseIf Abs(nForDate) <= 2 Then
        ' for backwards-compatibility: if True or 1 or 2 was passed in, then want to include snapshot
        nForDate = LastDailyDownload + 1
        Do While Not IsWeekday(nForDate)
            nForDate = nForDate + 1
        Loop
    End If
    If nForDate > LastDailyDownload Then
        bIncludeSnapshotData = 2 ' TLB: must now do this to override when salmon is running
    End If
    If nForDate < LastDailyDownload Then
        lLastDateOfData = nForDate
    Else
        lLastDateOfData = LastDailyDownload
    End If

    
    ' Get coded text and handle of values array from each Criteria
    aScanExpr.Create eGDARRAY_Strings
    aScanNames.Create eGDARRAY_Strings
    aScanArrays.Create eGDARRAY_Longs
    aArrayOfResults.Create eGDARRAY_Longs
    aScanExprW.Create eGDARRAY_Strings
    aScanNamesW.Create eGDARRAY_Strings
    aScanArraysW.Create eGDARRAY_Longs
    aArrayOfResultsW.Create eGDARRAY_Longs
    aGroupFld.Create eGDARRAY_Longs
    aNumDays.Create eGDARRAY_Longs
    aGroupFldW.Create eGDARRAY_Longs
    aNumDaysW.Create eGDARRAY_Longs
    nNumDays = 0&
        
    alFundIDs.Create eGDARRAY_Longs
    alFundLifetime.Create eGDARRAY_Longs
    aScanArraysF.Create eGDARRAY_Longs
    aGroupFldF.Create eGDARRAY_Longs
        
    For Each Criteria In m.Criterias
        With Criteria
            lKindID = FundamentalOnly(.EnglishText)
            strCodedText = Trim(.CodedText)
            If Len(strCodedText) = 0 Then .IsActive = False
            If (.IsDirty = True Or Len(strSymbolToCalc) > 0) And .IsActive = True Then
                If lKindID = 0 Then
                    If .IsWeekly = False Then
                        aScanExpr.Add strCodedText
                        aScanNames.Add .Name
                    Else
                        aScanExprW.Add strCodedText
                        aScanNamesW.Add .Name
                    End If
                    
                    ' get values array handle, clear array (so no longer
                    ' a const array), pre-size array, and store handle
                    hArray = .ValuesArray.ArrayHandle
                    If Len(strSymbolToCalc) = 0 Then
                        gdClear hArray, True
                    Else
                        nFld = g.SymbolPool.FieldNumForID("DSV:" & .ID)
                        If nFld >= 0 Then
                            gdCopy hArray, g.SymbolPool.ArrayTable.FieldArrayHandle(nFld)
                        End If
                    End If
                    gdSetSize hArray, m.ArrayTable.NumRecords - 1, True
                    If .IsWeekly = False Then
                        aScanArrays.Add hArray
                        ' store symbol group used for this criteria
                        aGroupFld.Add g.SymbolPool.FieldNumForID(.GroupID)
                        aArrayOfResults.Add 0
                    Else
                        aScanArraysW.Add hArray
                        ' store symbol group used for this criteria
                        aGroupFldW.Add g.SymbolPool.FieldNumForID(.GroupID)
                        aArrayOfResultsW.Add 0
                    End If
    
                    ' see if NumDays bigger
                    If .NumDays <= 0 Then
                        .NumDays = frmCriteria.AutoDetect(.CodedText)
                        If .NumDays > 0 Then .ToFile
                    End If
                    If .IsWeekly = False Then
                        aNumDays.Add .NumDays
                        If .NumDays > nNumDays Then nNumDays = .NumDays
                    Else
                        aNumDaysW.Add (.NumDays + 1) * 5
                        If (.NumDays + 1) * 5 > nNumDays Then nNumDays = (.NumDays + 1) * 5
                    End If
                Else
                    ' get values array handle, clear array (so no longer
                    ' a const array), pre-size array, and store handle
                    hArray = .ValuesArray.ArrayHandle
                    If Len(strSymbolToCalc) = 0 Then
                        gdClear hArray, True
                    Else
                        nFld = g.SymbolPool.FieldNumForID("DSV:" & .ID)
                        If nFld >= 0 Then
                            gdCopy hArray, g.SymbolPool.ArrayTable.FieldArrayHandle(nFld)
                        End If
                    End If
                    gdSetSize hArray, m.ArrayTable.NumRecords - 1, True
                    aScanArraysF.Add hArray
                    aGroupFldF.Add g.SymbolPool.FieldNumForID(.GroupID)
                    alFundIDs.Add lKindID
                    ' get lifetime (fill) for this data kind
                    If DM_GetDataKindLifetime(g.DMS, lKindID, iLifetime) <> 0 Then
                        alFundLifetime.Add iLifetime
                    Else
                        alFundLifetime.Add 0
                    End If
                End If
            End If
        End With
    Next Criteria
    
    ' debug file
    strFile = App.Path & "\chk\ScanExpr.chk"
    aScanExpr.ToFile strFile, False
    aScanExprW.ToFile strFile, True
    FileFromString strFile, "NumDays = " & CStr(nNumDays), True, True

gdStopProfile 701
gdStartProfile 702
    
    ' calc FromDate, adjusting for weekends and holidays
    ' (need to fudge a little to the safe side)
    nFromDate = 0
    ' (new method: load DJ and get num days from it)
    If DM_GetBars(barsDow, "$DJIA", 0, 0, nForDate, , , , bIncludeSnapshotData) Then
        ' fudge for foreign holidays (extra day per quarter), etc.
        'i = barsDow.Size - nNumDays - 2 - Int(nNumDays / 66)
        i = barsDow.Size - nNumDays - 2 - Int(nNumDays / 22)
        If i >= 0 Then
            nFromDate = barsDow(eBARS_DateTime, i)
        End If
    End If
    If nFromDate <= 0 Then
        ' (old method: try this algorithm)
        'nFromDate = lLastDateOfData - Int(nNumDays * 1.46 + 0.5) - 2
        nFromDate = lLastDateOfData - Int(nNumDays * 1.6 + 0.5) - 2
    End If
    
    '7/9/03 TLB: don't need anything special for a Friday holiday anymore.
    'calc date of last completed week (fudge a little in case of a Friday holiday)
    'If Weekday(lLastDateOfData) = vbThursday And lLastDateOfData >= Int(Now) - 4 _
        And Now > lLastDateOfData + 1.5 Then
    '        lLastCompletedWeek = lLastDateOfData + 1
    'Else
            lLastCompletedWeek = lLastDateOfData
    'End If
    
    Dim lBars As Long
    Dim BarsColl As New cGdTree
    astrBarNames(0) = "Market1"
    astrBarNames(1) = "Weekly"
    astrBarNames(2) = "Monthly"
    astrBarNames(3) = "Daily"
    MarketsInExpressions aScanExpr, nFromDate, bIncludeSnapshotData, astrBarNames, BarsColl, "Daily"
    MarketsInExpressions aScanExprW, nFromDate, bIncludeSnapshotData, astrBarNames, BarsColl, "Weekly"
    ' TLB 6/7/2011: check if "-067" is used
    i67 = -1
    i67Weekly = -1
    i67Monthly = -1
    For i = 0 To astrBarNames.Size - 1
        Select Case UCase(astrBarNames(i))
            Case Chr(34) & "-067,DAILY" & Chr(34)
                i67 = i
            Case Chr(34) & "-067,WEEKLY" & Chr(34)
                i67Weekly = i
            Case Chr(34) & "-067,MONTHLY" & Chr(34)
                i67Monthly = i
        End Select
    Next

gdStopProfile 702
    
    If aScanExpr.Size + aScanExprW.Size + alFundIDs.Size > 0 Then
gdStartProfile 703
        
        If aScanExpr.Size > 0 Then
            astrParms(0) = "DailyScans"
            If Not SetupExpressions(astrParms, astrBarNames, aScanExpr) Then
                ' try to find out which one's bombing
                For i = 0 To aScanExpr.Size - 1
                    ' (use temp array to run just one item at a time)
                    aTemp(0) = aScanExpr(i)
                    aTemp.Size = 1
                    If Not SetupExpressions(astrParms, astrBarNames, aTemp) Then
                        InfBox "i=[] ; h=Recalc Criteria ; An error exists in the Criteria expression for:|" _
                            & aScanNames(i)
                        Exit Function
                    End If
                Next
                InfBox "i=[] ; h=Recalc Criteria ; An error exists in a Criteria expression."
                Exit Function
            End If
        End If
        
        If aScanExprW.Size > 0 Then
            astrParmsW(0) = "WeeklyScans"
            If Not SetupExpressions(astrParmsW, astrBarNames, aScanExprW) Then
                ' try to find out which one's bombing
                For i = 0 To aScanExprW.Size - 1
                    ' (use temp array to run just one item at a time)
                    aTemp(0) = aScanExprW(i)
                    aTemp.Size = 1
                    If Not SetupExpressions(astrParmsW, astrBarNames, aTemp) Then
                        InfBox "i=[] ; h=Recalc Criteria ; An error exists in the Criteria expression for:|" _
                            & aScanNamesW(i)
                        Exit Function
                    End If
                Next
                InfBox "i=[] ; h=Recalc Criteria ; An error exists in a Weekly Criteria expression."
                Exit Function
            End If
        End If
    
        ' get list of symbols to do (if debugging)
        aSymbolRecs.Create eGDARRAY_Longs
        If bOnlySelectedSymbols Then
            With frmSymbolGrid.fgVirtual
                For i = 0 To .SelectedRows - 1
                    strSymbol = UCase(.TextMatrix(.SelectedRow(i), kSymbolCol))
                    nPoolRec = g.SymbolPool.PoolRecForSymbol(strSymbol)
                    If nPoolRec >= 0 Then aSymbolRecs.Add nPoolRec
                Next
            End With
        ElseIf Len(strSymbolToCalc) > 0 Then
            bOnlySelectedSymbols = True
            For i = 1 To 99999
                strSymbol = Parse(strSymbolToCalc, vbTab, i)
                If Len(strSymbol) = 0 Then Exit For
                nPoolRec = g.SymbolPool.PoolRecForSymbol(strSymbol)
                If nPoolRec >= 0 Then aSymbolRecs.Add nPoolRec
            Next
        End If
        aSymbolRecs.Sort eGdSort_Descending
    
        ' For each symbol:
        If Len(strSymbolToCalc) = 0 Then frmStatus.AddDetail "Recalculating Criteria for " & DateFormat(nForDate)
        
        If bOnlySelectedSymbols Then
            If Len(strSymbolToCalc) = 0 Then frmStatus.AddDetail "(only scanning selected symbols)"
        ElseIf FileExist(App.Path & "\SCANS.FEW") Then
            If Len(strSymbolToCalc) = 0 Then frmStatus.AddDetail "(only scanning some symbols)"
            bFewSymbols = True
        End If
        If Len(strSymbolToCalc) = 0 Then frmStatus.Status = eStatus_Running
        iPercentDone = -1 'so will show when 0%
        BenchMark
        aArrayOfBars.Create eGDARRAY_Longs
        
gdStopProfile 703
        
Dim alDataIDs As New cGdArray
Dim adValues As New cGdArray
Dim alDates As New cGdArray
        
        d = gdTickCount
        dPrevSleepTime = d
        dPrevFullYieldTime = d
        For nPoolRec = 0 To m.ArrayTable.NumRecords - 1 'Step 10
gdStartProfile 710
        
            nSymbolID = m.ArrayTable.Num(kSymbolID, nPoolRec)
            Bars.Size = 0
            
            If bFewSymbols Then
                If nPoolRec > 500 And nPoolRec < m.ArrayTable.NumRecords - 300 Then
                    If Left(Symbol(nPoolRec), 1) <> "$" Then
                        nSymbolID = 0
                    End If
                End If
            End If
            If bOnlySelectedSymbols Then
                If nPoolRec = aSymbolRecs.Num(aSymbolRecs.Size - 1) Then
                    aSymbolRecs.Size = aSymbolRecs.Size - 1
                Else
                    nSymbolID = 0
                End If
            End If
gdStopProfile 710

            If nSymbolID <> 0 Then
                strSymbol = Symbol(nPoolRec)
If strSymbol = "IBM" Then
    fh = fh
End If
                'If bOnlySelectedSymbols And strSymbol = "ACAM" Then
                If strSymbol = "xLH-067" Then
                    fh = FreeFile
                    Open App.Path & "\chk\" & strSymbol & ".chk" For Output As #fh
                    Print #fh, "Symbol = " & strSymbol
                ElseIf fh <> 0 Then
                    Close #fh
                    fh = 0
                End If
                
                If alFundIDs.Size > 0 Then
gdStartProfile 711
                    If DM_GetAllSnapData(nSymbolID, alDataIDs, adValues, alDates) Then
gdStopProfile 711
gdStartProfile 712
                        For i = 0 To alFundIDs.Size - 1
                            If m.ArrayTable.Num(aGroupFldF.Num(i), nPoolRec) = 1 Then
                                hArray = aScanArraysF.Num(i)
                                dValue = gdNullValue(hArray)
                                If alDataIDs.BinarySearch(alFundIDs.Num(i), lIndex) Then
                                    'see if date is "in range"
                                    lActiveDate = alDates.Num(lIndex)
                                    iLifetime = alFundLifetime(i)
                                    lBarDate = LastDailyDownload
                                    If lBarDate > 0 And lActiveDate > 0 Then
                                        If lBarDate = lActiveDate Then
                                            'exact date match
                                            dValue = adValues.Num(lIndex)
                                        ElseIf iLifetime > 0 And lBarDate > lActiveDate Then
                                            'post-fill data (up to "n" days after value)
                                            If lBarDate <= lActiveDate + iLifetime Then
                                                dValue = adValues.Num(lIndex)
                                            End If
                                        ElseIf iLifetime < 0 And lBarDate < lActiveDate Then
                                            'pre-fill data (up to "n" days prior to value)
                                            If lBarDate >= lActiveDate + iLifetime Then
                                                dValue = adValues.Num(lIndex)
                                            End If
                                        End If
                                    End If
                                End If
                                gdSetNum hArray, nPoolRec, dValue
                            End If
                        Next i
gdStopProfile 712
                    End If
gdStopProfile 711
                End If
            
                ' See how many days need to be loaded for this symbol
gdStartProfile 720
                nNumDays = 0
                For i = 0 To aNumDays.Size - 1
                    If m.ArrayTable.Num(aGroupFld.Num(i), nPoolRec) = 1 Then
                        If aNumDays.Num(i) > nNumDays Then
                            nNumDays = aNumDays.Num(i)
                        End If
                        If aArrayOfResults.Num(i) = 0 Then
                            ' create results array for this criteria
                            aArrayOfResults.Num(i) = gdCreateArray(eGDARRAY_Doubles)
                        End If
                    ElseIf aArrayOfResults.Num(i) <> 0 Then
                        ' destroy this results array since not needed
                        gdDestroyArray aArrayOfResults.Num(i)
                        aArrayOfResults.Num(i) = 0
                    End If
                Next
gdStopProfile 720
gdStartProfile 721
                For i = 0 To aNumDaysW.Size - 1
                    If m.ArrayTable.Num(aGroupFldW.Num(i), nPoolRec) = 1 Then
                        If aNumDaysW.Num(i) > nNumDays Then
                            nNumDays = aNumDaysW.Num(i)
                        End If
                        If aArrayOfResultsW.Num(i) = 0 Then
                            ' create results array for this criteria
                            aArrayOfResultsW.Num(i) = gdCreateArray(eGDARRAY_Doubles)
                        End If
                    ElseIf aArrayOfResultsW.Num(i) <> 0 Then
                        ' destroy this results array since not needed
                        gdDestroyArray aArrayOfResultsW.Num(i)
                        aArrayOfResultsW.Num(i) = 0
                    End If
                Next
gdStopProfile 721
gdStartProfile 722

                If nNumDays > 0 Then
                    nFromDate = 0
                    ' fudge for foreign holidays (extra day per quarter), etc.
                    'i = barsDow.Size - nNumDays - 2 - Int(nNumDays / 66)
                    i = barsDow.Size - nNumDays - 2 - Int(nNumDays / 22)
                    If i >= 0 Then
                        nFromDate = barsDow(eBARS_DateTime, i)
                    End If
                    ' load the data for this symbol
                    If Not DM_GetBars(Bars, nSymbolID, 0, nFromDate, nForDate, , , , bIncludeSnapshotData) Then
                        Bars.Size = 0
                    End If
                End If
gdStopProfile 722
                
                If strSymbol = "$DJIA" Or strSymbol = "IBM" Or strSymbol = "SP-067" Or strSymbol = "SP-057" Or strSymbol = "TRAD" Then
                    FileFromString strFile, strSymbol & " " & CStr(nNumDays) & " " & CStr(Bars.Size), True, True
                End If
            End If

            If Bars.Size > 0 Then
                nCount = nCount + 1
gdStartProfile 723
                Weekly.BuildBars "Weekly", Bars.BarsHandle
                Monthly.BuildBars "Monthly", Bars.BarsHandle
gdStopProfile 723

                ' load the 67's (if were used in the expressions)
                If i67 > 0 Or i67Weekly > 0 Or i67Monthly > 0 Then
                    If Bars.SecurityType <> "F" Or Right(strSymbol, 4) = "-067" Then
                        ' just make the -067 be same as Market1
                        Set Bars67 = Bars
                        Set Bars67Weekly = Weekly
                        Set Bars67Monthly = Monthly
                    Else
gdStartProfile 724
                        ' need the load the -067 data
                        s = Parse(strSymbol, "-", 1) & "-067"
                        If Not DM_GetBars(Bars67, s, 0, nFromDate, nForDate, , , , bIncludeSnapshotData) Then
                            Bars67.Size = 0
                            Bars67Weekly.Size = 0
                            Bars67Monthly.Size = 0
                        Else
                            Bars67Weekly.BuildBars "Weekly", Bars67.BarsHandle
                            Bars67Monthly.BuildBars "Monthly", Bars67.BarsHandle
                        End If
gdStopProfile 724
                    End If
                End If
                               
                ' make sure date of last bar is within last 5 days
                dLastDate = Bars(eBARS_DateTime, Bars.Size - 1)
                If dLastDate >= lLastDateOfData - 5 And aScanExpr.Size > 0 Then
gdStartProfile 730
                    ' run engine to evaluate expressions for this symbol
                    aArrayOfBars.Num(0) = Bars.BarsHandle '(in case changed)
                    aArrayOfBars.Num(1) = Weekly.BarsHandle
                    aArrayOfBars.Num(2) = Monthly.BarsHandle
                    aArrayOfBars.Num(3) = Bars.BarsHandle
                    For lBars = 4 To astrBarNames.Size - 1
                        Select Case lBars
                        Case i67
                            aArrayOfBars.Num(lBars) = Bars67.BarsHandle
                        Case i67Weekly
                            aArrayOfBars.Num(lBars) = Bars67Weekly.BarsHandle
                        Case i67Monthly
                            aArrayOfBars.Num(lBars) = Bars67Monthly.BarsHandle
                        Case Else
                            aArrayOfBars.Num(lBars) = BarsColl(lBars + 1).BarsHandle
                        End Select
                    Next lBars
                    astrParms.Size = 2
                    rc = RunExpressions(astrParms.ArrayHandle, _
                        astrBarNames.ArrayHandle, aArrayOfBars.ArrayHandle, _
                        aArrayOfResults.ArrayHandle, ByVal 0&, ByVal 0&)
gdStopProfile 730
                    
                    If rc <> 0 Then
                        'frmMain.Caption = CStr(rc) & ": " & aStrParms(aStrParms.Size - 1)
                    Else
gdStartProfile 731
                        ' set current value for each expression
                        For i = 0 To aScanArrays.Size - 1
                            ' get most recent value
                            hArray = aArrayOfResults.Num(i)
                            If hArray = 0 Then
                                d = kNullData
                            Else
                                d = gdGetNum(hArray, Bars.Size - 1)
                                If fh <> 0 Then
                                    Print #fh, ""
                                    Print #fh, "Expr = " & aScanExpr(i)
                                    For j = 0 To gdGetSize(hArray) - 1
                                        If Bars(eBARS_DateTime, j) > 0 Then
                                            Print #fh, Format(CVDate(Bars(eBARS_DateTime, j))) _
                                                & " " & CStr(gdGetNum(hArray, j))
                                        Else
                                            Print #fh, CStr(Bars(eBARS_DateTime, j)) _
                                                & " " & CStr(gdGetNum(hArray, j))
                                        End If
                                    Next
                                End If
                            End If
                                
                            ' store into the scan's array for this symbol
                            hArray = aScanArrays.Num(i)
                            If d = kNullData Then
                                d = gdNullValue(hArray)
                            Else
                                d = d
                            End If
                            gdSetNum hArray, nPoolRec, d
                        Next
gdStopProfile 731
                    End If
                End If
                
                ' make sure date of last bar is within last 5 days
                If dLastDate >= lLastDateOfData - 5 And aScanExprW.Size > 0 Then
gdStartProfile 735
                    ' run engine to evaluate expressions for this symbol
                    aArrayOfBars.Num(0) = Weekly.BarsHandle '(in case changed)
                    aArrayOfBars.Num(1) = Weekly.BarsHandle
                    aArrayOfBars.Num(2) = Monthly.BarsHandle 'NOTE: monthly should NOT really get used in weekly scans
                    aArrayOfBars.Num(3) = Bars.BarsHandle
                    For lBars = 4 To astrBarNames.Size - 1
                        Select Case lBars
                        Case i67, i67Weekly
                            aArrayOfBars.Num(lBars) = Bars67Weekly.BarsHandle
                        Case i67Monthly
                            aArrayOfBars.Num(lBars) = Bars67Monthly.BarsHandle
                        Case Else
                            aArrayOfBars.Num(lBars) = BarsColl(lBars + 1).BarsHandle
                        End Select
                    Next lBars
                    astrParmsW.Size = 2
                    rc = RunExpressions(astrParmsW.ArrayHandle, _
                        astrBarNames.ArrayHandle, aArrayOfBars.ArrayHandle, _
                        aArrayOfResultsW.ArrayHandle, ByVal 0&, ByVal 0&)
gdStopProfile 735
                    
                    If rc <> 0 Then
                        'frmMain.Caption = CStr(rc) & " w: " & aStrParmsW(aStrParmsW.Size - 1)
                    Else
gdStartProfile 736
                        ' set current value for each expression
                        For i = 0 To aScanArraysW.Size - 1
                            ' get most recent value
                            hArray = aArrayOfResultsW.Num(i)
                            If hArray = 0 Then
                                d = kNullData
                            Else
                                If Weekly(eBARS_DateTime, Weekly.Size - 1) > lLastCompletedWeek Then
                                    d = gdGetNum(hArray, Weekly.Size - 2)
                                Else
                                    d = gdGetNum(hArray, Weekly.Size - 1)
                                End If
    
                                If fh <> 0 Then
                                    Print #fh, ""
                                    Print #fh, "Expr = " & aScanExprW(i)
                                    For j = 0 To gdGetSize(hArray) - 1
                                        If Bars(eBARS_DateTime, j) > 0 Then
                                            Print #fh, Format(CVDate(Bars(eBARS_DateTime, j))) _
                                                & " " & CStr(gdGetNum(hArray, j))
                                        Else
                                            Print #fh, CStr(Bars(eBARS_DateTime, j)) _
                                                & " " & CStr(gdGetNum(hArray, j))
                                        End If
                                    Next
                                End If
                            End If
                            
                            ' store into the scan's array for this symbol
                            hArray = aScanArraysW.Num(i)
                            If d = kNullData Then
                                d = gdNullValue(hArray)
                            Else
                                d = d
                            End If
                            gdSetNum hArray, nPoolRec, d
                        Next
                    End If
gdStopProfile 736
                End If
            End If
            
            i = 100 * nPoolRec / m.ArrayTable.NumRecords
            If i <> iPercentDone Then
                iPercentDone = i
gdStartProfile 740
                If Len(strSymbolToCalc) = 0 Then frmStatus.UpdateProgress "Criteria", iPercentDone
gdStopProfile 740
            End If
            
            'we'll do a full idle (all threads) every 3 seconds,
            'and do just a DoEvents every 1/4 second
            d = gdTickCount
            If d > dPrevFullYieldTime + 3000 Then
gdStartProfile 741
                Sleep 0
                dPrevFullYieldTime = gdTickCount
                dPrevSleepTime = dPrevFullYieldTime
gdStopProfile 741
            ElseIf d > dPrevSleepTime + 250 Then
gdStartProfile 742
                DoEvents
                dPrevSleepTime = gdTickCount
gdStopProfile 742
            End If
            
            If Len(strSymbolToCalc) = 0 Then
                If frmStatus.Status = eStatus_Aborting Then Exit For
            Else
                'Exit For
            End If
''If iPercentDone > 10 Then Exit For
        Next

gdStartProfile 749
        ' clear the expression evaluator when done with it
        If aScanExpr.Size > 0 Then
            SetupExpressions astrParms '(clear expressions)
        End If
        
        If aScanExprW.Size > 0 Then
            SetupExpressions astrParmsW '(clear expressions)
        End If
gdStopProfile 749
        
        'BenchMark Str(nCount) & " symbols"
    End If
    Set barsDow = Nothing
    
    If fh <> 0 Then
        Close #fh
        fh = 0
    End If
    
    If Len(strSymbolToCalc) = 0 Then
        If frmStatus.Status = eStatus_Aborting Then
            frmStatus.Status = eStatus_Aborted
        ElseIf frmStatus.Status <> eStatus_Aborted Then
            frmStatus.Status = eStatus_Completed
        End If
        
        If frmStatus.Status = eStatus_Completed Then bContinue = True
    Else
        bContinue = True
    End If
    
    If bContinue = True Then
    
gdStartProfile 750
        ' now get results for each Criteria
        For Each Criteria In m.Criterias
            With Criteria
                If (.IsDirty Or Len(strSymbolToCalc) > 0) And .IsActive Then
                    ' calculate its percentiles array
                    .CalcPercentiles
                    ' and put the result arrays into the pool
                    .AddToPool True
                    .IsDirty = False
                    .ToFile True
                    
                    frmQuotes.UpdateFilter "DSV:" & Criteria.ID
                End If
            End With
        Next
gdStopProfile 750
gdStartProfile 751
    
        'now recalc filters
        For i = 1 To g.SymbolPool.Filters.Count
            Set Filter = g.SymbolPool.Filters(i)
            If Filter.IsActive = True Then
                Filter.CalcFilter
                Filter.AddToPool True
                
                frmQuotes.UpdateFilter "FIL:" & Filter.ID
            End If
        Next
        Set Filter = Nothing
gdStopProfile 751
        
        frmSymbolGrid.RefreshGrid
    
        ' Save pool now so won't be lost if crashes
        If Len(strSymbolToCalc) = 0 Then
            If NumRecords > 0 Then SaveFlagGroup
            Serialize True
            DirtyCriteria = False
        End If
        RecalcDirtyCriteria = True
    End If
        
    FileFromString strFile, gdGetProfiles(700, 799), True, True
        
    d = Round((gdTickCount - dStartTime) / 60000#, 2)
    FileFromString strFile, "Total Time = " & CStr(d) & " minutes", True, True
    
    ' destroy all the temporary result arrays
    For i = 0 To aArrayOfResults.Size - 1
        gdDestroyArray aArrayOfResults(i)
    Next
    aArrayOfResults.Size = 0
    For i = 0 To aArrayOfResultsW.Size - 1
        gdDestroyArray aArrayOfResultsW(i)
    Next
    aArrayOfResultsW.Size = 0
    
    CreateCriteriaRecalcFiles nForDate
    
ErrExit:
    Set Criteria = Nothing
    Exit Function
    
ErrSection:
    Set Criteria = Nothing
    If frmStatus.Status = eStatus_Running Then frmStatus.Status = eStatus_Aborted
    RaiseError "cSymbolPool.RecalcDirtyCriteria", eGDRaiseError_Raise
    
End Function

Private Sub MakeDummyCriteria()
On Error GoTo ErrSection:

    Dim i&, nFld&, nSize&
    Dim Criteria As cCriteria
#If 0 Then
    ' load some dummy Criteria
    nSize = mDataTable.NumRecords
    
    For nFld = 1 To mObjectPool.Count
        If mObjectPool.NodeType(nFld) = "cCriteria" Then
            Set Criteria = mObjectPool(nFld)
            With Criteria
                If .StorageType = eCriteria_Boolean Then
                    For i = 0 To nSize - 1
                        mDataTable.Num(nFld, i) = gdRandomNumber(0, 1)
                    Next
                'ElseIf nFld Mod 2 = 0 Then
                '    For i = 0 To nSize - 1
                '        mDB(nFld, i) = gdRandomNumber(0, 100)
                '    Next
                Else
                    For i = 0 To nSize - 1
                        mDataTable.Num(nFld, i) = gdRandomNumber(-50, 100) * gdRandomNumber(1, 10)
                    Next
                End If
                
                .Percentile.CalcPercentiles .aValues.ArrayHandle
            End With
        End If
    Next
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolPool.MakeDummyCriteria", eGDRaiseError_Raise
    
End Sub

' For the "LIMIT" lines, e.g.: LIMIT    Module1,Module2 sym1,sym2,...
' - the symbols will be appended for any LIMIT line with a matching enablement
' - TradeNav will then separate the symbols by 4 security types (futures, stocks, forex, indices)
' - so if a limited symbol list exists for a security type, only those will be displayed
' - and if no limited symbol list exists for a security type, all will be displayed
'Version 1
'LIMIT  LimitFut    E7-,EMD-,TF-,ES-,J7-,NQ-,QC-,QG-,QH-,QI-,QM-,QO-,XC-,XK-,XS-,XW-,XY-,YM-,G6B-,G6E-,G6J-,GX-,ZB-,ZC-,ZS-,ZW-
Private Sub GetSymbolFilters(strFuturesFilter$, strStocksFilter$, strIndexFilter$, strForexFilter$)
On Error GoTo ErrSection:

    Dim i&, iPos&, strText$, strFilter$, strSymbol$
    Dim bValidFile As Boolean
    Dim aLines As New cGdArray, aFlds As New cGdArray, aRecent As New cGdArray
    
    ' read in the SymbolFilter file, decrypt, and split lines into an array
    strText = DecryptFromHex(FileToString(App.Path & "\Provided\SymbolFilter.cfg"))
    aLines.SplitFields Replace(strText, vbCrLf, Chr(27)), Chr(27)
    
    ' look through each line and build the symbol filter from each matching enablement
    ' (ignore comments, and to be valid a line must begin with "VERSION")
    For i = 0 To aLines.Size - 1
        strText = Trim(aLines(i))
        If Len(strText) > 0 And Left(strText, 1) <> "'" Then
            aFlds.SplitFields strText, vbTab
            Select Case UCase(aFlds(0))
            Case "VERSION"
                ' this line indicates a valid file (i.e. user hasn't mucked with it)
                If Val(aFlds(1)) >= 1 Then
                    bValidFile = True
                End If
            Case "LIMIT"
                If HasModule(aFlds(1)) Then
                    strFilter = strFilter & "," & aFlds(2)
                End If
            End Select
        End If
    Next
    Set aFlds = Nothing
    
    ' if they've deleted or mucked with the file, then only give them a few symbols
    If Not bValidFile Then
        strFilter = ",ES-,ZB-,IBM,MSFT,$DJIA,$SPX,$COMPQ,$EUR-USD,"
    End If
           
    ' separate the symbols by security type (Futures, Stocks, Forex, Index)
    strFuturesFilter = ","
    strStocksFilter = ","
    strIndexFilter = ","
    strForexFilter = ","
    aLines.SplitFields strFilter, ","
    aLines.Sort eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues
    For i = 0 To aLines.Size - 1
        strSymbol = UCase(Trim(aLines(i)))
        If Left(strSymbol, 1) = "$" Then
            If IsForex(strSymbol) Then
                ' a Forex symbol
                strForexFilter = strForexFilter & strSymbol & ","
            Else
                ' an Index symbol
                strIndexFilter = strIndexFilter & strSymbol & ","
            End If
        ElseIf InStr(strSymbol, "-") > 0 Then
            ' a Future
            strFuturesFilter = strFuturesFilter & strSymbol & ","
            ' also add the synthetic (if exists)
            strSymbol = ConvertSynthetic(strSymbol, True)
            If Len(strSymbol) > 0 Then
                strFuturesFilter = strFuturesFilter & strSymbol & ","
            End If
        ElseIf Len(strSymbol) > 0 Then
            ' a Stock
            strStocksFilter = strStocksFilter & strSymbol & ","
        End If
    Next
    If strFuturesFilter = "," Then strFuturesFilter = ""
    If strStocksFilter = "," Then strStocksFilter = ""
    If strIndexFilter = "," Then strIndexFilter = ""
    If strForexFilter = "," Then strForexFilter = ""

    ' but to be "safe" when the list changes, add in the recent symbols list
    ' (e.g. if a position, open order, etc. might happen to exist for a newly filtered symbol)
    If g.RealTime Is Nothing Then
        aRecent.FromFile App.Path & "\Custom\RecentSymbols.Lst"
    Else
        Set aRecent = g.RealTime.GetRecentSymbols
    End If
    For i = 0 To aRecent.Size - 1
        strSymbol = Parse(aRecent(i), vbTab, 1)
        strSymbol = Parse(strSymbol, " ", 1)
        If IsAlpha(strSymbol, 1) Then
            iPos = InStr(strSymbol, "-")
            If iPos > 0 Then
                strSymbol = Left(strSymbol, iPos)
            End If
        End If
        
        If Left(strSymbol, 1) = "$" Then
            If IsForex(strSymbol) Then
                ' a Forex symbol
                If Len(strForexFilter) > 0 And InStr(strForexFilter, "," & strSymbol & ",") = 0 Then
                    strForexFilter = strForexFilter & strSymbol & ","
                End If
            Else
                ' an Index symbol
                If Len(strIndexFilter) > 0 And InStr(strIndexFilter, "," & strSymbol & ",") = 0 Then
                    strIndexFilter = strIndexFilter & strSymbol & ","
                End If
            End If
        ElseIf InStr(strSymbol, "-") > 0 Then
            ' a Future
            If Len(strFuturesFilter) > 0 And InStr(strFuturesFilter, "," & strSymbol & ",") = 0 Then
                strFuturesFilter = strFuturesFilter & strSymbol & ","
                ' also add the synthetic (if exists)
                strSymbol = ConvertSynthetic(strSymbol, True)
                If Len(strSymbol) > 0 Then
                    strFuturesFilter = strFuturesFilter & strSymbol & ","
                End If
            End If
        ElseIf Len(strSymbol) > 0 Then
            ' a Stock
            If Len(strStocksFilter) > 0 And InStr(strStocksFilter, "," & strSymbol & ",") = 0 Then
                strStocksFilter = strStocksFilter & strSymbol & ","
            End If
        End If
    Next

ErrExit:
    Set aLines = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolPool.GetSymbolFilters"
End Sub

Private Function SetAccessFlags() As Boolean
On Error GoTo ErrSection:

    Dim strKey As String                ' Registry key
    Dim strOldDAS As String             ' Old Data Access String
    Dim strDAS As String                ' Data Access String
    Dim lTable As Long                  ' Database table
    Dim lRC As Long                     ' Return code
    Dim strSecType As String            ' Security Type
    Dim bAccess As Boolean              ' Access from the database
    Dim dLastDay As Double              ' Last day of data in the database
    Dim lExpireDate As Long             ' If last day of data before this, then consider "expired"
    Dim i&, iPos&, strSymbol$, strIgnoreIfLike$, nSymbolID&
    Dim aIfLike() As String
    Dim bVPT As Boolean
    Dim aItems As New cGdArray
    Dim bNoForex As Boolean
    Dim strFuturesFilter$, strStocksFilter$, strIndexFilter$, strForexFilter$
    Dim dChk#, iChkCount&
    Dim strSFE$
    Dim strBrokerCode As String         ' Broker code for forex symbols
    Dim strBrokerAccess As String       ' Broker access codes
    
    ' TLB 4/3/2009: read symbol filters (to only show limited symbols based on enablements)
    GetSymbolFilters strFuturesFilter, strStocksFilter, strIndexFilter, strForexFilter
    
    ' ignore Forex symbols for BetterTrades
    bVPT = HasModule("EXP,VPT")
    If (ExtremeCharts = 1 And Not bVPT) Or HasModule("NOFOREX") Then
        bNoForex = True
    End If
    
    ' get masks of symbols to ignore (from download filter)
    strIgnoreIfLike = Trim(UCase(GetIniFileProperty("IgnoreSymbols", "", "General", g.strIniFile)))
    ' TLB 8/27/2010: exclude SFE Futures unless has installed TWS
    strSFE = Trim(UCase(GetProvidedProperty("SFE")))
    If Len(strSFE) > 0 Then
        If SfeAllowed Then
            strSFE = ""
        End If
        If Len(strSFE) > 0 Then
            strIgnoreIfLike = strIgnoreIfLike & "," & strSFE
        End If
    End If
    If Len(strIgnoreIfLike) > 0 Then
        aIfLike = Split(strIgnoreIfLike, ",")
        For i = LBound(aIfLike) To UBound(aIfLike)
            aIfLike(i) = UCase(Trim(aIfLike(i)))
        Next
    Else
        ReDim aIfLike(-1 To -1) As String ' (means none)
    End If
    
    ' Build the new Data Access String
    strDAS = ""
    If Len(g.strAuthorizationString) = 0 Then
        GetAuthorizationStringFromRegistry
    End If
    aItems.SplitFields g.strAuthorizationString, ","
    For i = 0 To aItems.Size - 1
        strSecType = Trim(aItems(i))
        If Len(strSecType) > 0 And Len(strSecType) <= 2 Then
            If InStr("SFIM", Left(strSecType, 1)) > 0 Then
                If InStr(strDAS, Left(strSecType, 1)) = 0 Then
                    strDAS = strDAS & Left(strSecType, 1)
                End If
            End If
        End If
    Next
    '(added 7/23/01 by TLB)
    If strDAS = "" Then strDAS = "SFIM"
    If Not AllowDivAndMF Then
        strDAS = Replace(strDAS, "M", "")
    End If
    
    ' Get the old DAS out of the registry
    'strKey = "Software\Genesis Financial Data Services\DAS"
    'strOldDAS = GetRegistryValue(rkLocalMachine, strKey, "DAS", "")
        
    ' If the DAS has changed or we are forced to, reset the access flags
    ''If (strOldDAS <> strDAS) Or (bOnlyIfChanged = False) Then
    If True Then
        ' Save the new Data Access String
        strKey = "Software\Genesis Financial Data Services\DAS"
        SetRegistryValue rkLocalMachine, strKey, "DAS", strDAS
    
        ' Get the symbol expire from the registry
        strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
        If CLng(GetRegistryValue(rkLocalMachine, strKey, "chkSymbolExpire", vbChecked)) = 0 Then
            lExpireDate = 0
        Else
            'get 'x' # of days from last date of data
            dLastDay = LastDailyDownload
            If dLastDay <= 0 Then dLastDay = Date
            lExpireDate = dLastDay - CLng(GetRegistryValue(rkLocalMachine, strKey, "SymbolExpire", 31#))
        End If
    
        ' Open the database table with no tag
        lTable = g.Universe.tblSymbols
        If lTable = 0 Then Exit Function
        lRC = TagSelect(lTable, 0)
        
dChk = gdTickCount
        
        ' Walk through all of the records and change the access flags on only
        ' those that have changed
        lRC = d4top(lTable)
        Do While lRC = r4success
            If d4deleted(lTable) = 0 Then
                strSymbol = Trim(UCase(f4str(g.Universe.fldSymbol)))
                strSecType = Trim(UCase(f4str(g.Universe.fldSecType)))
                nSymbolID = f4long(g.Universe.fldSymbolID)
                If Len(strSymbol) = 0 Or nSymbolID = 0 Then
                    ' TLB 9/26/2007: a blank symbol or ID "shouldn't" happen but it did somehow,
                    ' so we should delete it since it can cause a problem for the DMDLL
                    d4delete lTable
                Else
                    If g.strInstalledString <> "" Then
                        If InStr(g.strInstalledString, "," & Left(strSecType, 1) & ",") = 0 Then
                            f4assignDouble g.Universe.fldEodLast, 0#
                        End If
                    End If
                    
                    dLastDay = f4double(g.Universe.fldEodLast)
                    bAccess = False
                    If nSymbolID < 0 Then
                        ' custom index (i.e. defined by this user)
                        bAccess = True
                    ElseIf InStr(strDAS, Left(strSecType, 1)) > 0 _
                            And dLastDay > 0 And dLastDay >= lExpireDate Then
                        ' ignore symbols based on download filter
                        bAccess = Not IgnoreSymbol(strSymbol, aIfLike)
                        If bAccess Then
                            ' ignore symbols not in the "filter" list (if exists for enablement)
                            iPos = InStr(strSymbol, "-")
                            If Left(strSymbol, 1) = "$" Then
                                ' an Index
                                If IsForex(strSymbol) Then
                                    ' Forex@Broker symbols require broker enablement
                                    If Mid(strSymbol, 9, 1) = "@" Then
                                        strBrokerCode = Parse(strSymbol, "@", 2)
                                        strBrokerAccess = GetProvidedProperty(UCase(strBrokerCode) & "Access", "")
'                                        If UCase(strBrokerCode) = "PFG" Then
'                                            bAccess = HasModule("B_PFG,B_CTGPFG,B_FINPFG")
'                                        ElseIf UCase(strBrokerCode) = "IB" Then
'                                            bAccess = HasModule("B_INTB,B_IDEAL")
'                                        ElseIf UCase(strBrokerCode) = "CNX" Then
'                                            bAccess = HasModule("B_ALPCNX,B_PENCNX")
                                        If Len(strBrokerAccess) > 0 Then
                                            bAccess = HasModule(strBrokerAccess)
                                        Else
                                            bAccess = HasModule("B_" & strBrokerCode)
                                        End If
                                    ElseIf bNoForex Then
                                        bAccess = False
                                    ElseIf Len(strForexFilter) > 0 Then
                                        If InStr(strForexFilter, "," & strSymbol & ",") = 0 Then
                                            bAccess = False
                                        End If
                                    End If
                                ElseIf Len(strIndexFilter) > 0 Then
                                    If InStr(strIndexFilter, "," & strSymbol & ",") = 0 Then
                                        bAccess = False
                                    End If
                                End If
                            ElseIf iPos > 0 Then
                                ' a Future
                                If Len(strFuturesFilter) > 0 Then
                                    ' Limit futures to only certain base symbols
                                    If InStr(strFuturesFilter, "," & Left(strSymbol, iPos) & ",") = 0 Then
                                        iChkCount = iChkCount + 1
                                        bAccess = False
                                    End If
                                End If
                                ' if a SpreadSymbol, limit to less than 10 contracts out
                                If IsSpreadSymbol(strSymbol) Then
                                    bAccess = False
                                    ' don't allow any spread symbols prior to version 6.5
                                    If App.Major > 6 Or App.Minor > 4 Then
                                        ' ignore spreads > 9 contracts out
                                        i = Len(strSymbol)
                                        If i > 10 Then '(length of these should never be < 11)
                                            If Mid(strSymbol, i - 2, 1) = "-" Then
                                                bAccess = True
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                ' a Stock
                                If Len(strStocksFilter) > 0 Then
                                    If InStr(strStocksFilter, "," & strSymbol & ",") = 0 Then
                                        bAccess = False
                                    End If
                                End If
                            End If
                        End If
If IsIDE Then
    'If Left(strSymbol, 4) = "G6A-" Then bAccess = False
End If
                        ' exception for Vantage Point
                        If bVPT And Not bAccess Then
                            If strSymbol = "$XAO@ASX" Or strSymbol = "$STIN@SGX" Then
                                bAccess = True
                            End If
                        End If
                    End If
                    If bAccess <> f4Logical(g.Universe.fldAccess) Then
                        f4assignLogical g.Universe.fldAccess, bAccess
                    End If
                End If
            End If
            
            lRC = d4skip(lTable, 1&)
        Loop
        
If IsIDE Then
    dChk = gdTickCount - dChk
    StartupLog "SetAccessFlags = " & Str(Int(dChk)) & " ms -- symbols hidden: " & Str(iChkCount)
End If
        
        g.strInstalledString = ""
        
        SyncCodebaseCaches

        SetAccessFlags = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.SetAccessFlags", eGDRaiseError_Raise
    
End Function

Public Sub MakeDebugFile()
On Error GoTo ErrSection:

    Dim i&, h&, iSize&
    Dim SymGroup As cSymbolGroup
    Dim Criteria As cCriteria
    Dim Filter As cFilter
    Dim aStrings As New cGdArray
    
    aStrings.Add "========== TABLE FIELDS ============="
    With m.ArrayTable
        For i = 0 To .NumFields - 1
            h = .FieldArrayHandle(i)
            If gdIsConstantValue(h) Then
                iSize = 1
            Else
                iSize = gdGetSize(h)
            End If
            aStrings.Add CStr(i) & vbTab _
                & Chr(.FieldType(i)) & vbTab _
                & iSize & vbTab _
                & .FieldName(i) & vbTab & m.aFieldIDs(i)
        Next
    End With
    
    aStrings.Add ""
    aStrings.Add "========== SYMBOL GROUPS ============="
    For i = 1 To m.SymbolGroups.Count
        Set SymGroup = m.SymbolGroups(i)
        If Not SymGroup Is Nothing Then
            With SymGroup
                aStrings.Add CStr(i) & vbTab _
                    & .ID & vbTab & .Custom & vbTab _
                    & .GroupType & vbTab & .IsActive & vbTab _
                    & .Name & vbTab & .Desc
            End With
        End If
    Next
    Set SymGroup = Nothing

    aStrings.Add ""
    aStrings.Add "========== CRITERIA ============="
    For i = 1 To m.Criterias.Count
        Set Criteria = m.Criterias(i)
        If Not Criteria Is Nothing Then
            With Criteria
                aStrings.Add CStr(i) & vbTab _
                    & .ID & vbTab & .Custom & vbTab _
                    & .IsBoolean & vbTab & .IsActive & vbTab _
                    & .Name & vbTab & .Desc
            End With
        End If
    Next
    Set Criteria = Nothing

    aStrings.Add ""
    aStrings.Add "========== FILTERS ============="
    For i = 1 To m.Filters.Count
        Set Filter = m.Filters(i)
        If Not Filter Is Nothing Then
            With Filter
                aStrings.Add CStr(i) & vbTab _
                    & .ID & vbTab & .Custom & vbTab _
                    & .Conditions.Size & vbTab & .IsActive & vbTab _
                    & .Name & vbTab & .Desc
            End With
        End If
    Next
    Set Filter = Nothing
    
    aStrings.ToFile App.Path & "\chk\Pool.chk"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolPool.MakeDebugFile", eGDRaiseError_Raise
    
End Sub

Public Function AddCustomIndex(ByVal nSymbolID&, ByVal strSymbol$) As Boolean
On Error GoTo ErrSection:

    Dim nRec&, nInsertSymbolAt&, nInsertDescAt&, i&, nDBFRec&, hArray&
    Dim aTimes As New cGdArray
    
    ' find record in Symbols DBF
    'i = TagSeek(g.Universe.tblSymbols, g.Universe.tagSymbolID, CStr(nSymbolID), False)
    i = TagSeek(g.Universe.tblSymbols, g.Universe.tagSymbol, UCase(strSymbol), True)
    If f4long(g.Universe.fldSymbolID) <> nSymbolID Then Exit Function
    nDBFRec = d4recNo(g.Universe.tblSymbols)
    
    With m.ArrayTable
        ' Find where to insert symbol
        nRec = PoolRecForSymbol(f4str(g.Universe.fldSymbol), False)
        d4go g.Universe.tblSymbols, nDBFRec
        
        ' Find where Desc gets inserted
        nInsertDescAt = PoolRecForDesc(f4str(g.Universe.fldDesc), False)
        nInsertDescAt = .Num(kDesc, nInsertDescAt)
        d4go g.Universe.tblSymbols, nDBFRec
        
        ' insert at this row of the pool
        For i = 0 To .NumFields - 1
            hArray = .FieldArrayHandle(i)
            If hArray <> 0 And Not gdIsConstantValue(hArray) Then
                gdInsertNum hArray, gdNullValue(hArray), nRec
            End If
        Next
        .NumRecords = .NumRecords + 1
        
        ' Fix symbol (consecutive numbers since already in symbol order)
        ' and desc (bump others up by one)
        For i = 0 To .NumRecords - 1
            .Num(kSymbol, i) = i
            If .Num(kDesc, i) >= nInsertDescAt Then
                .Num(kDesc, i) = .Num(kDesc, i) + 1
            End If
        Next
        .Num(kDesc, nRec) = nInsertDescAt

        ' set rest of builtin fields
        .Num(kSymbolID, nRec) = nSymbolID
        .Num(kDbRecNum, nRec) = nDBFRec
        .Num(kAllSymbols, nRec) = 1
        .Num(kAllIndices, nRec) = 1
        .Num(kFirstDate, nRec) = f4long(g.Universe.fldEodFirst)
        .Num(kLastDate, nRec) = f4long(g.Universe.fldEodLast)
        .Num(kSecType, nRec) = eSYMType_Index
        
        ' resort SymIndex
        i = gdSortAsIndex(.FieldArrayHandle(kSymIndex), .FieldArrayHandle(kSymbolID), True, eGdSort_Default, 0, -1)
    End With

    ' reload all groups, criteria, and filters (to reattach arrays)
    LoadSymbolGroups False
    LoadCriterias False
    LoadFilters

    ' recalc criteria just for this symbol
    RecalcDirtyCriteria False, strSymbol
    
    AddCustomIndex = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.AddCustomIndex", eGDRaiseError_Raise
    
End Function

Public Function RemoveCustomIndex(ByVal nSymbolID&) As Boolean
On Error GoTo ErrSection:

    Dim nRec&, i&, hArray&, bSuccess As Boolean, strSymbol$

    With m.ArrayTable
        ' find rec in pool
        nRec = PoolRecForSymbolID(nSymbolID)
        strSymbol = Symbol(nRec)
        If nRec >= 0 Then
            bSuccess = True
            ' delete this record in the pool
            For i = 0 To .NumFields - 1
                hArray = .FieldArrayHandle(i)
                If hArray <> 0 And Not gdIsConstantValue(hArray) Then
                    gdDeleteItems hArray, nRec, 1
                End If
            Next
            .NumRecords = .NumRecords - 1
        
            ' resort SymIndex
            i = gdSortAsIndex(.FieldArrayHandle(kSymIndex), .FieldArrayHandle(kSymbolID), True, eGdSort_Default, 0, -1)
        End If
    End With

    ' reload all groups, criteria, and filters (to reattach arrays)
    LoadSymbolGroups False
    LoadCriterias False
    LoadFilters
    
    ' recalc criteria just for this symbol (to redo percentiles, etc.)
    RecalcDirtyCriteria False, strSymbol
    
    RemoveCustomIndex = bSuccess

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.RemoveCustomIndex", eGDRaiseError_Raise
    
End Function

Public Sub SaveFlagGroup()
On Error GoTo ErrSection:

    Dim SymbolGroup As New cSymbolGroup
    Dim alSymbolID As New cGdArray
    Dim lField As Long
    Dim lIDField As Long
    Dim lIndex As Long
    Dim lCounter As Long
    
    lField = FieldNumForID("GRP:_FLAGS_.GRP")
        
    If lField >= 0 Then
        alSymbolID.Create eGDARRAY_Longs, ArrayTable.FieldArray(lField).CountOf(1)
        lIndex = 0&
        For lCounter = 0 To ArrayTable.FieldArray(lField).Size - 1
            If ArrayTable(lField, lCounter) = 1 Then
                alSymbolID(lIndex) = SymbolID(lCounter)
                lIndex = lIndex + 1
            End If
        Next lCounter
        
        Set SymbolGroup = PoolObject("GRP:_FLAGS_.GRP")
        ''SymbolGroup.Edit AddSlash(App.Path) & "Custom", "_FLAGS_.GRP", , , alSymbolID, False, False
        SymbolGroup.Save alSymbolID, False
    End If
    
ErrExit:
    Set alSymbolID = Nothing
    Set SymbolGroup = Nothing
    Exit Sub
    
ErrSection:
    Set alSymbolID = Nothing
    Set SymbolGroup = Nothing
    RaiseError "cSymbolPool.SaveFlagGroup", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FundamentalOnly
'' Description: If this is a fundamental only criteria, return the kind id
'' Inputs:      English text of the Criteria
'' Returns:     Kind ID of the Fundamental or 0 if not fundamental only
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FundamentalOnly(ByVal strEnglish As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long
    Dim strKind As String
    
    If Left(UCase(strEnglish), 19) = "GET CURRENT VALUE (" Then
        If Right(strEnglish, 1) = ")" And InStr(Left(strEnglish, Len(strEnglish) - 1), ")") = 0 Then
            strKind = Parse(strEnglish, Chr(34), 2)
            lReturn = DM_GetKindIDFromName(strKind)
        End If
    End If
    
    FundamentalOnly = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.FundamentalOnly", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolInField
'' Description: Determine whether a symbol is in a group/filter/criteria
'' Inputs:      Symbol/Symbol ID, Field ID
'' Returns:     True if in Field, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SymbolInField(ByVal vSymbol As Variant, ByVal strFieldID As String) As Boolean
On Error GoTo ErrSection:

    Dim lSymbolID As Long               ' Symbol ID for the symbol/ID passed in
    Dim lPoolRec As Long                ' Pool record for the given symbol ID
    Dim lFieldNum As Long               ' Field number for the given group ID
    
    lSymbolID = GetSymbolID(vSymbol)
    lPoolRec = PoolRecForSymbolID(lSymbolID)
    lFieldNum = FieldNumForID(strFieldID)
    
    SymbolInField = (ArrayTable.Num(lFieldNum, lPoolRec) = 1)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.SymbolInField", eGDRaiseError_Raise
    
End Function

Private Function GoToSymbolRec(ByVal Symbol) As Boolean
On Error GoTo ErrSection:
    
    Dim nRecord&, nRec&
    nRecord = Me.PoolRecForSymbolID(GetSymbolID(Symbol))
    If nRecord >= 0 Then
        nRec = m.ArrayTable.Num(kDbRecNum, nRecord)
        If GotoRec(g.Universe.tblSymbols, nRec) Then
            GoToSymbolRec = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.GoToSymbolRec", eGDRaiseError_Raise
End Function

Public Property Get TickFirstDate(ByVal Symbol) As Long
On Error GoTo ErrSection:
    
    'Note: the date in the Symbols table is not consistenly maintained ...
    #If 0 Then
        If GoToSymbolRec(Symbol) Then
            TickFirstDate = f4long(g.Universe.fldTickFirst)
        End If
    #Else
        ' ... so must use an alternate method
        TickFirstDate = GetFirstTickDate(Symbol)
    #End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cSymbolPool.TickFirstDate", eGDRaiseError_Raise
End Property

Public Property Get TickLastDate(ByVal Symbol) As Long
On Error GoTo ErrSection:
    
    If GoToSymbolRec(Symbol) Then
        TickLastDate = f4long(g.Universe.fldTickLast)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cSymbolPool.TickLastDate", eGDRaiseError_Raise
End Property

Public Property Get EodFirstDate(ByVal Symbol) As Long
On Error GoTo ErrSection:
    
    If GoToSymbolRec(Symbol) Then
        EodFirstDate = f4long(g.Universe.fldEodFirst)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cSymbolPool.EodFirstDate", eGDRaiseError_Raise
End Property

Public Property Get EodLastDate(ByVal Symbol) As Long
On Error GoTo ErrSection:
    
    If GoToSymbolRec(Symbol) Then
        EodLastDate = f4long(g.Universe.fldEodLast)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cSymbolPool.EodLastDate", eGDRaiseError_Raise
End Property

Private Function IgnoreSymbol(ByVal strSymbol$, aLike() As String) As Boolean
On Error GoTo ErrSection:

    Dim i&, iPos&, strLike$
        
    If UBound(aLike) >= 0 Then
        For i = LBound(aLike) To UBound(aLike)
            strLike = aLike(i)
            If Len(strLike) > 0 Then
                If Left(strLike, 1) = "!" Then
                    IgnoreSymbol = True ' default to true
                    For iPos = 2 To Len(strLike)
                        If InStr(strSymbol, Mid(strLike, iPos, 1)) > 0 Then
                            IgnoreSymbol = False ' is not a match
                            Exit For
                        End If
                    Next
                    If IgnoreSymbol Then Exit Function
                ElseIf strSymbol Like strLike Then
                    IgnoreSymbol = True ' is a match, so ignore it
                    Exit Function
                End If
            End If
        Next
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.IgnoreSymbol", eGDRaiseError_Raise
End Function


Public Function GetSymbolsForGroup(ByVal vFieldIdOrNum As Variant) As String
On Error GoTo ErrSection:

    Dim i&, nFldID&, nCount&, nSize&, hOnOff&, strGroupID$
    Dim aOnOff As cGdArray
    Dim mb As New cMemBuffer
     
    ' get the On/Off array for this group/criteria/filter
    If VarType(vFieldIdOrNum) = vbString Then
        ' e.g. "GRP:ALL STOCKS.GRP"
        strGroupID = vFieldIdOrNum
        If InStr(strGroupID, ".") = 0 Then
            strGroupID = strGroupID & ".GRP"
        End If
        If InStr(strGroupID, ":") = 0 Then
            strGroupID = "GRP:" & strGroupID
        End If
        nFldID = g.SymbolPool.FieldNumForID(strGroupID)
    Else
        nFldID = Val(vFieldIdOrNum)
    End If
    Set aOnOff = g.SymbolPool.ArrayTable.FieldArray(nFldID)
    If Not aOnOff Is Nothing Then
        ' pre-allocate the memory buffer so all allocated at once
        mb.Allocated = aOnOff.CountOfNonNull * 10
        nCount = 0
        nSize = aOnOff.Size
        hOnOff = aOnOff.ArrayHandle
        ' for each symbol which is in this group
        For i = 0 To nSize - 1
            If gdGetNum(hOnOff, i) > 0 Then ' aOnOff(i) > 0 Then
                ' append the symbol into the memory buffer
                If nCount > 0 Then
                    mb.PutStr vbTab & Me.Symbol(i)
                Else
                    mb.PutStr Me.Symbol(i)
                End If
                nCount = nCount + 1
            End If
        Next
        ' return the memory buffer (big string)
        GetSymbolsForGroup = mb.GetStr
    End If

ErrExit:
    Set aOnOff = Nothing
    Set mb = Nothing
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.GetSymbolsForGroup"

End Function

' For the "Criteria Server" to create files for the web Stock Screener to use ...
Public Function CreateCriteriaRecalcFiles(Optional ByVal nForDate& = 0) As Long
On Error GoTo ErrSection:

    Dim s$, strSymbol$, strPath$, strDatePath$, strType$, strDesc$
    Dim i&, iFld&, iRec&, d#, iCount&, nSymbolID&, dNull#, nDataStart&
    Dim fSingle As Single
    Dim fhInfo%, fhData%
    Dim aSymbols As New cGdArray
    Dim Criteria As cCriteria
    Dim Bars As New cGdBars

    ' get path from flag file
    aSymbols.FromFile App.Path & "\CriteriaRecalcFiles.flg"
    strPath = Trim(aSymbols(0))
    If Len(strPath) = 0 Then
        Exit Function
    End If
    nDataStart = JulFromLong(Val(aSymbols(1)))
    aSymbols.Size = 0
    strPath = AddSlash(strPath)
    MakeDir strPath
    MakeDir strPath & "Data\"
    KillFile strPath & "*.*"
    
    ' determine which date
    If nForDate >= 19000000 Then
        nForDate = JulFromLong(nForDate)
    End If
    If nForDate = 0 Then
        nForDate = LastDailyDownload
    ElseIf nForDate > LastDailyDownload Or Abs(nForDate) <= 2 Then
        nForDate = LastDailyDownload + 1
        Do While Not IsWeekday(nForDate)
            nForDate = nForDate + 1
        Loop
    End If

    fhInfo = FreeFile
    Open strPath & "CriteriaInfo.TXT" For Output As #fhInfo
    
    ' write out the info for each criteria
    s = "Fld#" & vbTab & "ID" & vbTab & "Name" & vbTab & "Description" _
        & vbTab & "EnablementCodes" & vbTab & "CategoryID" & vbTab & "CategoryName" _
        & vbTab & "IsBoolean"
    If fhInfo <> 0 Then Print #fhInfo, s
    For iFld = 0 To m.aFieldIDs.Size - 1
        s = m.aFieldIDs(iFld)
        strType = Parse(s, ":", 1)
        If strType = "DSV" Or strType = "DSP" Then
            Set Criteria = Criterias(Mid(s, 5))
            If Not Criteria Is Nothing Then
                If Criteria.IsActive Then
                    iCount = iCount + 1
                    frmTest.AddList Str(iFld) & vbTab & s
                    dNull = m.ArrayTable.FieldArray(iFld, False).NullValue
                    If strType <> "DSP" Then
                        s = Str(iFld) & vbTab & Criteria.ID & vbTab & Criteria.Name & vbTab & Criteria.Desc _
                            & vbTab & Criteria.Required & vbTab & Criteria.CategoryID & vbTab & Criteria.CategoryName _
                            & vbTab & Str(Criteria.IsBoolean)
                        If fhInfo <> 0 Then Print #fhInfo, s
                    End If
                
                    ' write out the values for each symbol
                    fhData = FreeFile
                    If strType = "DSP" Then
                        s = strPath & FileBase(Criteria.ID) & ".PCT"
                    Else
                        s = strPath & FileBase(Criteria.ID) & ".DAT"
                    End If
                    Open s For Binary As #fhData
                    For iRec = 0 To m.ArrayTable.NumRecords - 1
                        nSymbolID = SymbolID(iRec)
                        If nSymbolID > 0 And SecType(iRec) = eSYMType_Stock Then
                            strSymbol = Symbol(iRec)
                            If InStr(strSymbol, "@") = 0 And Len(strSymbol) > 0 Then
                                If iCount = 1 Then
                                    strDesc = Desc(iRec)
                                    aSymbols.Add strSymbol & vbTab & Str(nSymbolID) & vbTab & strDesc
                                End If
                                d = m.ArrayTable.Num(iFld, iRec)
                                If d = dNull Then
                                    s = ""
                                    fSingle = -999999
                                Else
                                    d = RoundToSigDigits(d, 7)
                                    fSingle = d
                                    's = Str(d)
                                End If
                                If fhData <> 0 Then
                                    Put #fhData, , fSingle
                                End If
                            End If
                        End If
                    Next
                    If fhData Then
                        Close #fhData
                        fhData = 0
                    End If
                    DoEvents
                End If
            End If
        End If
    Next
    Set Criteria = Nothing
    frmTest.AddList Str(iCount) & " fields"
    If fhInfo <> 0 Then Close #fhInfo
    aSymbols.ToFile strPath & "Symbols.TXT"
    
    ' copy all the files into a folder for that specific date (to keep history)
    strDatePath = strPath & Format(nForDate, "YYYYMMDD") & "\"
    MakeDir strDatePath
    KillFile strDatePath & "*.*"
    FileCopy strPath & "*.*", strDatePath
    
    ' write data for past "n" days of each symbol to a file
    If nDataStart > 0 And nForDate >= LastDailyDownload Then
        For i = 0 To aSymbols.Size - 1
            nSymbolID = Val(Parse(aSymbols(i), vbTab, 2))
            DM_GetBars Bars, nSymbolID, , nDataStart, nForDate, , , , (nForDate > LastDailyDownload)
            If Bars.Size > 0 Then
                Bars.ToFile "GDB", strPath & "Data\" & Str(nSymbolID) & ".GDB"
            End If
            If i Mod 100 = 0 Then
                frmTest.AddList Str(i) & vbTab & aSymbols(i)
                DoEvents
            End If
        Next
    End If
    
    
ErrExit:
    CreateCriteriaRecalcFiles = iCount
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.CreateCriteriaRecalcFiles"
End Function

Public Function OtherFieldsStart()
    OtherFieldsStart = kOtherFields
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NumberRecordsForID
'' Description: Determine the number of records for the given field ID
'' Inputs:      Field ID
'' Returns:     Number of records
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NumberRecordsForID(ByVal strFieldID As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lFieldNum As Long               ' Field number for the group in the symbol pool

    lReturn = 0
    lFieldNum = g.SymbolPool.FieldNumForID(strFieldID)
    If lFieldNum >= 0 Then
        lReturn = ArrayTable.FieldArray(lFieldNum).CountOf(1)
    End If
    
    NumberRecordsForID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolPool.NumberItemsForID"
    
End Function
