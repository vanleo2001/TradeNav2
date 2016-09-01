VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSymbolGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum eSymbolGroupType
    eGROUP_Normal = 0
    eGROUP_Builtin = 1
    eGROUP_QuoteList = 2
    eGROUP_Flag = 3
End Enum

Private Type mPrivate
    strID As String                     ' Filename (without path)
    strName As String                   ' Displayed name
    strDesc As String                   ' Full description
    
    bCustom As Boolean                  ' Is this a custom symbol group?
    eGroupType As eSymbolGroupType      ' Symbol group type
    
    aSymbolIds As cGdArray              ' Array of symbol ID's
    astrSymbols As cGdArray             ' Array of symbols without symbol ID's
    aResults As cGdArray                ' Pool array (T/F for each symbol)
    
    ' Custom Index stuff added 11/8/2001 by DAJ
    lSymID As Long                      ' Symbol ID for the custom index
    bIndex As Boolean                   ' Is this symbol group also a custom index?
    adPriceWeights As cGdArray          ' Weights on the prices for a custom index
    adVolWeights As cGdArray            ' Weights on the volumes for a custom index
    alFlags As cGdArray                 ' Special flags for a custom index
    dPriceDivisor As Double             ' Total Price Divisor for a custom index
    dVolDivisor As Double               ' Total Volume Divisor for a custom index
    
    bIsActive As Boolean                ' Is the symbol group active?
    strRequired As String               ' Module code required for the symbol group to show
End Type
Private m As mPrivate

Private Sub Class_Initialize()
    
    Set m.aResults = New cGdArray
    m.aResults.Create eGDARRAY_TinyInts, 0, 0
    m.aResults.MakeConstantArray m.aResults.NullValue
    
    Set m.aSymbolIds = New cGdArray
    m.aSymbolIds.Create eGDARRAY_Longs
    
    Set m.astrSymbols = New cGdArray
    m.astrSymbols.Create eGDARRAY_Strings
    
    ' Custom Index stuff added 11/8/2001 by DAJ
    m.lSymID = 0&
    Set m.adPriceWeights = New cGdArray
    m.adPriceWeights.Create eGDARRAY_Doubles
    Set m.adVolWeights = New cGdArray
    m.adVolWeights.Create eGDARRAY_Doubles
    Set m.alFlags = New cGdArray
    m.alFlags.Create eGDARRAY_Longs
    
    m.bCustom = True 'default
    m.eGroupType = eGROUP_Normal
    m.bIsActive = True
    m.bIndex = False
    m.dPriceDivisor = 1#
    m.dVolDivisor = 1#
    
End Sub

Private Sub Class_Terminate()
    
    Set m.aResults = Nothing
    Set m.aSymbolIds = Nothing
    Set m.astrSymbols = Nothing
    
    ' Custom Index stuff added 11/8/2001 by DAJ
    Set m.adPriceWeights = Nothing
    Set m.adVolWeights = Nothing
    Set m.alFlags = Nothing

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

Public Property Get Custom() As Boolean
    Custom = m.bCustom
End Property
Public Property Let Custom(ByVal bIsCustom As Boolean)
    m.bCustom = bIsCustom
End Property

Public Property Get GroupType() As eSymbolGroupType
    GroupType = m.eGroupType
End Property
Public Property Let GroupType(ByVal eValue As eSymbolGroupType)
    m.eGroupType = eValue
End Property

Public Property Get SymbolIDs() As cGdArray
    Set SymbolIDs = m.aSymbolIds
End Property
Public Property Get Symbols() As cGdArray
    Set Symbols = m.astrSymbols
End Property
Public Property Get PriceWeights() As cGdArray
    Set PriceWeights = m.adPriceWeights
End Property
Public Property Get VolWeights() As cGdArray
    Set VolWeights = m.adVolWeights
End Property
Public Property Get Flags() As cGdArray
    Set Flags = m.alFlags
End Property

Public Property Get IsActive() As Boolean
    IsActive = m.bIsActive And HasModule(m.strRequired)
End Property
Public Property Let IsActive(ByVal bIsActive As Boolean)
    m.bIsActive = bIsActive
End Property

Public Property Get IsIndex() As Boolean
    IsIndex = m.bIndex
End Property
Public Property Let IsIndex(ByVal bIsIndex As Boolean)
    m.bIndex = bIsIndex
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.lSymID
End Property
Public Property Let SymbolID(ByVal lSymbolID As Long)
    m.lSymID = lSymbolID
End Property

Public Property Get PriceDivisor() As Double
    PriceDivisor = m.dPriceDivisor
End Property
Public Property Let PriceDivisor(ByVal dPriceDivisor As Double)
    m.dPriceDivisor = dPriceDivisor
End Property

Public Property Get VolDivisor() As Double
    VolDivisor = m.dVolDivisor
End Property
Public Property Let VolDivisor(ByVal dVolDivisor As Double)
    m.dVolDivisor = dVolDivisor
End Property

Public Property Get Required() As String
    Required = m.strRequired
End Property

Public Sub AddToPool(Optional ByVal bReplaceArrays As Boolean = False)
On Error GoTo ErrSection:

    Dim strFldID$, nFld&, i&, nRecNum&

    ' add object to collection
    m.strID = UCase(m.strID)
    With g.SymbolPool.SymbolGroups
        i = .Index(m.strID)
        If i > 0 Then
            .Item(i) = Me
        Else
            .Add Me, m.strID
        End If
    End With
    
    ' must build a "list" array (T/F for each symbol) for pool
    If bReplaceArrays Then
        'size results array
        m.aResults.Clear '(to clear ConstValue status)
        m.aResults.Size = g.SymbolPool.NumRecords
        For i = 0 To m.aSymbolIds.Size - 1
            nRecNum = g.SymbolPool.PoolRecForSymbolID(m.aSymbolIds(i))
            If nRecNum >= 0 Then
                m.aResults.Num(nRecNum) = 1
            End If
        Next
    End If
    
    ' add Results array to ArrayTable
    With g.SymbolPool
        strFldID = "GRP:" & m.strID
        nFld = .FieldNumForID(strFldID)
        If nFld < 0 Or bReplaceArrays Then
            nFld = .ArrayTable.AttachField(m.aResults, nFld, m.strName)
            .FieldID(nFld) = strFldID
        Else 'If nFld >= 0 Then
            .ArrayTable.FieldName(nFld) = m.strName
        End If
    End With
    
    m.bIsActive = True
    If nFld < 0 Then
        ' TLB 10/18/2011: this is probably obsolete -- don't even think we should ever need to do this??
        nFld = g.SymbolPool.ArrayTable.FieldNum(m.strName)
    End If
    If nFld = -1 Then
        ' should never get here??
        m.bIsActive = False
    ElseIf m.eGroupType <> eGROUP_Builtin And m.eGroupType <> eGROUP_Flag Then
        If g.SymbolPool.ArrayTable.FieldArray(nFld).CountOf(1) = 0 Then
            m.bIsActive = False
        End If
    ElseIf m.eGroupType = eGROUP_Builtin Then
        If g.SymbolPool.ArrayTable.FieldArray(nFld).CountOf(1) = 0 Then
            m.bIsActive = False
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolGroup.AddToPool", eGDRaiseError_Raise

End Sub

'To make a Special Group Type:
'- Builtin: for builtin lists (e.g. "All Stocks")
'- QuoteList:
Public Sub MakeSpecialType(ByVal strID$, ByVal strName$, _
        Optional ByVal eGroupType As eSymbolGroupType = eGROUP_Builtin)
On Error GoTo ErrSection:

    m.eGroupType = eGroupType
    m.strName = strName
    m.strID = UCase(strID)
    Select Case eGroupType
    Case eGROUP_QuoteList
        m.bCustom = True
        AddSymbolID 50    ' $DJIA
        AddSymbolID 27    ' $COMPQ
        AddSymbolID 207   ' $SPX
        AddSymbolID 11936 ' IBM
        AddSymbolID 41180 ' SP-067
        AddSymbolID 41183 ' TQ-067
    Case eGROUP_Flag
        m.bCustom = True
    Case Else
        m.bCustom = False
    End Select
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolGroup.MakeSpecialType", eGDRaiseError_Raise

End Sub

' Load from file
Public Function FromFile(ByVal strPath$, ByVal strFile$, _
        Optional bLoadSymbols = False) As Boolean
On Error GoTo ErrSection:
    
    Dim iPos&, strParm$, strText$, nSymbolID&, strSymbol$
    Dim bResave As Boolean, fh As Integer
    Dim Bars As New cGdBars

#If 0 Then
    ' check if just wanting to reload the current file
    If Len(strPath) = 0 And Len(strFile) = 0 And Len(m.strID) > 0 Then
        strFile = m.strID
        If m.bCustom Then
            strPath = App.Path & "\Custom\"
        Else
            strPath = App.Path & "\Provided\"
        End If
    End If
#End If

    ' determine if Provided or Custom
    If Len(strPath) = 0 Then
        strPath = App.Path & "\Provided\"
        If Not FileExist(strPath & strFile) Then
            strPath = App.Path & "\Custom\"
        End If
    End If

    m.aSymbolIds.Clear
    m.astrSymbols.Clear
    
    ' Custom Index stuff added 11/8/2001 by DAJ
    m.adPriceWeights.Clear
    m.adVolWeights.Clear
    m.alFlags.Clear

    m.strID = UCase(strFile)
    strPath = AddSlash(UCase(strPath))
    If Right(strPath, 8) = "\CUSTOM\" Then
        m.bCustom = True
    Else
        m.bCustom = False
    End If
    strFile = strPath & strFile

    ' to remove obsolete files
    If FileLength(strFile) < 5 Then
        KillFile strFile
    Else
        fh = FreeFile
        Open strFile For Input As #fh
        Do While Not EOF(fh)
            Line Input #fh, strText
            strText = Trim(strText)
            iPos = InStr(strText, "=")
            If iPos > 1 Then
                strParm = UCase(Trim(Left(strText, iPos - 1)))
                strText = Parse(Mid(strText, iPos + 1), vbTab, 1)
                Select Case strParm
                    Case "NAME":
                        m.strName = strText
                    Case "DESC":
                        m.strDesc = strText
                    Case "GROUPTYPE":
                        m.eGroupType = ValOfText(strText)
                    'Case "ISACTIVE":
                    '    m.bIsActive = ValOfText(strText)
                    
                    ' Custom Index stuff added 11/8/2001 by DAJ
                    Case "ISINDEX":
                        m.bIndex = CBool(CLng(strText))
                    Case "PRICEDIVISOR":
                        m.dPriceDivisor = ValOfText(strText)
                        If m.dPriceDivisor = 0# Then m.dPriceDivisor = 1#
                    Case "VOLDIVISOR":
                        m.dVolDivisor = ValOfText(strText)
                        If m.dVolDivisor = 0# Then m.dVolDivisor = 1#
                    Case "REQUIRED":
                        m.strRequired = strText
                End Select
            ElseIf Not bLoadSymbols Then
                'if don't need to load symbols, stop reading now
                Exit Do
            ElseIf Len(strText) > 0 Then
                'see if a Symbol or a SymbolID
                strSymbol = Parse(strText, vbTab, 1)
                If IsAlpha(strSymbol) Then
                    ' set bars prop so will use gdSymbol class to convert to
                    ' Genesis symbology (stock classes, contract century, etc.)
                    Bars.Prop(eBARS_Symbol) = strSymbol
                    strSymbol = Bars.Prop(eBARS_Symbol)
                    'see if symbol exists (if so, flag to resave SymbolID's)
                    nSymbolID = GetSymbolID(strSymbol)
                    If nSymbolID <> 0 Then bResave = True
                Else
                    nSymbolID = ValOfText(strSymbol)
                    
                    ' temporary fix for goof-up
                    If nSymbolID = 0 Then
                        strSymbol = Parse(strText, vbTab, 2)
                        If InStr(strSymbol, " ") = 0 Then
                            If Left(strSymbol, 2) = "_I" Then
                                '(fix for index and forex symbols)
                                strSymbol = "$" & Mid(strSymbol, 3)
                                If Mid(strSymbol, 5, 2) = "_P" Then
                                    strSymbol = Left(strSymbol, 4) & "-" & Mid(strSymbol, 7)
                                End If
                            End If
                            Bars.Prop(eBARS_Symbol) = strSymbol
                            strSymbol = Bars.Prop(eBARS_Symbol)
                            'see if symbol exists (if so, flag to resave SymbolID's)
                            nSymbolID = GetSymbolID(strSymbol)
                            If nSymbolID <> 0 Then bResave = True
                        End If
                    End If
                End If
                
                If nSymbolID <> 0 Then
                    m.aSymbolIds.Add nSymbolID
                
                    ' Custom Index stuff added 11/8/2001
                    ' Keep the weight arrays in sync with the SymbolID array
                    If Parse(strText, vbTab, 3) = "" Then
                        m.adPriceWeights.Add 1#
                    Else
                        m.adPriceWeights.Add CDbl(ValOfText((Parse(strText, vbTab, 3))))
                    End If
                    
                    If Parse(strText, vbTab, 4) = "" Then
                        m.adVolWeights.Add 1#
                    Else
                        m.adVolWeights.Add CDbl(ValOfText((Parse(strText, vbTab, 4))))
                    End If

                    m.alFlags.Add CLng(ValOfText((Parse(strText, vbTab, 5))))
                Else
                    strSymbol = Parse(strText, vbTab, 2)
                    If Len(strSymbol) > 0 Then
                        If InStr(strSymbol, "(") <> 0 Then
                            'backward-compatibility: strip off "(SO)" and insert space
                            strSymbol = Parse(strSymbol, "(", 1)
                            If Len(strSymbol) >= 3 Then
                                strSymbol = Left(strSymbol, Len(strSymbol) - 2) & " " & Right(strSymbol, 2)
                            End If
                        End If
                        m.astrSymbols.Add strSymbol
                    End If
                End If
            End If
            FromFile = True
        Loop
        Close #fh
    
        If m.eGroupType = eGROUP_Flag Then m.strName = "FLAGGED Symbols"

        If bResave Then ToFile
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolGroup.FromFile", eGDRaiseError_Raise

End Function

' Save to file
Public Sub ToFile()
On Error GoTo ErrSection:

    Dim aLines As New cGdArray
    Dim strPath$, strFile$, strSymbol$, nSymbolID&, i&, iSorted&
    Dim SymInfo As vbSymbolInfo
    Dim aSorted As New cGdArray
    
    strFile = m.strID
    If Len(strFile) = 0 Then
        strFile = NewCustomObjectName(".GRP")
        m.bCustom = True
        m.strID = UCase(strFile)
    End If
    strPath = App.Path & "\"
    If m.bCustom Then
        strPath = strPath & "Custom\"
    Else
        strPath = strPath & "Provided\"
    End If
    
    With aLines
        .Add "Name=" & Parse(m.strName, vbTab, 1)
        .Add "Desc=" & Parse(m.strDesc, vbTab, 1)
        .Add "Required=" & Parse(m.strRequired, vbTab, 1)
        .Add "GroupType=" & CStr(m.eGroupType)
        '.Add "IsActive=" & CStr(m.bIsActive)
        
        ' Custom Index stuff added 11/8/2001 by DAJ
        .Add "IsIndex=" & CStr(CLng(m.bIndex))
        .Add "PriceDivisor=" & CStr(m.dPriceDivisor)
        .Add "VolDivisor=" & CStr(m.dVolDivisor)
        
        ' Write the option symbols (which have no SymbolID)
        m.astrSymbols.Sort eGdSort_IgnoreCase
        For i = 0 To m.astrSymbols.Size - 1
            .Add "0" & vbTab & m.astrSymbols(i)
        Next i
        
        ' Write all the symbols (SymbolID is what's
        ' important - Symbol written just for reference)
        ' (first sort the symbols by ID using an "index")
        aSorted.Create eGDARRAY_Longs
        gdSortAsIndex aSorted.ArrayHandle, m.aSymbolIds.ArrayHandle, 1, 0, 0, m.aSymbolIds.Size
        For iSorted = 0 To aSorted.Size - 1
            i = aSorted.Num(iSorted)
            nSymbolID = m.aSymbolIds(i)
            strSymbol = g.SymbolPool.SymbolForID(nSymbolID)
            If Len(strSymbol) = 0 Then
                'if not in pool, see if in DBF
                If SU_GetSymbolInf(nSymbolID, SymInfo) Then
                    strSymbol = SymInfo.Symbol
                End If
            End If
            .Add CStr(nSymbolID) & vbTab & strSymbol & vbTab & _
                m.adPriceWeights(i) & vbTab & m.adVolWeights(i) & vbTab & m.alFlags(i)
        Next
        
        .ToFile strPath & strFile
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolGroup.ToFile", eGDRaiseError_Raise

End Sub

Public Function RemoveSymbol(ByVal strPath$, ByVal strID$, ByVal Symbol As Variant) As Boolean
On Error GoTo ErrSection:

    Dim lSymbolID As Long
    Dim strSymbol As String
    Dim lPos As Long

    ' Load from file (if existing), including symbols
    If Len(strID) > 0 Then
        If Not FromFile(strPath, strID, True) Then
            InfBox "i=[] ; h=ERROR ; Could not be loaded:|" & strID
            Exit Function
        End If
    Else
        m.strID = ""
    End If
    
    If VarType(Symbol) = vbString Then
        lSymbolID = g.SymbolPool.SymbolIDforSymbol(Symbol)
        strSymbol = Symbol
    Else
        lSymbolID = Symbol
        strSymbol = g.SymbolPool.SymbolForID(lSymbolID)
    End If
    
    If lSymbolID = 0 Then
        If m.astrSymbols.BinarySearch(strSymbol, lPos) Then
            m.astrSymbols.Remove lPos
            RemoveSymbol = True
        End If
    Else
        If m.aSymbolIds.BinarySearch(lSymbolID, lPos) Then
            m.aSymbolIds.Remove lPos
            RemoveSymbol = True
        End If
    End If
    
    If RemoveSymbol Then
        ' Save to file
        ToFile
        
        ' Add back into pool
        AddToPool True
        
        ' Refresh symbol grid dropdown and list
        frmSymbolGrid.RefreshGrid
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolGroup.RemoveSymbol", eGDRaiseError_Raise

End Function

Public Function AddSymbols(ByVal strPath$, ByVal strID$, astrSymbols As cGdArray) As Boolean
On Error GoTo ErrSection:

    Dim lSymbolID As Long
    Dim strSymbol As String
    Dim lPos As Long
    Dim lIndex As Long
    Dim lNumRows As Long
    Dim lCurRow As Long

    ' Load from file (if existing), including symbols
    If Len(strID) > 0 Then
        If Not FromFile(strPath, strID, True) Then
            InfBox "i=[] ; h=ERROR ; Could not be loaded:|" & strID
            Exit Function
        End If
    Else
        m.strID = ""
    End If
    
    If m.eGroupType = eGROUP_QuoteList Then
        If astrSymbols.Size + m.astrSymbols.Size + m.aSymbolIds.Size > MaxSymbolsAllowed Then
            Err.Raise vbObjectError + 1000, , "There cannot be more than " & MaxSymbolsAllowed & " symbols on the quote board"
        End If
    End If
    
    For lIndex = 0 To astrSymbols.Size - 1
        lSymbolID = GetSymbolID(astrSymbols(lIndex))
        
        If lSymbolID = 0 Then
            If Not m.astrSymbols.BinarySearch(astrSymbols(lIndex), lPos) Then
                m.astrSymbols.Add astrSymbols(lIndex), lPos
                AddSymbols = True
            End If
        Else
            If AddSymbolID(lSymbolID) Then
                AddSymbols = True
            End If
        End If
    Next lIndex
        
    If AddSymbols Then
        ' Save to file
        ToFile
        
        ' Add back into pool
        AddToPool True
        
        ' Refresh symbol grid dropdown and list
        frmSymbolGrid.RefreshGrid
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolGroup.AddSymbols", eGDRaiseError_Raise

End Function

Friend Sub SetPrivateData(mCopyFrom As mPrivate)
    
    ' get copy of all private data
    m = mCopyFrom
    ' now make gdArrays a "copy" of the original
    m.strID = ""
    Set m.aResults = mCopyFrom.aResults.MakeCopy
    Set m.astrSymbols = mCopyFrom.astrSymbols.MakeCopy
    Set m.aSymbolIds = mCopyFrom.aSymbolIds.MakeCopy
    
    ' Custom Index stuff added 11/8/2001 by DAJ
    Set m.adPriceWeights = mCopyFrom.adPriceWeights.MakeCopy
    Set m.adVolWeights = mCopyFrom.adVolWeights.MakeCopy
    Set m.alFlags = mCopyFrom.alFlags.MakeCopy

End Sub

Public Function MakeCopy() As cSymbolGroup
    
    Dim aCopy As New cSymbolGroup
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

End Function

Public Function AddSymbolID(ByVal nSymbolID&) As Boolean
On Error GoTo ErrSection:
    
    Dim lInsertAt As Long
    
    If m.aSymbolIds.BinarySearch(nSymbolID, lInsertAt) = False Then
        m.aSymbolIds.Add nSymbolID, lInsertAt
        m.adPriceWeights.Add 1#, lInsertAt
        m.adVolWeights.Add 1#, lInsertAt
        m.alFlags.Add 0&, lInsertAt
        AddSymbolID = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolGroup.AddSymbolID", eGDRaiseError_Raise

End Function

Public Sub Save(Optional alToAdd As cGdArray = Nothing, Optional ByVal bKeepExisting As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    If Not alToAdd Is Nothing Then
        If bKeepExisting = True Then
            For lIndex = 0 To alToAdd.Size - 1
                AddSymbolID alToAdd(lIndex)
            Next lIndex
        Else
            m.aSymbolIds.Clear
            m.adPriceWeights.Clear
            m.adVolWeights.Clear
            m.alFlags.Clear
            
            For lIndex = 0 To alToAdd.Size - 1
                AddSymbolID alToAdd(lIndex)
            Next lIndex
        End If
    End If

    ToFile
    If m.eGroupType <> eGROUP_Flag Then AddToPool True
    If FormIsLoaded("frmSymbolGrid") Then frmSymbolGrid.RefreshGrid
    
#If 0 Then
    If m.eGroupType = eGROUP_QuoteList Then
        With frmQuotes
            ' wait until TotalRefresh is not active
            Do While .IsBusy
                Sleep 0.1
            Loop
            .fgQuotes.Redraw = flexRDNone
            If .LoadGrid Then .TotalRefresh False 'True
            .fgQuotes.Redraw = flexRDBuffered
        End With
    End If
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolGroup.Save", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Rename
'' Description: Allow the user to rename the symbol group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Rename()
On Error GoTo ErrSection:

    Dim strNewName As String            ' New name for the symbol group
    Dim strOldName As String            ' Old name of the symbol group
    Dim strText As String               ' Text for the dialog to display to the user
    Dim lSymbolID As Long               ' Symbol ID of the custom index
    Dim lMousePointer As Long           ' Previous mouse pointer
    
    lMousePointer = Screen.MousePointer
    
    strText = "Rename the current Symbol Group as..."
    If m.bIndex = True Then
        strText = strText & vbCrLf & vbCrLf & "(Since you have chosen to make a Custom Index out" _
                                & " of this symbol group, the name must be less than" _
                                & " eight characters in length with no spaces)" & vbCrLf
    End If
    strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
    
    Do While m.bIndex = True
        If InStr(strNewName, " ") > 0 Then
            InfBox "Custom Index symbol cannot contain spaces", "!", , "Error"
        ElseIf InStr(strNewName, "'") > 0 Then
            InfBox "Custom Index symbol cannot contain apostrophes", "!", , "Error"
        ElseIf Len(strNewName) > 8 Then
            InfBox "Custom Index symbol cannot exceed eight characters in length", "!", , "Error"
        Else
            Exit Do
        End If
        
        strText = vbCrLf & vbCrLf & "(Since you have chosen to make a Custom Index out" _
                                & " of this symbol group, the name must be less than" _
                                & " eight characters in length with no spaces)" & vbCrLf
        strText = "Rename the current Symbol Group as..." & strText
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & strNewName & " ; " & strText)
    Loop
    
    If Left(strNewName, 1) = "#" Then
        strNewName = Right(strNewName, Len(strNewName) - 1)
    End If
    
    If (Len(Trim(strNewName)) > 0) And (Trim(UCase(strNewName)) <> UCase(m.strName)) Then
        Screen.MousePointer = vbHourglass
        
        strOldName = UCase(m.strName)
        m.strName = Trim(strNewName)
        
        If m.bIndex = True Then
            ' If the user is renaming the symbol group, delete the old custom index first...
            If Len(strOldName) > 0 Then
                lSymbolID = g.SymbolPool.SymbolIDforSymbol(UCase("#" & strOldName))
                If SU_DeleteComposite(lSymbolID, UCase("#" & strOldName)) Then
                    g.SymbolPool.RemoveCustomIndex lSymbolID
                End If
            End If
        
            lSymbolID = 0&
            If SU_SetCompositeInf(lSymbolID, UCase("#" & m.strName), m.strDesc, m.dPriceDivisor, _
                                    m.aSymbolIds, m.adPriceWeights, m.alFlags, _
                                    m.dVolDivisor, m.adVolWeights) = False Then
                InfBox "Problems saving #" & m.strName, "!", , "Error"
            Else
                m.lSymID = lSymbolID
                If g.SymbolPool.AddCustomIndex(m.lSymID, UCase("#" & m.strName)) = False Then
                    InfBox "Problems saving #" & m.strName, "!", , "Error"
                End If
                UpdateVisibleCharts eRedo9_ReloadData, m.lSymID
            End If
        End If
    
        ToFile
        AddToPool ' to replace the name in the fields table
        frmSymbolGrid.RefreshGrid
    End If

ErrExit:
    Screen.MousePointer = lMousePointer
    Exit Sub
    
ErrSection:
    Screen.MousePointer = lMousePointer
    RaiseError "cSymbolGroup.Rename"
    
End Sub
