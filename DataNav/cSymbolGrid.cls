VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSymbolGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements IVSFlexDataSource

Private Type mPrivate
    fg As VSFlexGrid

    ' array of record numbers for current "index"
    ' (a sorted/filtered set of db record #'s)
    aIndex As cGdArray
    strFilterID As String 'optional filter
    bNewFilter As Boolean

    ' just to track current sortation
    nSortedCol As Integer
    bSortedDescending As Boolean
    eSortFlags As eGdSort_Flags
End Type
Private m As mPrivate

Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.fg = Nothing
    Set m.aIndex = New cGdArray
    m.aIndex.Create eGDARRAY_Longs, 0
    m.strFilterID = ""
    m.nSortedCol = kSymbolCol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolGrid.Initialize", eGDRaiseError_Raise
    
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.fg = Nothing
    Set m.aIndex = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolGrid.Terminate", eGDRaiseError_Raise
    
End Sub

' # of Rows is how big the index (sorted filter) is.
Public Property Get NumRows() As Long
    NumRows = m.aIndex.Size
End Property

Public Property Let NumRows(ByVal nNumRows As Long)
    m.aIndex.Size = nNumRows
    'm.fg.Rows = nNumRows + m.fg.FixedRows
    m.fg.FlexDataSource = Me
End Property

' Tells grid how many columns to have.
Private Function IVSFlexDataSource_GetFieldCount() As Long
    IVSFlexDataSource_GetFieldCount = m.fg.Cols
End Function

' Tells grid name of columns.
Private Function IVSFlexDataSource_GetFieldName(ByVal nCol As Long) As String
On Error GoTo ErrSection:

    ' get field number of the db (stored in ColData of grid)
    Dim nFieldNum&, strName$
    
    nFieldNum = Val(m.fg.ColData(nCol))
    strName = g.SymbolPool.ArrayTable.FieldName(nFieldNum)
    If nCol <> kFlagCol Then
        If UCase(strName) = "DESCRIPTION" Then
            strName = "Description @ Exchange"
        End If
        strName = Space(4) & strName
    ElseIf m.nSortedCol = kFlagCol Then
        strName = " "
    Else
        strName = "Flag"
    End If
    IVSFlexDataSource_GetFieldName = strName

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolGrid.IVSFlexDataSource.GetFieldName", eGDRaiseError_Raise
    
End Function

' Tells grid how many rows to have.
Private Function IVSFlexDataSource_GetRecordCount() As Long
    IVSFlexDataSource_GetRecordCount = NumRows
End Function

' Tells grid what to display in each row and column.
Private Function IVSFlexDataSource_GetData(ByVal nCol As Long, ByVal nRow As Long) As String
On Error GoTo ErrSection:

    Dim vData As Variant
    Dim nFieldNum&, nRecord&
    
If 0 Then
        vData = CStr(nCol) & " = " & CStr(m.fg.ColData(nCol))
Else
    ' get number from the db (at field and record)
    nFieldNum = Val(m.fg.ColData(nCol))
    nRecord = m.aIndex(nRow)
    vData = g.SymbolPool.DataItem(nFieldNum, nRecord, "")
    Select Case VarType(vData)
        Case vbDate
            vData = DateFormat(vData, MM_DD_YY)
        Case vbDouble
            'vData = Round(vData, 4)
        Case vbString
    End Select
End If
    
    IVSFlexDataSource_GetData = vData

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolGrid.IVSFlexDataSource.GetData", eGDRaiseError_Raise
    
End Function

' The user has edited a cell.
Private Sub IVSFlexDataSource_SetData(ByVal nCol As Long, ByVal nRow As Long, ByVal strNewData As String)
On Error GoTo ErrSection:

    Dim nFieldNum&, nRecord&
    
    ' get number from the db (at field and record)
    nFieldNum = Val(m.fg.ColData(nCol))
    nRecord = m.aIndex(nRow)
    
    ' are we going to do anything with it?
    'mDB(nFieldNum, nRecord) = Val(strNewData)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolGrid.IVSFlexDataSource.SetData", eGDRaiseError_Raise
    
End Sub

Public Sub SortOnCol(ByVal nCol&, Optional ByVal nDirection% = 0, _
        Optional ByVal nSymbolIDtoShow& = 0)
On Error GoTo ErrSection:

    Dim i&, n&, nCurRedraw%, nSortField&, nFilterField&
    Dim eSortFlags As eGdSort_Flags
    Static nPrevCol&, bPrevDescending As Boolean

    If nCol = -1 Then
        If 1 Then
            ' we are setting back to current sort
            nCol = m.nSortedCol
            If m.bSortedDescending Then
                nDirection = -1
            Else
                nDirection = 1
            End If
        Else
            ' we are setting back to previous sort
            nCol = nPrevCol
            If bPrevDescending Then
                nDirection = 1
            Else
                nDirection = -1
            End If
        End If
    Else
        ' save previous sort
        nPrevCol = m.nSortedCol
        bPrevDescending = m.bSortedDescending
    End If

    ' exit if invalid column
    If nCol < 0 Or nCol >= m.fg.Cols Then Exit Sub

    ' save some stuff to restore later
    If nSymbolIDtoShow = 0 Then
        nSymbolIDtoShow = g.SymbolPool.SymbolID(m.aIndex(m.fg.Row - 1))
    End If
    nCurRedraw = m.fg.Redraw
    m.fg.Redraw = flexRDNone

    ' determine if descending
    If nCol = m.nSortedCol And nDirection = 0 Then
        ' just toggle direction
        m.bSortedDescending = Not m.bSortedDescending
    ElseIf nDirection = -1 Then
        m.bSortedDescending = True
    Else
        m.bSortedDescending = False
    End If
    m.nSortedCol = nCol

    ' create filtered index from filter field #
    If m.bNewFilter Then
        m.bNewFilter = False
        nFilterField = -1
        If Len(m.strFilterID) > 0 Then
            nFilterField = g.SymbolPool.FieldNumForID(m.strFilterID)
        End If
        Set m.aIndex = g.SymbolPool.ArrayTable.CreateIndex(nFilterField)
    End If
    
    ' sort index
    nSortField = Val(m.fg.ColData(m.nSortedCol))
    m.eSortFlags = eGdSort_Stable 'NO: Or eGdSort_DeleteNullValues
    If m.bSortedDescending Then
        m.eSortFlags = m.eSortFlags Or eGdSort_Descending
    End If
    g.SymbolPool.ArrayTable.SortIndex m.aIndex, nSortField, m.eSortFlags
    ' resize grid (might be different now)
    NumRows = m.aIndex.Size
    'm.FG.Parent.Caption = "Symbols (" & CStr(NumRows) & ")"
    
    With m.fg
        ' show sortation arrow in column heading
        .FillStyle = flexFillSingle
        For i = 0 To .Cols - 1
            .Select 0, i
            If i = m.nSortedCol Then
                If m.bSortedDescending Then
                    .CellPicture = frmMain.img16.ListImages("kSortedDownArrow").Picture
                Else
                    .CellPicture = frmMain.img16.ListImages("kSortedUpArrow").Picture
                End If
                '.CellPictureAlignment = flexPicAlignRightTop
                .CellPictureAlignment = flexPicAlignLeftTop
            Else
                .CellPicture = Nothing
            End If
        Next
    
        ' nicety: set selection back to same symbol
        ' as before (if it exists in the new list)
        If .Visible Then
            If 0 Then ' .Rows > .FixedRows Then
                .Row = .FixedRows ' set to first row
            Else
                .Row = -1 ' no rows in grid
            End If
            If nSymbolIDtoShow >= 0 Then
                For i = 0 To m.aIndex.Size - 1
                    If g.SymbolPool.SymbolID(m.aIndex.Num(i)) = nSymbolIDtoShow Then
                        .Row = i + 1
                        Exit For
                    End If
                Next
            End If
    
            ' restore things
            If .Row >= 0 Then
                .Select .Row, 0, .Row, .Cols - 1
                .ShowCell .Row, 0
            End If
        End If
        .Redraw = nCurRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolGrid.SortOnCol", eGDRaiseError_Raise
    
End Sub

Public Sub InitGrid(vsGrid As VSFlexGrid, ByVal strFields$, _
                    Optional ByVal bAllowMultiSelect As Boolean = True, _
                    Optional ByVal bShowFlags As Boolean = True)
On Error GoTo ErrSection:

    Dim s$, i%
    
    Set m.fg = vsGrid
    With m.fg
        ' Init grid settings
        .Redraw = flexRDNone
        .FixedCols = 0
        .FixedRows = 1
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExSortShowAndMove
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .AllowUserFreezing = flexFreezeColumns
        .SheetBorder = RGB(128, 128, 128)
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .ExtendLastCol = True
        .PicturesOver = True
        .AllowSelection = bAllowMultiSelect
        .Editable = flexEDNone '= flexEDKbdMouse
    
        ' Set up fields to display
        If InStr(UCase(strFields), "INF:SYMBOL|") = 0 Then
            'use default string
            strFields = "GRP:_FLAGS_.GRP|INF:Symbol|INF:Description|INF:FirstDate|INF:LastDate"
        ElseIf Parse(UCase(strFields), "|", 1) <> "GRP:_FLAGS_.GRP" Then
            strFields = "GRP:_FLAGS_.GRP|" & strFields
        End If
        InitFields strFields
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
    
        .ColHidden(kFlagCol) = Not bShowFlags
    
        If .Cols >= 2 Then .FrozenCols = 2
    
        ' tie grid to our class data-wise
        .FlexDataSource = Me
    
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolGrid.InitGrid", eGDRaiseError_Raise
    
End Sub

Private Sub InitFields(ByVal strFields$)
On Error GoTo ErrSection:

    Dim i&, nFld&, strID$, iCount&
    Dim aFields As New cGdArray
    
    aFields.SplitFields strFields, "|"
    With m.fg
        For i = 0 To aFields.Size - 1
            strID = Parse(aFields(i), "\", 1)
            nFld = g.SymbolPool.FieldNumForID(strID)
            
            ' If there is no Flags group yet, create one
            If nFld = -1 And strID = "GRP:_FLAGS_.GRP" Then
                Dim SymbolGroup As New cSymbolGroup
                
                SymbolGroup.MakeSpecialType "_FLAGS_.GRP", "FLAGGED Symbols", eGROUP_Flag
                SymbolGroup.ToFile
                SymbolGroup.AddToPool True
                nFld = g.SymbolPool.FieldNumForID(strID)
                Set SymbolGroup = Nothing
            End If
            
            If nFld >= 0 Then
                .Cols = iCount + 1
                .ColData(iCount) = CStr(nFld) & vbTab & aFields(i)
                .ColAlignment(iCount) = flexAlignCenterCenter
                Select Case nFld
                    Case 2: 'Symbol
                        .ColAlignment(iCount) = flexAlignLeftCenter
                    Case 4: 'Desc
                        .ColWidth(iCount) = 2.4 * .ColWidth(kSymbolCol)
                        .ColAlignment(iCount) = flexAlignLeftCenter
                    Case 7: 'SecType
                        .ColWidth(iCount) = 60 * 10
                    Case Else:
                        If strID = "GRP:_FLAGS_.GRP" Then
                            .ColDataType(iCount) = flexDTBoolean
                            .ColWidth(iCount) = 40 * 10
                        End If
                End Select
                iCount = iCount + 1
            End If
        Next
    End With
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolGrid.InitFields", eGDRaiseError_Raise
    
End Sub

Public Property Get FilterID() As String
    FilterID = m.strFilterID
End Property

Public Property Let FilterID(ByVal strFilterID As String)
    m.strFilterID = strFilterID
    m.bNewFilter = True
End Property

Public Sub ShowRec(ByVal nRec&)
On Error Resume Next

    Dim i&, nRow&

    With m.fg
        nRow = -1
    
        ' find it
        If nRec >= 0 Then
            For i = 0 To m.aIndex.Size - 1
                If m.aIndex.Num(i) = nRec Then
                    nRow = i + 1
                    Exit For
                End If
            Next
        End If
    
        ' show it
        If nRow >= 0 Then
            .Row = nRow
            .Select .Row, 0, .Row, .Cols - 1
            .ShowCell .Row, 0
        End If
    End With

End Sub

Public Function Search(ByVal strFindStr As String) As Long
On Error GoTo ErrSection:

    Dim iLo As Long
    Dim iHi As Long
    Dim iMid As Long
    Dim bFound As Boolean
    Dim iFlag As Integer                ' -1: <, 0: =, 1: >
    Dim iDirection As Integer           ' -1: Descending, 1: Ascending
    
    iLo = m.fg.FixedRows    ' usually 0
    iHi = m.fg.Rows - 1     ' largest elem# in array
    bFound = False
    
    If m.bSortedDescending = True Then iDirection = -1 Else iDirection = 1
    
    Do While iLo <= iHi And Not bFound
        iMid = iLo + (iHi - iLo) \ 2  ' so sum does not overflow
        If UCase(strFindStr) < UCase(m.fg.Cell(flexcpText, iMid, m.nSortedCol)) Then
            iFlag = -1
        ElseIf UCase(strFindStr) > UCase(m.fg.Cell(flexcpText, iMid, m.nSortedCol)) Then
            iFlag = 1
        Else
            iFlag = 0
        End If
        
        Select Case iFlag * iDirection
            Case -1:
                iHi = iMid - 1  ' in lower half
            Case 1:
                iLo = iMid + 1  ' in upper half
                iMid = iLo  ' position to insert
            Case 0:
                bFound = True
                ' now back up to the very first match
                Do While iMid > m.fg.FixedRows
                    If UCase(m.fg.Cell(flexcpText, iMid - 1, m.nSortedCol)) <> UCase(strFindStr) Then Exit Do
                    iMid = iMid - 1
                Loop
        End Select
    Loop
    
    If iMid < m.fg.Rows Then
        Search = iMid
    Else
        Search = m.fg.Rows - 1
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolGrid.Search", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToggleFlags
'' Description: Toggle the flag column on and off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ToggleFlags(Optional ByVal iUseRow& = -1)
On Error GoTo ErrSection:

    Dim lField As Long                  ' Field number in the symbol pool
    Dim lRecord As Long                 ' Record number in the symbol pool
    Dim lIndex As Long                  ' Index into a for loop
    Dim iValue As Integer               ' Value to change them all to
    
    ' Get the location in the symbol pool to toggle
    lField = CLng(Val(m.fg.ColData(kFlagCol)))
    If iUseRow < 0 Then iUseRow = m.fg.RowSel
    If iUseRow >= 0 Then
        lRecord = m.aIndex(iUseRow - 1)
        If g.SymbolPool.ArrayTable(lField, lRecord) = 1 Then
            iValue = 0
        Else
            iValue = 1
        End If
        
        With m.fg
            .Redraw = flexRDNone
            If .IsSelected(iUseRow) Then
                For lIndex = 0 To .SelectedRows - 1
                    lRecord = m.aIndex(.SelectedRow(lIndex) - 1)
                    g.SymbolPool.ArrayTable(lField, lRecord) = iValue
                Next lIndex
            Else
                g.SymbolPool.ArrayTable(lField, lRecord) = iValue
            End If
            .Redraw = flexRDBuffered
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolGrid.ToggleFlags", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearAllFlags
'' Description: Clear all of the flags in the flag column
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearAllFlags()
On Error GoTo ErrSection:

    Dim lField As Long                  ' Field number in the symbol pool
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRecord As Long                 ' Record number into the symbol pool
    
    ' Get the field number for the symbol pool
    lField = CLng(Val(m.fg.ColData(kFlagCol)))
    
    ' clear all the flags for symbols in current list
    For lIndex = 0 To m.aIndex.Size - 1
        lRecord = m.aIndex(lIndex)
        g.SymbolPool.ArrayTable(lField, lRecord) = 0
    Next lIndex
    m.fg.Refresh

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolGrid.ClearAllFlags", eGDRaiseError_Raise
    
End Sub
