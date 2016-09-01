VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cOptimizer.cls
'' Description: Class for handling an optimization run
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 06/16/2011   DAJ         Added code for the Highlight Bar Reporter
'' 06/22/2011   DAJ         Fix for running reports with the Highlight Bar Reporter
'' 06/22/2011   DAJ         Changed "# Days" to "# Bars", sort on "# Bars" column by default
'' 07/25/2013   DAJ         Send ToDate to cSystem.LoadForMultRun, Fix for Use Values
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements IVSFlexDataSource

Private Const kNAME_DELIM = vbTab

Private Enum eSN_OptFieldTypes
    eSN_OptFieldTypes_Double = 1
    eSN_OptFieldTypes_TinyInt = 2
    eSN_OptFieldTypes_String = 3
End Enum

Public Enum eGDOptMode
    eGDOptMode_Optimization = 0
    eGDOptMode_MultipleRun
    eGDOptMode_StrategyBasket
    eGDOptMode_GameMode
    eGDOptMode_HighlightBarReport
End Enum

Private Type mPrivate
    Mode As eGDOptMode
    OptForm As Form
    vsTests As VSFlexGrid
    OptCols As cOptCols
    Stats As cStats
    lTotalIterations As Long
    aSystemInputs As cGdArray
    bFirstSystem As Boolean
    StatusBar As cStatusBar
    dRefreshFrequency As Double
    bOptimizationInProgress As Boolean
    aInputs As cGdArray

    ' "memory database" to hold all the data
    Table As cGdTable

    ' array of record numbers for current "index"
    ' (a sorted/filtered set of db record #'s)
    aIndex As cGdArray
    bRecalcIndex As Boolean

    ' just to track current sortation
    iSortedCol As Integer
    bSortedDescending As Boolean
    bPyramiding As Boolean
    
    System As cSystem
    lToDateField As Long
    lFirstOptCol As Long
    lTestRuleID As Long
    TestRunMode As eGDRunMode
    
    bColorSymbols As Boolean
End Type
Private m As mPrivate

'Positions of Inputs string
Private Const C_RULENAME = 0
Private Const C_INPUTNAME = 1
Private Const C_INPUTVALUE = 2
Private Const C_INPUTUSED = 3
Private Const DefaultRows = 200                 'Default rows per field array

Property Let vsTests(pData As VSFlexGrid)
    Set m.vsTests = pData
End Property
Property Let TotalIterations(pData As Long)
    m.lTotalIterations = pData
End Property
Property Get TotalIterations() As Long
    TotalIterations = m.lTotalIterations
End Property
Property Let SystemInputs(pData As cGdArray)
    Set m.aSystemInputs = pData
End Property
Property Let Pyramiding(ByVal pData As Boolean)
    m.bPyramiding = pData
End Property

Public Property Let Mode(ByVal eMode As eGDOptMode)
    m.Mode = eMode
End Property
Public Property Get Mode() As eGDOptMode
    Mode = m.Mode
End Property

' Field names used for columns in grid
' # of Rows is how big the index (sorted filter) is.
Public Property Get NumRows() As Long
    NumRows = m.aIndex.Size
End Property
Public Property Let NumRows(ByVal nNumRows As Long)
    m.aIndex.Size = nNumRows
    'm.vsTests.Rows = nNumRows + m.vsTests.FixedRows
    m.vsTests.FlexDataSource = Me
End Property

Public Property Get TestRuleID() As Long
    TestRuleID = m.lTestRuleID
End Property
Public Property Let TestRuleID(ByVal lRuleID As Long)
    m.lTestRuleID = lRuleID
End Property

Public Property Get TestRunMode() As eGDRunMode
    TestRunMode = m.TestRunMode
End Property
Public Property Let TestRunMode(ByVal eTestRunMode As eGDRunMode)
    m.TestRunMode = eTestRunMode
End Property

'Tells grid how many columns to have.
Private Function IVSFlexDataSource_GetFieldCount() As Long
    IVSFlexDataSource_GetFieldCount = m.OptCols.Count
End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal nCol As Long) As String
    ' get field number of the db (stored in ColData of grid)
    Dim nFieldNum As Long
    Dim strReturn As String
    
    nFieldNum = m.vsTests.ColData(nCol)
    If (m.Table.FieldName(nFieldNum) = "Test#") And (m.OptForm.OptMode = eGDOptMode_GameMode) Then
        strReturn = "Game"
    Else
        strReturn = m.Table.FieldName(nFieldNum)
    End If
    
    IVSFlexDataSource_GetFieldName = strReturn
    
End Function


' Tells grid how many rows to have.
Private Function IVSFlexDataSource_GetRecordCount() As Long
    IVSFlexDataSource_GetRecordCount = NumRows
End Function

' Tells grid what to display in each row and column.
Private Function IVSFlexDataSource_GetData(ByVal nCol As Long, ByVal nRow As Long) As String

    Dim vData As Variant, dbNum As Double
    Dim nFieldNum&, nRecord&
    Dim aFiles As New cGdArray
    Dim strFile$, i&
    
    If nCol = 0 And m.OptForm.OptMode = eGDOptMode_GameMode Then
        'use string directly in table so everything stays in sync when table gets sorted - 6475, 6476
        nRecord = m.aIndex.Num(nRow)
        If nRecord >= 0 And nRecord < m.Table.NumRecords Then
            strFile = m.Table(m.Table.NumFields - 1, nRecord)
            IVSFlexDataSource_GetData = strFile
        End If
        
        Exit Function
    End If
    
    ' get number from the db (at field and record)
    nFieldNum = m.vsTests.ColData(nCol)
    nRecord = m.aIndex.Num(nRow)
    If m.Table.FieldType(nFieldNum) <> eGDARRAY_Strings Then
        dbNum = m.Table.Num(nFieldNum, nRecord)
        
        'Determine if column is a alternating rule, set to "Y" if used.
        If Left(m.vsTests.TextMatrix(0, nCol), 2) = "R-" Then
            If dbNum = -1 Then
                IVSFlexDataSource_GetData = "Y"
            Else
                IVSFlexDataSource_GetData = "N"
            End If
        ElseIf dbNum = -999999 Then
            IVSFlexDataSource_GetData = ""
        ElseIf nFieldNum = m.lToDateField Then
            If m.Table.Num(m.OptCols.Item("ToEnd").FieldNumber, nRecord) Then
                IVSFlexDataSource_GetData = "End of Data"
            Else
                IVSFlexDataSource_GetData = dbNum
            End If
        Else
            IVSFlexDataSource_GetData = dbNum
        End If
        If nCol = 13 And nRow > 0 Then
            'm.vsTests.Cell(flexcpFontBold, nRow, nCol, nRow, nCol) = (dbNum > 1)
        End If
    Else
        IVSFlexDataSource_GetData = m.Table.Item(nFieldNum, nRecord)
    End If
    
    
    If nCol = m.iSortedCol And nRow > 0 Then
        'm.vsTests.Cell(flexcpFontBold, nRow, nCol, nRow, nCol) = True
        nRow = m.vsTests.Rows - 1
        'm.vsTests.Cell(flexcpFontBold, nRow, nCol, nRow, nCol) = True
    Else
        'm.vsTests.Cell(flexcpFontBold, nRow, nCol, nRow, nCol) = False
    End If
        
End Function

' The user has edited a cell.
Private Sub IVSFlexDataSource_SetData(ByVal nCol As Long, ByVal nRow As Long, ByVal strNewData As String)
    
    Dim nFieldNum&, nRecord&
    
    ' get number from the db (at field and record)
    nFieldNum = m.vsTests.ColData(nCol)
    nRecord = m.aIndex(nRow)
    
    If m.OptForm.OptMode = eGDOptMode_GameMode Then
        m.aInputs(nRow) = vbTab & strNewData & vbTab & 0# & vbTab & "False"
    End If
    
    ' are we going to do anything with it?
    'mDB(nFieldNum, nRecord) = Val(strNewData)

End Sub

'This is called each time the column properties change (hide/unhide/filter)
Public Sub RefreshColumns()
On Error GoTo ErrSection:

    Dim X               As Integer
    Dim cNewOptCols     As cOptCols
    
    'Reload changes to fixed statistic columns
    Set cNewOptCols = New cOptCols
    cNewOptCols.Load
    For X = 1 To cNewOptCols.Count
        With cNewOptCols.Item(X)
            m.OptCols.Item(X).Hide = .Hide
            m.OptCols.Item(X).Operator = .Operator
            m.OptCols.Item(X).FilterValue = .FilterValue
        End With
    Next X
    
    'Set Hide/UnHide settings
    m.vsTests.Redraw = flexRDNone
    For X = 1 To m.vsTests.Cols - 1
        'm.vsTests.ColHidden(X - 1) = m.OptCols.Item(X).Hide
        m.vsTests.ColHidden(X) = m.OptCols.Item(m.vsTests.TextMatrix(0, X)).Hide
    Next X
    m.vsTests.Redraw = flexRDBuffered
    
    SetFilter
    If m.bSortedDescending Then
        SortOnCol m.iSortedCol, -1, False
    Else
        SortOnCol m.iSortedCol, 1, False
    End If
    
    Set cNewOptCols = Nothing
    m.bOptimizationInProgress = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptimizer.RefreshColumns", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Public Sub SortOnCol(ByVal nCol&, Optional ByVal nDirection% = 0, Optional ByVal bResetSameRecord = True)
On Error GoTo ErrSection:

    Dim nCurRecord&, i&, n&, nCurRedraw%, nSortField&
    Dim eSortFlags As eGdSort_Flags
    Dim bReformat As Boolean
    Static nPrevCol&, bPrevDescending As Boolean

    'If the inputs column headings are clicked during the optimization then don't sort.
    If m.Mode <> eGDOptMode_HighlightBarReport Then
        If m.bOptimizationInProgress And IsInputColumn(nCol) Then
            Exit Sub
        End If
    End If
    
    If nCol = -1 Then
        If 1 Then
            ' we are setting back to current sort
            nCol = m.iSortedCol
            If m.bSortedDescending Then
                nDirection = -1
            Else
                nDirection = 1
            End If
        Else
            ' we are setting back to previous sort
            nCol = nPrevCol
            If bPrevDescending Then
                nDirection = -1
            Else
                nDirection = 1
            End If
        End If
    Else
        ' save previous sort
        nPrevCol = m.iSortedCol
        bPrevDescending = m.bSortedDescending
    End If

    ' exit if invalid column
    If nCol < 0 Or nCol >= m.vsTests.Cols Then Exit Sub

    ' save some stuff to restore later
    nCurRecord = m.aIndex(m.vsTests.Row - 1)
    nCurRedraw = m.vsTests.Redraw
    m.vsTests.Redraw = flexRDNone

    ' determine if descending
    If nCol = m.iSortedCol And nDirection = 0 Then
        ' just toggle direction
        m.bSortedDescending = Not m.bSortedDescending
    ElseIf nDirection = -1 Then
        m.bSortedDescending = True
    Else
        m.bSortedDescending = False
    End If
    If m.iSortedCol <> nCol Then
        m.iSortedCol = nCol
        bReformat = True
    End If

    ' create filtered index from filter field #
    If m.bRecalcIndex Then
        m.bRecalcIndex = False
        Set m.aIndex = m.Table.CreateIndex(0)
    End If
    
    ' sort index
    nSortField = m.vsTests.ColData(m.iSortedCol)
    eSortFlags = eGdSort_Stable 'Or eGdSort_DeleteNullValues
    If m.bSortedDescending Then
        eSortFlags = eSortFlags Or eGdSort_Descending
    End If
    m.Table.SortIndex m.aIndex, nSortField, eSortFlags

    ' resize grid (might be different now)
    NumRows = m.aIndex.Size
    
    ' Show sortation arrow in column heading
    With m.vsTests
        For i = 0 To .Cols - 1
            If i = m.iSortedCol Then
                If m.bSortedDescending Then
                    .Cell(flexcpPicture, 0, i) = frmMain.img16.ListImages("kSortedDownRight").Picture
                Else
                    .Cell(flexcpPicture, 0, i) = frmMain.img16.ListImages("kSortedUpRight").Picture
                End If
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightTop
               .PicturesOver = True
            Else
                .Cell(flexcpPicture, 0, i) = Nothing
            End If
        Next
    End With
    
    ' nicety: set selection back to same record
    ' as before (if it exists in the new list)
    'With m.vsTests
    '    If NumRows > 0 Then
    '        .Row = .FixedRows
    '        If bResetSameRecord And nCurRecord >= 0 Then
    '            For i = 0 To m.aIndex.Size - 1
    '                If m.aIndex(i) = nCurRecord Then
    '                    .Row = i + 1
    '                    Exit For
    '                End If
    '            Next
    '        End If
    '    End If
    'End With
    
    'Always show the first row after sorting...
    If m.vsTests.Rows > 1 Then m.vsTests.Row = 1
    
    ' restore things
    With m.vsTests
        .Select .Row, 0, .Row, .Cols - 1
        '.AutoSize 0, .Cols - 1     '<<This Bogs things down!!!
        CheckColumnWidths '6/21/01 TLB: so DO THIS INSTEAD
        .ShowCell .Row, 0
        .Redraw = nCurRedraw
    End With

    'Save the column as the optimization column (if not an input column)
    If IsInputColumn(nCol) = False Then
        With m.OptCols
            .OptColumn = .Item(m.vsTests.TextMatrix(0, nCol)).FieldName
            If m.bSortedDescending Then
                .OptOrder = "Desc"
            Else
                .OptOrder = "Asc"
            End If
            .Item(m.vsTests.TextMatrix(0, nCol)).FieldOrder = .OptOrder
            m.OptForm.sbField(5).Caption = "Optimization Column: " & .OptColumn
            m.OptForm.sbField(5).Refresh
        End With
    End If
    
    bReformat = True
    If bReformat Then
        FormatVisibleRows
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptimizer.SortOnCol", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'Testing initialization
Public Sub Init()
On Error GoTo ErrSection:
    
    Dim X               As Long
    Dim strInputName    As String
    
    m.bFirstSystem = True
    m.iSortedCol = 0
    m.lFirstOptCol = 1&
    
    'Initialization
    Set m.Stats = New cStats
    Set m.OptForm = frmOptimizer
    
    'Controls how frequent the test results grid is refreshed.  The lower
    'the number (.1) the more frequent
    'm.dRefreshFrequency = Int(m.lTotalIterations * 0.25)
    m.dRefreshFrequency = 1#
    
    'Progress Bar initialization
    Set m.StatusBar = New cStatusBar
    m.StatusBar.StatusBarControl = m.OptForm.vsStatusBar
    m.OptForm.vsStatusBar.Visible = True
    
    'Load statistic fields into fields collection
    Set m.OptCols = New cOptCols
    m.OptCols.Load (m.Mode = eGDOptMode_HighlightBarReport)
    
    'Build in-memory test results database based on fields collection
    For X = 1 To m.OptCols.Count
        With m.OptCols.Item(X)
            AddNewField .FieldNumber, .FieldName, eSN_OptFieldTypes_Double, .Hide, .FieldFormat
        End With
    Next X
    
    InitGrid
    
    m.Table.NumRecords = DefaultRows
    
    'then set back to 0
    m.Table.NumRecords = 0
    NumRows = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptimizer.Init", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Private Sub AddNewField(lFieldNum As Long, strFieldName As String, _
    FieldType As eSN_OptFieldTypes, bHidden As Boolean, ByVal strFieldFormat As String)
On Error GoTo ErrSection:
    
    Dim iPos As Integer                 ' Position of something in a string
    
    ' Create field array in the table for the new field
    Select Case FieldType
        Case eSN_OptFieldTypes_Double
            m.Table.CreateField eGDARRAY_Doubles, lFieldNum, strFieldName
        Case eSN_OptFieldTypes_TinyInt
            m.Table.CreateField eGDARRAY_TinyInts, lFieldNum, strFieldName
        Case eSN_OptFieldTypes_String
            m.Table.CreateField eGDARRAY_Strings, lFieldNum, strFieldName
    End Select
    
    ' Add field as a new column at the end of the grid
    With m.vsTests
        .Cols = lFieldNum
        .ColData(lFieldNum - 1) = lFieldNum
        .ColHidden(lFieldNum - 1) = bHidden
        ' 6/6/01: FlexGrid's ColFormat does not take string like
        ' VB's "pos;(neg)", so if has ";(" then strip off the first part
        iPos = InStr(strFieldFormat, ";(")
        If iPos > 0 Then
            strFieldFormat = Mid(strFieldFormat, iPos + 1)
        End If
        .ColFormat(lFieldNum - 1) = strFieldFormat
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptimizer.AddNewField", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'Apply array filter here...
Public Sub SetFilter()
On Error GoTo ErrSection:

    Dim i&, nRecs&, v#
    Dim Operator()      As String
    Dim FieldNumber()   As Long
    Dim OperValue()     As Single
    Dim curFilter       As Long
    Dim X               As Long
    Dim TotFilters      As Integer
    Dim aFilter As New cGdArray

    'Load Filters from OptCols collection into array for speed
    ReDim FieldNumber(m.OptCols.Count)
    ReDim Operator(m.OptCols.Count)
    ReDim OperValue(m.OptCols.Count)
    For X = 1 To m.OptCols.Count
        If m.OptCols.Item(X).Operator <> "" Then
            TotFilters = TotFilters + 1
            FieldNumber(TotFilters) = m.OptCols.Item(X).FieldNumber
            Operator(TotFilters) = m.OptCols.Item(X).Operator
            OperValue(TotFilters) = m.OptCols.Item(X).FilterValue
        End If
    Next X
    
    'Create filters array of true/false values
    aFilter.Create eGDARRAY_TinyInts, m.Table.NumRecords, 0
    For i = 0 To m.Table.NumRecords - 1
    
        'Default this test to NOT Filtered
        aFilter(i) = True
        
        'Loop through all filters for all stats and filter...
        For curFilter = 1 To TotFilters
            v = m.Table.Num(FieldNumber(curFilter), i)
            Select Case Operator(curFilter)
                Case ">"
                    If v > OperValue(curFilter) Then
                    Else
                        aFilter(i) = False
                        Exit For
                    End If
                Case ">="
                    If v >= OperValue(curFilter) Then
                    Else
                        aFilter(i) = False
                        Exit For
                    End If
                Case "<"
                    If v < OperValue(curFilter) Then
                    Else
                        aFilter(i) = False
                        Exit For
                    End If
                Case "<="
                    If v <= OperValue(curFilter) Then
                    Else
                        aFilter(i) = False
                        Exit For
                    End If
                Case "<>"
                    If v <> OperValue(curFilter) Then
                    Else
                        aFilter(i) = False
                        Exit For
                    End If
                Case "="
                    If v = OperValue(curFilter) Then
                    Else
                        aFilter(i) = False
                        Exit For
                    End If
            End Select
        Next curFilter
    
    Next i
    
    'now add filter as "0" column in table
    m.Table.AttachField aFilter, 0, "Filter"
    m.bRecalcIndex = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptimizer.SetFilter", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'Called by the EngineCallback routine for each test that occurs...
' Returns:
'   0 - Optimization in progress
'   1 - Abort pressed
'   2 - Error occured
Public Function Add(ByVal pIteration As Long, pTrades As cGdArray, _
    pParmValues As cGdArray) As Byte
On Error GoTo ErrSection:
    
    Dim svFldsCnt           As Integer
    Dim X                   As Long
    Dim Y                   As Long
    Dim InputNbr            As Long
    Dim RuleNbr             As Long
    Dim aFields             As cGdArray
    Dim strInput            As String
    Dim strName             As String
    Dim Dec1                As Integer
    Dim Dec2                As Integer
    Dim InputStartingPos    As Long
    Dim n                   As Long
    Dim iPos                As Integer
    Dim strTemp             As String
    Dim lInsertAt           As Long
    Dim lCol                As Long
    Dim strColumnName       As String
    Dim strColumnDesc       As String
    Static dLastGridRefresh As Double
    Static strMultFutSymbols As String

    'Initialization...
    Add = kSN_OPTIMIZATION_IN_PROGRESS
    m.Table.NumRecords = pIteration
    n = pIteration - 1
    Set aFields = New cGdArray
    
    If m.Mode = eGDOptMode_GameMode Then
        If m.bFirstSystem Then
            Set m.aInputs = New cGdArray
            m.bOptimizationInProgress = True
            
            'add a new field to the table so that the game mode text file name also gets sorted
            m.Table.CreateField eGDARRAY_Strings, m.Table.NumFields, "GameModeFile"
            
            m.bFirstSystem = False
        End If
        strMultFutSymbols = ""
    
    ElseIf m.bFirstSystem Then          'Remove later when parm values passed on init call
        strMultFutSymbols = ""
        
        'Save the Inputs and alternating rules one time.  Force a copy of
        'the array with assignment of element 0.
        Set m.aInputs = New cGdArray
        Set m.aInputs = pParmValues
        m.aInputs(0) = m.aInputs(0)
        
        Set aFields = New cGdArray
        m.bOptimizationInProgress = True
        lInsertAt = 1&

        For X = 0 To pParmValues.Size - 1
            aFields.SplitFields pParmValues(X), vbTab

            Select Case aFields(C_INPUTNAME)
                Case "<alternate>": ' Alternating Rule
                    RuleNbr = RuleNbr + 1
                    m.OptCols.Add False, "R-" & RuleNbr, aFields(C_RULENAME), "", "Desc", "N", "N", True
                    AddNewField m.OptCols.Count, "R-" & RuleNbr, eSN_OptFieldTypes_TinyInt, False, ""
                    m.vsTests.TextMatrix(0, m.OptCols.Count - 1) = "R-" & RuleNbr
                    m.vsTests.ColPosition(m.OptCols.Count - 1) = lInsertAt
                    lInsertAt = lInsertAt + 1&
                
                Case "<system>":    ' System Name
                    m.OptCols.Add False, "Strategy", "", "", "Ascend", "N", "N", False
                    AddNewField m.OptCols.Count, "Strategy", eSN_OptFieldTypes_String, False, ""
                    m.vsTests.TextMatrix(0, m.OptCols.Count - 1) = "Strategy"
                    m.vsTests.ColPosition(m.OptCols.Count - 1) = lInsertAt
                    lInsertAt = lInsertAt + 1&
                
                Case "<symbol>":    ' Symbol
                    m.OptCols.Add False, "Symbol", "", "", "Ascend", "N", "N", False
                    AddNewField m.OptCols.Count, "Symbol", eSN_OptFieldTypes_String, False, ""
                    m.vsTests.TextMatrix(0, m.OptCols.Count - 1) = "Symbol"
                    m.vsTests.ColPosition(m.OptCols.Count - 1) = lInsertAt
                    lInsertAt = lInsertAt + 1&
                
                Case "<period>":    ' Period
                    m.OptCols.Add False, "Period", "", "", "Ascend", "N", "N", False
                    AddNewField m.OptCols.Count, "Period", eSN_OptFieldTypes_String, False, ""
                    m.vsTests.TextMatrix(0, m.OptCols.Count - 1) = "Period"
                    m.vsTests.ColPosition(m.OptCols.Count - 1) = lInsertAt
                    lInsertAt = lInsertAt + 1&
                
                Case "<fromdate>":  ' From date
                    m.OptCols.Add False, "FromDate", "", "", "Ascend", "N", "N", False
                    AddNewField m.OptCols.Count, "FromDate", eSN_OptFieldTypes_Double, False, DateFormat("format")
                    m.vsTests.TextMatrix(0, m.OptCols.Count - 1) = "From Date"
                    m.vsTests.ColPosition(m.OptCols.Count - 1) = lInsertAt
                    lInsertAt = lInsertAt + 1&
                
                Case "<todate>":    ' To Date
                    m.OptCols.Add False, "ToDate", "", "", "Ascend", "N", "N", False
                    AddNewField m.OptCols.Count, "ToDate", eSN_OptFieldTypes_Double, False, DateFormat("format")
                    m.vsTests.TextMatrix(0, m.OptCols.Count - 1) = "To Date"
                    m.lToDateField = m.OptCols.Item("ToDate").FieldNumber
                    m.vsTests.ColPosition(m.OptCols.Count - 1) = lInsertAt
                    lInsertAt = lInsertAt + 1&
                
                Case "<toend>":     ' To End of Data
                    m.OptCols.Add True, "ToEnd", "", "", "Ascend", "N", "N", False
                    AddNewField m.OptCols.Count, "ToEnd", eSN_OptFieldTypes_TinyInt, True, ""
                    m.vsTests.TextMatrix(0, m.OptCols.Count - 1) = "To End"
                    m.vsTests.ColPosition(m.OptCols.Count - 1) = lInsertAt
                    lInsertAt = lInsertAt + 1&
                
                Case "<overrides>"
                    m.OptCols.Add False, "Custom Inputs", "", "", "Ascend", "N", "N", False
                    AddNewField m.OptCols.Count, "Custom Inputs", eSN_OptFieldTypes_String, False, ""
                    m.vsTests.TextMatrix(0, m.OptCols.Count - 1) = "Custom Inputs"
                    m.vsTests.ColPosition(m.OptCols.Count - 1) = lInsertAt
                    lInsertAt = lInsertAt + 1&
                    
                Case "<filenames>"
                    m.OptCols.Add True, "Filenames", "", "", "Ascend", "N", "N", False
                    AddNewField m.OptCols.Count, "Filenames", eSN_OptFieldTypes_String, True, ""
                    m.vsTests.TextMatrix(0, m.OptCols.Count - 1) = "Filenames"
                    m.vsTests.ColPosition(m.OptCols.Count - 1) = lInsertAt
                    lInsertAt = lInsertAt + 1&
                
                Case Else:          ' Optimized Inputs
                    If aFields(C_INPUTUSED) = "true" Then
                        InputNbr = InputNbr + 1
                        
                        If (m.Mode = eGDOptMode_HighlightBarReport) And (InputNbr = 1) Then
                            strColumnName = "# Bars"
                        Else
                            strColumnName = "I-" & Str(InputNbr)
                        End If
                        strColumnDesc = aFields(C_RULENAME) & kNAME_DELIM & aFields(C_INPUTNAME)
                            
                        m.OptCols.Add False, strColumnName, strColumnDesc, "#,##0;(#,##0)", "Desc", "N", "N", aFields(C_INPUTUSED)
                        AddNewField m.OptCols.Count, strColumnName, eSN_OptFieldTypes_Double, False, ""
                        m.vsTests.TextMatrix(0, m.OptCols.Count - 1) = strColumnName
                        m.vsTests.ColPosition(m.OptCols.Count - 1) = lInsertAt
                        lInsertAt = lInsertAt + 1&
                    End If
            End Select
        Next X
        
        m.lFirstOptCol = lInsertAt
        
''        If m.Mode = eGDOptMode_MultipleRun Then
''            m.vsTests.ColPosition(m.vsTests.Cols - 1) = 1
''            m.vsTests.ColPosition(m.vsTests.Cols - 1) = 1
''        End If
        
        If m.Mode = eGDOptMode_HighlightBarReport Then
            m.OptCols.OptColumn = "# Bars"
            m.OptCols.OptOrder = "Ascend"
        End If
        
        'Get Trades header information
        aFields.SplitFields pTrades(0), vbTab
        m.OptForm.sbField(0).Caption = "Total Tests: " & m.lTotalIterations
        
        If (m.Mode = eGDOptMode_Optimization) Or (m.Mode = eGDOptMode_HighlightBarReport) Then
            m.OptForm.sbField(1).Caption = aFields(1)      'System Name
            m.OptForm.sbField(2).Caption = "Data From: " & Format(CVDate(Val(aFields(3))), "mm/dd/yyyy") 'Fromdate
            m.OptForm.sbField(3).Caption = "Data To: " & Format(CVDate(Val(aFields(4))), "mm/dd/yyyy") 'Todate
            
            ' Added following block of code in to make sure that the ampersand
            ' shows up in S&P 500 instead of having the P underlined
            ' 5/14/2001 DAJ
            iPos = InStr(aFields(7), "&")
            If iPos <> 0 Then
                strTemp = Left(aFields(7), iPos) & "&" & Mid(aFields(7), iPos + 1)
            Else
                strTemp = aFields(7)
            End If
            
            m.OptForm.sbField(4).Caption = "Symbol: " & strTemp 'aFields(7)      'Symbol
            m.OptForm.sbField(5).Caption = "Optimization Column: " & m.OptCols.OptColumn
        End If
        
        'Set the previous optimization column
        lCol = ColumnNumber(m.OptCols.OptColumn)
        If lCol <> -1& Then
            m.iSortedCol = lCol
            If m.OptCols.OptOrder = "Desc" Then
                m.bSortedDescending = True
            Else
                m.bSortedDescending = False
            End If
        End If
        
        m.bFirstSystem = False
    End If
    
    'Calculate system performance statistics...
    With m.Stats
        .BeginBalance = 100000
        .MinProfit = 0
        .gdCalculate pTrades
        n = pIteration - 1
        m.Table.Num(m.OptCols.Item("Test#").FieldNumber, n) = pIteration
        m.Table.Num(m.OptCols.Item("Win%").FieldNumber, n) = .WinPct
        m.Table.Num(m.OptCols.Item("Trades").FieldNumber, n) = .TotTrades
        m.Table.Num(m.OptCols.Item("Wins").FieldNumber, n) = .Wins
        m.Table.Num(m.OptCols.Item("Losses").FieldNumber, n) = .Losses
        m.Table.Num(m.OptCols.Item("Payout").FieldNumber, n) = .PLRatio
        m.Table.Num(m.OptCols.Item("Avg Win").FieldNumber, n) = .WinAvg
        m.Table.Num(m.OptCols.Item("Return%").FieldNumber, n) = .ReturnPct
        m.Table.Num(m.OptCols.Item("Avg Loss").FieldNumber, n) = .LossAvg
        m.Table.Num(m.OptCols.Item("Large Win").FieldNumber, n) = .WinLargest
        m.Table.Num(m.OptCols.Item("Large Loss").FieldNumber, n) = .LossLargest
        m.Table.Num(m.OptCols.Item("CW").FieldNumber, n) = .WinConsec
        m.Table.Num(m.OptCols.Item("CL").FieldNumber, n) = .LossConsec
        If .WinRuns = 0 Then
            m.Table.Num(m.OptCols.Item("Avg CW").FieldNumber, n) = 0
        Else
            ''m.Table.Num(m.OptCols.Item("Avg CW").FieldNumber, n) = .WinConsec / .WinRuns
            m.Table.Num(m.OptCols.Item("Avg CW").FieldNumber, n) = .Wins / .WinRuns
        End If
        If .LossRuns = 0 Then
            m.Table.Num(m.OptCols.Item("Avg CL").FieldNumber, n) = 0
        Else
            ''m.Table.Num(m.OptCols.Item("Avg CL").FieldNumber, n) = .LossConsec / .LossRuns
            m.Table.Num(m.OptCols.Item("Avg CL").FieldNumber, n) = .Losses / .LossRuns
        End If
        m.Table.Num(m.OptCols.Item("Gross Prof").FieldNumber, n) = .WinTotal
        m.Table.Num(m.OptCols.Item("Gross Loss").FieldNumber, n) = .LossTotal
        m.Table.Num(m.OptCols.Item("Net Profit").FieldNumber, n) = .TotalNetProfit
        m.Table.Num(m.OptCols.Item("Profit Factor").FieldNumber, n) = .ProfitFactor
        m.Table.Num(m.OptCols.Item("Avg Trade").FieldNumber, n) = .AvgTrade
        m.Table.Num(m.OptCols.Item("Avg Bars/Trade").FieldNumber, n) = .AvgBarsInTrade
        m.Table.Num(m.OptCols.Item("Avg Trades/Yr").FieldNumber, n) = .AvgTradesPerYear
        m.Table.Num(m.OptCols.Item("Max Drawdown").FieldNumber, n) = .MaxDrawdownClosed
        m.Table.Num(m.OptCols.Item("Max Intra").FieldNumber, n) = .MaxDrawdownIntra
        
        ' set flag so next SortOnCol will recalc the index
        m.bRecalcIndex = True
    End With
    
    If m.Mode = eGDOptMode_GameMode Then
        'aardvark 6475, 6476 fix - add file name directly to table so stays in order when table gets sorted
        '   previous code sorted the table, but the file name in m.aInputs were not ever changed to match table data
        If m.Table.FieldName(m.Table.NumFields - 1) = "GameModeFile" Then
            If n >= 0 And n < pParmValues.Size Then
                m.aInputs(n) = pParmValues(n)   'not sure if this is needed, but old code did it so just do it
                
                strName = Parse(pParmValues(n), ".", 1)
                strName = Replace(strName, vbTab, "")
                
                m.Table(m.Table.NumFields - 1, m.Table.NumRecords - 1) = strName
            End If
        Else
            DebugLog "Game mode table does not have GameModeFile field."
            Add = kSN_OPTIMIZATION_ERROR
            GoTo ErrExit
        End If
    Else
        'Save the input values for the current test...
        'These always start one column after the statistic fields
        InputNbr = 0: RuleNbr = 0
        For X = 0 To pParmValues.Size - 1
            aFields.SplitFields pParmValues(X), vbTab
            
            Select Case aFields(C_INPUTNAME)
                Case "<alternate>": ' Alternating Rule
                    RuleNbr = RuleNbr + 1
                    If aFields(C_INPUTVALUE) = "true" Then
                        m.Table.Num(m.OptCols.Item("R-" & RuleNbr).FieldNumber, n) = -1
                    Else
                        m.Table.Num(m.OptCols.Item("R-" & RuleNbr).FieldNumber, n) = 0
                    End If
                    
                Case "<system>":    ' System
                    m.Table.Item(m.OptCols.Item("Strategy").FieldNumber, n) = aFields(C_INPUTVALUE)
                
                Case "<symbol>":    ' Symbol
                    strTemp = aFields(C_INPUTVALUE)
                    m.Table.Item(m.OptCols.Item("Symbol").FieldNumber, n) = strTemp
                    If strMultFutSymbols <> "..." Then
                        If InStr(strTemp, "-") > 0 Then
                            If strMultFutSymbols = "" Then
                                strMultFutSymbols = strTemp ' set to first futures symbol
                            ElseIf strMultFutSymbols <> strTemp Then
                                strMultFutSymbols = "..." ' multiple symbols
                                m.OptForm.tbToolbar.Tools("ID_Colored").Visible = True
                                m.OptForm.tbToolbar.Tools("ID_Colored").Enabled = True
                            End If
                        End If
                    End If
                    
                Case "<period>":    ' Period
                    m.Table.Item(m.OptCols.Item("Period").FieldNumber, n) = aFields(C_INPUTVALUE)
                    
                Case "<fromdate>":  ' From Date
                    m.Table.Item(m.OptCols.Item("FromDate").FieldNumber, n) = aFields(C_INPUTVALUE)
                    
                Case "<todate>":    ' To Date
                    m.Table.Item(m.OptCols.Item("ToDate").FieldNumber, n) = aFields(C_INPUTVALUE)
                    
                Case "<toend>":     ' To End of Data
                    m.Table.Item(m.OptCols.Item("ToEnd").FieldNumber, n) = aFields(C_INPUTVALUE)
                    
                Case "<overrides>"  ' Parameter overrides
                    m.Table.Item(m.OptCols.Item("Custom Inputs").FieldNumber, n) = aFields(C_INPUTVALUE)
                    
                Case "<filenames>"  ' Filename for the report file
                    m.Table.Item(m.OptCols.Item("Filenames").FieldNumber, n) = aFields(C_INPUTVALUE)
                    
                Case Else:          ' Optimized Input
                    If aFields(C_INPUTUSED) = "true" Then
                        InputNbr = InputNbr + 1
                        
                        If (m.Mode = eGDOptMode_HighlightBarReport) And (InputNbr = 1) Then
                            strColumnName = "# Bars"
                        Else
                            strColumnName = "I-" & Str(InputNbr)
                        End If
                        
                        If aFields(C_INPUTVALUE) = "" Then
                            m.Table.Num(m.OptCols.Item(strColumnName).FieldNumber, n) = -999999
                        Else
                            m.Table.Num(m.OptCols.Item(strColumnName).FieldNumber, n) = Val(aFields(C_INPUTVALUE))
                        End If
                    End If
            End Select
        Next X
    End If
    
    'Immediately show first system tests, then just once per second
    If pIteration <= 1 Or (gdTickCount - dLastGridRefresh) / 1000 > m.dRefreshFrequency Then
        
        If m.bSortedDescending Then
            SortOnCol m.iSortedCol, -1, False
        Else
            SortOnCol m.iSortedCol, 1, False
        End If
        
        'Update progress meter
        m.StatusBar.Value = (pIteration / m.lTotalIterations) * 100
        DoEvents
        dLastGridRefresh = gdTickCount
    End If
    
    'Last test end processing...
    If pIteration = m.lTotalIterations Then
    
        'If one or more filters exist then filter final results...
        If m.OptCols.FiltersExist Then
            SetFilter
        End If
        If m.bSortedDescending Then
            SortOnCol m.iSortedCol, -1, False
        Else
            SortOnCol m.iSortedCol, 1, False
        End If
        Add = kSN_OPTIMIZATION_COMPLETED
        m.bOptimizationInProgress = False
    
    End If
    
    lCol = 13
    'm.vsTests.Cell(flexcpFontBold, m.vsTests.TopRow, lCol, m.vsTests.BottomRow, lCol) = True
    
ErrExit:
    Set aFields = Nothing
    Exit Function

ErrSection:
    m.bOptimizationInProgress = False
    Add = kSN_OPTIMIZATION_ERROR
    RaiseError "cOptimizer.Add", eGDRaiseError_Raise
    Resume ErrExit

End Function

Public Sub BeforeScrollTip(pRow As Long)
On Error GoTo ErrSection:
    
    With m.vsTests
        .ScrollTipText = " Row " & pRow & ": " & _
                m.OptCols.Item(1).FieldName & ": " & .Cell(flexcpTextDisplay, pRow, 0)
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptimizer.BeforeScrollTip", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Public Sub ShowReport()
On Error GoTo ErrSection:
    
    Dim X               As Long
    Dim Y               As Long
    Dim aFields         As cGdArray
    Dim aReBuiltInputs  As cGdArray
    Dim RuleNames()     As String
    Dim InputNames()    As String
    Dim InputValues()   As Double
    Dim DecPos1         As Integer
    Dim InputName       As String
    Dim AlternatingStart    As Long
    Dim Start           As Long
    
    Dim strSystemName As String
    Dim strSymbol As String
    Dim lSystemID As Long
    Dim strPeriod As String
    Dim dFromDate As Double
    Dim dToDate As Double
    Dim bToEndOfData As Boolean
    Dim Bars As cGdBars
    Dim strOverrides As String
    Dim dLoadFrom As Double
    
    If (m.Mode = eGDOptMode_Optimization) Or (m.Mode = eGDOptMode_HighlightBarReport) Then
        'Build arrays of inputs for the currently selected test row...
        If m.OptCols.InputStartingPos > 0 Then
            ReDim RuleNames(m.OptCols.Count - m.OptCols.InputStartingPos + 1)
            ReDim InputNames(m.OptCols.Count - m.OptCols.InputStartingPos + 1)
            ReDim InputValues(m.OptCols.Count - m.OptCols.InputStartingPos + 1)
            Start = m.OptCols.InputStartingPos
        Else
            'This section handles no inputs but 1 or more alternating rules
            ReDim RuleNames(m.OptCols.Count - m.OptCols.AltStartingPos + 1)
            ReDim InputNames(m.OptCols.Count - m.OptCols.AltStartingPos + 1)
            ReDim InputValues(m.OptCols.Count - m.OptCols.AltStartingPos + 1)
            Start = m.OptCols.AltStartingPos
        End If
        
        For X = Start To m.OptCols.Count
            Y = Y + 1
            InputName = m.OptCols.Item(X).FieldDesc
            DecPos1 = InStr(InputName, kNAME_DELIM)
            If DecPos1 > 0 Then
                RuleNames(Y) = Left(InputName, DecPos1 - 1)
                InputNames(Y) = Mid(InputName, DecPos1 + Len(kNAME_DELIM))
                InputValues(Y) = m.Table.Num(X, m.vsTests.TextMatrix(m.vsTests.Row, 0) - 1)
            Else
                'Save alternating flag
                If AlternatingStart = 0 Then AlternatingStart = Y
                RuleNames(Y) = InputName
                InputValues(Y) = m.Table.Num(X, m.vsTests.TextMatrix(m.vsTests.Row, 0) - 1)
            End If
        Next X
        
        'aRebuiltInputs will have the updated inputs required by the engine to
        'run a test.  Extract each input row from m.aInputs (has the originally
        'saved inputs from the 1st test).  Update the inputs from the arrays
        'of the saved test.
        Set aReBuiltInputs = New cGdArray
        For X = 0 To m.aInputs.Size - 1
            Set aFields = New cGdArray
            aFields.SplitFields m.aInputs(X), vbTab
    
            'Alternating rules section...
            If aFields(C_INPUTNAME) = "<alternate>" Then
                For Y = AlternatingStart To UBound(RuleNames)
                    If aFields(C_RULENAME) = RuleNames(Y) Then
                        If InputValues(Y) = 0 Then
                            aFields(C_INPUTVALUE) = "false"
                        Else
                            aFields(C_INPUTVALUE) = "true"
                        End If
                        Exit For
                    End If
                Next Y
                
                'Add Alternating Rules row to new cGDArray...
                aReBuiltInputs.Add aFields(C_RULENAME) & vbTab & _
                                   aFields(C_INPUTNAME) & vbTab & _
                                   aFields(C_INPUTVALUE)
            Else
                'Update the input in the saved array
                For Y = 1 To UBound(RuleNames)
                    If aFields(C_RULENAME) = RuleNames(Y) And _
                       aFields(C_INPUTNAME) = InputNames(Y) Then
                        aFields(C_INPUTVALUE) = InputValues(Y)
                        Exit For
                    End If
                Next Y
                
                'Add inputs row to new cGDArray...
                aReBuiltInputs.Add aFields(C_RULENAME) & vbTab & _
                                   aFields(C_INPUTNAME) & vbTab & _
                                   Trim(aFields(C_INPUTVALUE)) & vbTab & _
                                   aFields(C_INPUTUSED)
            End If
        Next X
        
        With g.CurrentSystem
            .RunEngine False, groSingleIteration, False, .TradesFile(False), , , m.lTestRuleID, , , , aReBuiltInputs
        End With
    Else
        With m.vsTests
            For X = 0 To .Cols - 1
                If .TextMatrix(0, X) = "Strategy" Then
                    Y = X
                    Exit For
                End If
            Next X
            
            strSystemName = .TextMatrix(.Row, Y)
            strSymbol = .TextMatrix(.Row, Y + 1)
            strPeriod = m.Table(m.OptCols.Item("Period").FieldNumber, m.aIndex(.Row - 1))
            dFromDate = m.Table(m.OptCols.Item("FromDate").FieldNumber, m.aIndex(.Row - 1))
            bToEndOfData = m.Table(m.OptCols.Item("ToEnd").FieldNumber, m.aIndex(.Row - 1))
            If bToEndOfData Then
                dToDate = 0#
            Else
                dToDate = m.Table(m.OptCols.Item("ToDate").FieldNumber, m.aIndex(.Row - 1))
            End If
            strOverrides = m.Table(m.OptCols.Item("Custom Inputs").FieldNumber, m.aIndex(.Row - 1))
            
            lSystemID = SystemIDForName(strSystemName)
            If lSystemID <> 0 Then
#If 0 Then ' OLD method
                If m.Mode = eGDOptMode_StrategyBasket Then
                    Set m.System = New cSystem
                    m.System.Load lSystemID
                    
                    Set Bars = New cGdBars
            
                    If GetPeriodicity(strPeriod) > ePRD_Days + 1 Then
                        dLoadFrom = 0#
                    Else
                        dLoadFrom = Int(dFromDate - (m.System.BarsLoadedBeforeTrading + 1) * 1.46)
                    End If
                    
                    DM_GetBars Bars, strSymbol, strPeriod, dLoadFrom, dToDate
                    m.System.LoadForMultRun lSystemID, Bars, dFromDate, bToEndOfData, dToDate, strOverrides
                    m.System.RunMode = m.TestRunMode
                    m.System.Test , m.lTestRuleID, , False
                Else
                    Set m.System = g.CurrentSystem.MakeCopy
                    Set Bars = New cGdBars
                    
                    If GetPeriodicity(strPeriod) > ePRD_Days + 1 Then
                        dLoadFrom = 0#
                    Else
                        dLoadFrom = Int(dFromDate - (m.System.BarsLoadedBeforeTrading + 1) * 1.46)
                    End If
                    
                    DM_GetBars Bars, strSymbol, strPeriod, dLoadFrom, dToDate
                    m.System.Bars = Bars
                    m.System.RunMode = m.TestRunMode
                    m.System.Test , m.lTestRuleID, , False
                End If
#Else ' NEW method
                If m.Mode = eGDOptMode_StrategyBasket Then
                    Set m.System = New cSystem
                    m.System.LoadBarsForMultRun lSystemID, strSymbol, strPeriod, dFromDate, bToEndOfData, dToDate, strOverrides, False
                Else
                    Set m.System = g.CurrentSystem.MakeCopy
                    m.System.FromDate = dFromDate
                    m.System.ToDate = dToDate
                    m.System.ToEndOfData = bToEndOfData
                    m.System.LoadMarket1Bars strSymbol, strPeriod, False
                End If
                m.System.RunMode = m.TestRunMode
                m.System.Test , m.lTestRuleID, , False
#End If
            End If
        End With
    End If

ErrExit:
    Set aReBuiltInputs = Nothing
    Set aFields = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cOptimizer.ShowReport", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'Save the inputs of the selected test and pass to System Manager
Public Sub SaveTest()
On Error GoTo ErrSection:
    
    Dim InputRuleNames()    As String
    Dim InputNames()        As String
    Dim InputValues()       As Double
    Dim AltRules()          As String
    Dim AltRulesUse()       As Boolean
    Dim X                   As Integer
    Dim Y                   As Integer
    Dim DecPos1             As Integer
    Dim DecPos2             As Integer
    Dim InputName           As String
    Dim frm As frmSystemManager
    
    'Count number of inputs and size the arrays
    Y = 0
    ReDim InputRuleNames(Y)
    ReDim InputNames(Y)
    ReDim InputValues(Y)
    If m.OptCols.InputStartingPos > 0 Then
        For X = m.OptCols.InputStartingPos To m.OptCols.Count
            If IsInputColumn(X, False) Then
                Y = Y + 1
            End If
        Next X
        ReDim InputRuleNames(Y)
        ReDim InputNames(Y)
        ReDim InputValues(Y)
    
        'Save the inputs
        Y = 0
        For X = m.OptCols.InputStartingPos To m.OptCols.Count
            If IsInputColumn(X, False) Then
                Y = Y + 1
                InputName = m.OptCols.Item(X).FieldDesc
                DecPos1 = InStr(InputName, kNAME_DELIM)
                InputRuleNames(Y) = Left(InputName, DecPos1 - 1)
                InputNames(Y) = Mid(InputName, DecPos1 + Len(kNAME_DELIM))
                InputValues(Y) = m.Table.Num(X, m.vsTests.TextMatrix(m.vsTests.Row, 0) - 1)
            Else
                
            End If
        Next X
    End If
    
    'Count the number of alternating rules and size the arrays
    Y = 0
    ReDim AltRules(Y)
    ReDim AltRulesUse(Y)
    If m.OptCols.AltStartingPos > 0 Then
        For X = m.OptCols.AltStartingPos To m.OptCols.Count
            If Left(m.OptCols.Item(X).FieldName, 2) = "R-" Then
                Y = Y + 1
            End If
        Next X
        ReDim AltRules(Y)
        ReDim AltRulesUse(Y)
        
        'Save the alternating rules
        Y = 0
        For X = m.OptCols.AltStartingPos To m.OptCols.Count
            If Left(m.OptCols.Item(X).FieldName, 2) = "R-" Then
                Y = Y + 1
                AltRules(Y) = m.OptCols.Item(X).FieldDesc
                AltRulesUse(Y) = m.Table.Num(X, m.vsTests.TextMatrix(m.vsTests.Row, 0) - 1)
            End If
        Next X
    End If
    
    Set frm = SysMgrForm
    If frm Is Nothing Then
        Err.Raise vbObjectError + 1000, , "Strategy Editor for this Strategy no longer loaded"
    Else
        frm.OptimizationUpdate _
            InputRuleNames, InputNames, InputValues, AltRules, AltRulesUse
        On Error Resume Next
        If frm.WindowState = vbMinimized Then frm.WindowState = vbNormal
        frm.SetFocus
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cOptimizer.SaveTest", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'Show column description
Public Sub vsTests_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:
    
    Dim txtPreview As RichTextBox
    Dim RetVal As Variant
    Dim strColName As String
    Dim strColDesc As String
    
    RetVal = LockWindowUpdate(m.OptForm.hWnd)
    Set txtPreview = m.OptForm.rtfColDesc
    
    'Get column information
    strColName = m.vsTests.TextMatrix(0, NewCol)
    strColDesc = m.OptCols.Item(strColName).FieldDesc
    
    With txtPreview
        .SelStart = 0
        .SelLength = 7
        .SelItalic = False
        .Text = "Column: " & strColName & "   " & strColDesc
        .SelStart = 8
        .SelLength = Len(strColName)
        .SelBold = True
        .SelStart = .SelStart + Len(strColName) + 1
        .SelLength = Len("   " & strColDesc)
        .SelItalic = True
        .SelBold = False
        .SelLength = 0
    End With
    
ErrExit:
    Set txtPreview = Nothing
    RetVal = LockWindowUpdate(0)
    Exit Sub
    
ErrSection:
    Set txtPreview = Nothing
    RetVal = LockWindowUpdate(0)
    RaiseError "cOptimizer.vsTests_AfterRowColChange", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'Column description
Public Function ColumnDesc(ByVal lCol As Long) As String
On Error GoTo ErrSection:
    
    Dim strReturn As String             ' Return value for the function
    Dim strColName As String            ' Column name
    Dim strColDesc As String            ' Column description
    Dim iPos As Integer                 ' Position of something in a string
    
    strReturn = ""
    If lCol >= 0 Then
        strColName = m.vsTests.TextMatrix(0, lCol)
        strColDesc = m.OptCols.Item(strColName).FieldDesc
        iPos = InStr(strColDesc, vbTab)
        If iPos > 0 Then
            strColDesc = Left(strColDesc, iPos - 1) & " -> " & Mid(strColDesc, iPos + 1)
        End If
        
        strReturn = strColName & " =  " & strColDesc
    End If
    
    ColumnDesc = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptimizer.ColumnDesc"

End Function

Private Sub InitGrid()
On Error GoTo ErrSection:
    
    With m.vsTests
        .Redraw = flexRDNone
        .FixedCols = 0
        .FixedRows = 1
        .ScrollTips = True
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExSortShowAndMove
        .ScrollTrack = True
        .PicturesOver = True
        .SelectionMode = flexSelectionListBox
        .AllowUserFreezing = flexFreezeColumns
        .SheetBorder = RGB(128, 128, 128)
        .FrozenCols = 1             '???
        .FlexDataSource = Me        'tie grid to our class data-wise
        .AutoSize 1, .Cols - 1, , 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptimizer.InitGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Public Sub SaveColumns()
On Error GoTo ErrSection:

    m.OptCols.SaveCols m.vsTests
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptimizer.SaveColumns", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Private Sub Class_Initialize()
    
    m.lTestRuleID = 0&
    Set m.vsTests = Nothing
    Set m.Table = New cGdTable
    Set m.aIndex = New cGdArray
    m.aIndex.Create eGDARRAY_Floats 'Should accomodate sorting on things
                                    'like Winpct, Win avg Consec Losses.
End Sub

Private Sub Class_Terminate()
    
    Set m.vsTests = Nothing
    Set m.aIndex = Nothing
    Set m.Table = Nothing
    Set m.Stats = Nothing
    Set m.StatusBar = Nothing
    Set m.OptForm = Nothing
    Set m.OptCols = Nothing
    Set m.aSystemInputs = Nothing
    Set m.System = Nothing
    m.lToDateField = -1

End Sub

' This routine allows us to avoid doing .AutoSize which really bogs down
' (when tens of thousands of rows, one call to .AutoSize takes minutes!)
Private Sub CheckColumnWidths()
On Error GoTo ErrSection:

    Dim nCol&, nRow&, nWidth&, strText$, nHighest&, nIteration&
    Dim nFieldNum&, nRecord&, nItem&
    Static nLastIterationChecked&

    With m.vsTests
        nFieldNum = m.OptCols.Item("Test#").FieldNumber
        If m.aIndex.Size <= 1 Then
            nLastIterationChecked = 0
            .AutoSize 0, .Cols - 1, , 75
        End If
        
        ' look for all new iterations (since last checked)
        For nItem = 0 To m.aIndex.Size - 1
            ' get iteration# from the table
            nRecord = m.aIndex.Num(nItem)
            nIteration = m.Table.Num(nFieldNum, nRecord)
            If nIteration > nHighest Then nHighest = nIteration
            
            ' if a new iteration, check width of each column
            If nIteration > nLastIterationChecked Then
                nRow = nItem + .FixedRows
                For nCol = 0 To .Cols - 1
                    strText = .Cell(flexcpTextDisplay, nRow, nCol)
                    nWidth = m.OptForm.TextWidth(strText) + 180
                    If nWidth > .ColWidth(nCol) Then
                        .ColWidth(nCol) = nWidth
                    End If
                Next
            End If
        Next
        nLastIterationChecked = nHighest
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptimizer.CheckColumnWidths", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Public Sub LoadSystem()
On Error GoTo ErrSection:
    
    Dim strSymbol As String
    Dim strSystemName As String
    Dim strPeriod As String
    Dim dFromDate As Double
    Dim dToDate As Double
    Dim bToEndOfData As Boolean
    Dim lSystemID As Long
    Dim frm As frmSystemManager
    Dim X&, Y&
    Dim Bars As cGdBars
    Dim dLoadFrom As Double
    Dim strOverrides As String

    With m.vsTests
        For X = 0 To .Cols - 1
            If .TextMatrix(0, X) = "Strategy" Then
                Y = X
                Exit For
            End If
        Next X
        
        strSystemName = .TextMatrix(.Row, Y)
        strSymbol = .TextMatrix(.Row, Y + 1)
        
        strPeriod = m.Table(m.OptCols.Item("Period").FieldNumber, .Row - 1)
        dFromDate = m.Table(m.OptCols.Item("FromDate").FieldNumber, .Row - 1)
        dToDate = m.Table(m.OptCols.Item("ToDate").FieldNumber, .Row - 1)
        bToEndOfData = m.Table(m.OptCols.Item("ToEnd").FieldNumber, .Row - 1)
        strOverrides = m.Table(m.OptCols.Item("Custom Inputs").FieldNumber, .Row - 1)
    End With
        
    lSystemID = SystemIDForName(strSystemName)
    If lSystemID <> 0 Then
        Set m.System = New cSystem
#If 1 Then
        m.System.LoadBarsForMultRun lSystemID, strSymbol, strPeriod, dFromDate, bToEndOfData, dToDate, strOverrides, False
#Else
        m.System.Load lSystemID
        
        Set Bars = New cGdBars
        
        If GetPeriodicity(strPeriod) > ePRD_Days + 1 Then
            dLoadFrom = 0#
        Else
            dLoadFrom = Int(dFromDate - (m.System.BarsLoadedBeforeTrading + 1) * 1.46)
        End If
        
        If bToEndOfData Then
            DM_GetBars Bars, strSymbol, strPeriod, dLoadFrom
        Else
            DM_GetBars Bars, strSymbol, strPeriod, dLoadFrom, dToDate
        End If
        m.System.LoadForMultRun lSystemID, Bars, dFromDate, bToEndOfData, dToDate, strOverrides
#End If
        Set frm = New frmSystemManager
        frm.ShowSystem m.System
    End If

ErrExit:
    Set frm = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cOptimizer.LoadSystem", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

Private Function SysMgrForm() As frmSystemManager
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim frm As Form
    Dim strID As String
    
    If g.CurrentSystem Is Nothing Then Exit Function
    
    strID = g.CurrentSystem.SystemNumber
    For lIndex = Forms.Count - 1 To 0 Step -1
        Set frm = Forms(lIndex)
        If frm.Name = "frmSystemManager" Then
            If frm.ID = strID Then
                Set SysMgrForm = frm
                Exit For
            End If
        End If
    Next
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptimizer.SysMgrForm", eGDRaiseError_Raise

End Function

Public Function SelectedFiles() As cGdArray
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrFiles As New cGdArray       ' Array of selected files
    
    With m.vsTests
        For lIndex = 0 To .SelectedRows - 1
            astrFiles.Add m.Table(m.OptCols.Item("Filenames").FieldNumber, m.aIndex(.SelectedRow(lIndex) - 1))
        Next lIndex
    End With
    
    Set SelectedFiles = astrFiles

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptimizer.SelectedFiles", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsInputColumn
'' Description: Determine if the given column is an input column
'' Inputs:      Column, From Grid?
'' Returns:     True if Input Column, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsInputColumn(ByVal lColumn As Long, Optional ByVal bFromGrid As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strColumnName As String         ' Column name
    
    bReturn = False
    If bFromGrid Then
        strColumnName = m.vsTests.TextMatrix(0, lColumn)
    Else
        strColumnName = m.OptCols.Item(lColumn).FieldName
    End If
    
    If (m.Mode = eGDOptMode_HighlightBarReport) And (UCase(strColumnName) = "# BARS") Then
        bReturn = True
    ElseIf UCase(Left(strColumnName, 2)) = "I-" Then
        bReturn = True
    End If
    
    IsInputColumn = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptimizer.IsInputColumn"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColumnNumber
'' Description: Deterine the column number for the given name
'' Inputs:      Column Name
'' Returns:     Column Number (-1 if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ColumnNumber(ByVal strColumnName As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lCol As Long                    ' Index into a for loop
    
    lReturn = -1&
    With m.vsTests
        For lCol = 0 To .Cols - 1
            If UCase(.TextMatrix(0, lCol)) = UCase(strColumnName) Then
                lReturn = lCol
                Exit For
            End If
        Next lCol
    End With
    
    ColumnNumber = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptimizer.ColumnNumber"
    
End Function

Public Sub FormatVisibleRows()

    Dim i&, s$, d#, nRow&, nColor&
    Dim bColorSymbols As Boolean
    Static nPFcolumn&

If Not IsIDE Then
    On Error Resume Next ' just don't crash on errors when running in production
End If
    
    With m.vsTests
        
        If IsIDE Then
            i = m.OptCols.Item("Profit Factor").FieldNumber
            i = .ColData(i)
            i = 10
            i = m.iSortedCol
            s = .Cell(flexcpText, 0, 13)
            i = 13
            '.Cell(flexcpFontBold, .TopRow, i, .BottomRow, i) = True
        End If
        
        If .Rows <= .FixedRows + 1 Then
            For i = 0 To .Cols - 1
                s = .Cell(flexcpText, 0, i)
                If UCase(Trim(s)) = "PROFIT FACTOR" Then
                    nPFcolumn = i
                    Exit For
                End If
            Next
        End If
        
        If m.bColorSymbols Then
'            .BackColorAlternate = 0
        End If
        
        For nRow = .TopRow To .BottomRow
            If nRow >= .FixedRows And nRow < .Rows Then
                nColor = 0
                If m.bColorSymbols Then
                    s = .Cell(flexcpText, nRow, 2)
                    If InStr(s, "-") > 0 Then
                        nColor = GetFuturesCategoryColor(s)
                    Else
                        nColor = RGB(255, 255, 255) 'white
                    End If
                End If
                .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = nColor
                
                ' bold profit factors > 1
                d = .Cell(flexcpValue, nRow, nPFcolumn)
                .Cell(flexcpFontBold, nRow, nPFcolumn) = (d > 1)
            End If
        Next
        s = ""
        '.ScrollTipText = " Row " & pRow & ": " & _
                m.OptCols.Item(1).FieldName & ": " & .Cell(flexcpTextDisplay, pRow, 0)
    End With

End Sub

Public Property Get ColorSymbols() As Boolean
    ColorSymbols = m.bColorSymbols
End Property

Public Property Let ColorSymbols(ByVal bColorSymbols As Boolean)
    If m.bColorSymbols <> bColorSymbols Then
        m.bColorSymbols = bColorSymbols
        FormatVisibleRows
    End If
End Property
