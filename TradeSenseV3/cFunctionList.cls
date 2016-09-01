VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private Type mPrivate
    vsFunctionList As VSFlexGrid
    rtDescription As RichTextBox
    Functions As cFunctions
    vsStr As VSFlexString
    vListID As Variant
    astrList As cGdArray
    Lists As cLists
    iHighlightedRow As Integer
    lRetVal As Long
    StrH As cStrh
    vPrevListID As Variant
    bUsage As Boolean
    bPort As Boolean                    'True=Use by Portfolio Navigator
    bSys As Boolean                     'True=Use by System Navigator
    bChart As Boolean                   'True=Use by Charting
    bScan As Boolean                    'True=Use by Criteria Scanning
    bShowNewFunction As Boolean         ' Show/Hide the new function button
End Type
Private m As mPrivate

Property Get ListID() As Variant
    ListID = m.vListID
End Property

'Returns the currently selected row from the functions list box
Property Get Selected() As String
    
    'Text selection list
    With m.vsFunctionList
        'Don't insert double quote if the list is OF.
        If Val(m.vListID) > 0 And Val(m.vListID) <> 19 Then
            Selected = .TextMatrix(.Row, 1) & """"
        Else
            'Function List selection
            Selected = .TextMatrix(.Row, 1)
        End If
    End With

End Property

Property Let FunctionsRef(pData As cFunctions)
    Set m.Functions = pData
End Property
Property Let ListID(pData As Variant)
    m.vListID = pData
End Property
Property Let Lists(pData As cLists)
    Set m.Lists = pData
End Property
Property Let Usage(pData As Byte)
    m.bUsage = pData
    m.bPort = GetBit(pData, 1)
    m.bSys = GetBit(pData, 2)
    m.bChart = GetBit(pData, 3)
    m.bScan = GetBit(pData, 4)
End Property

Public Property Get ShowNewFunction() As Boolean
    ShowNewFunction = m.bShowNewFunction
End Property
Public Property Let ShowNewFunction(ByVal bShowNewFunction As Boolean)
    m.bShowNewFunction = bShowNewFunction
End Property

'Down arrow key pressed
Public Sub SearchDown(pGridHasFocus As Boolean)
On Error GoTo ErrSection:
    
    If Not pGridHasFocus Then
        'Show next function in list...
        With m.vsFunctionList
            .Redraw = flexRDNone
            'If not at end of functions list
            If .Row + 1 < .Rows Then
                'select the next row
                .Row = .Row + 1
            End If
            .Redraw = flexRDDirect
            'make sure our current row is visible
            .ShowCell .Row, 1
      End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cFunctionList.SearchDown", eGDRaiseError_Raise, g.strAppPath

End Sub

'Page down pressed
Public Sub PageDown(pGridHasFocus As Boolean)
On Error GoTo ErrSection:

    Dim NewRow  As Long, nRow As Long
    Dim nVisibleRows As Integer
    
    If Not pGridHasFocus Then
        nVisibleRows = m.vsFunctionList.BottomRow - m.vsFunctionList.TopRow
        'Show next function in list...
        With m.vsFunctionList
            .Redraw = flexRDNone
            
            'If not at end of functions list
            NewRow = .Row + nVisibleRows
            If .Row + nVisibleRows >= .Rows Then
                NewRow = .Rows - 1
            End If
        
            'select the next row
            .Row = NewRow
            
            'make selected row the top row
            .TopRow = NewRow
            .Redraw = flexRDDirect
            'make sure our current row is visible
            .ShowCell .Row, 1
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cFunctionList.PageDown", eGDRaiseError_Raise, g.strAppPath

End Sub

'This search is used when the down arrow is pressed in the editing Rich text box
'and not in the grid inself
Public Sub SearchUp(pGridHasFocus As Boolean)
On Error GoTo ErrSection:
    
    If Not pGridHasFocus Then
        With m.vsFunctionList
            .Redraw = flexRDNone
            If .Row - 1 >= 0 Then
            
                'select the next row
                .Row = .Row - 1
            End If
            .Redraw = flexRDDirect
            'make sure our current row is visible
            .ShowCell .Row, 1
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cFunctionList.SearchUp", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub PageUp(pGridHasFocus As Boolean)
On Error GoTo ErrSection:
    
    Dim NewRow      As Long
    Dim nVisibleRows As Integer
    
    If Not pGridHasFocus Then
        nVisibleRows = m.vsFunctionList.BottomRow - m.vsFunctionList.TopRow
        With m.vsFunctionList
            .Redraw = flexRDNone
            NewRow = .Row - nVisibleRows
            If .Row - nVisibleRows < 0 Then
                NewRow = 0
            End If
        
            'select the next row
            .Row = NewRow
            'make selected row the top row
            .TopRow = NewRow
            .Redraw = flexRDDirect
            'make sure our current row is visible
            .ShowCell .Row, 1
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cFunctionList.PageUp", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub Click()
On Error GoTo ErrSection:
    
    'click is handled in the grid itself and it seems to do the right thing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cFunctionList.Click", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Function ReLoadFunctions() As Boolean
On Error GoTo ErrSection:
    
    Dim astrList As cGdArray
    Dim ListName As String
    Dim X As Integer
    Dim bFunctionView As Boolean
    
    'quick check
    If IsEmpty(m.vListID) Then Exit Function
    Set astrList = New cGdArray
    astrList.Create eGDARRAY_Strings
    
    ''Set astrList = m.Lists.Item(CStr(m.vListID)).ListItems
    astrList.CopyFromHandle m.Lists.Item(CStr(m.vListID)).ListItemsHandle
    ListName = m.Lists.Item(CStr(m.vListID)).ListName
    
    If astrList.Size = 0 Then Exit Function
    
    With m.vsFunctionList
        .Redraw = flexRDNone
        .Rows = 0
        .Clear
        .Rows = astrList.Size
        For X = 0 To .Rows - 1
            If ListName = "Functions" Then
                bFunctionView = True
            Else
                bFunctionView = False
            End If
            .TextMatrix(X, 1) = astrList(X)
        Next X
        If Not ListName = "Of" Then
            .Sort = flexSortGenericAscending
        End If
        .Row = 0
        .Col = 1
        .Redraw = flexRDDirect
    End With
    m.vPrevListID = m.vListID
    m.iHighlightedRow = 0
    
    'need to set up the widgets on the function list form to display
    'list of functions or something else
    frmFunctionList.FunctionView bFunctionView
    ReLoadFunctions = True

ErrExit:
    Exit Function

ErrSection:

End Function

Public Sub Show(pX As Long, pY As Long)
On Error GoTo ErrSection:
    
    Dim astrList As cGdArray
    Dim ListName As String
    Dim X As Integer
    
    If Not ReLoadFunctions Then Exit Sub
    
    'Show function list form X,Y coordinates
    With frmFunctionList
        .ShowNewFunction = m.bShowNewFunction
        
        'force below parameter list if visible
        If frmInputsLine.Visible Then
            .Move pX, frmInputsLine.Top + frmInputsLine.Height
        Else
            .Move pX, pY
        End If
        FormOnTop .hWnd, True, 0, 0
        .Caption = "List: " & m.Lists.Item(CStr(m.vListID)).ListName
        .Visible = True
    End With
    
    With m.vsFunctionList
        .Redraw = flexRDNone
        .Row = 0
        .RowSel = 0
        .Redraw = flexRDDirect
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cFunctionList.Show", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub Search(pStr As String, pCaretPos As Integer)
On Error GoTo ErrSection:

    Dim ScanWord            As String
    Dim Y                   As Integer
    Const invalidStartChars = "1234567890()*^%$#@!~-+[]{}<>?=,./\| "

    ScanWord = m.StrH.curFunction(pStr, pCaretPos)
    
    'tjr 2/03 - added to fix bug of having a '(' in ScanWord
    'we know functions dont start with invalid chars, so lets remove any
    While Len(ScanWord) > 0 And InStr(1, invalidStartChars, Left(ScanWord, 1), vbTextCompare) > 0
        ScanWord = Right(ScanWord, Len(ScanWord) - 1)
    Wend
    
    'Search function array
    If Len(ScanWord) > 0 Then
        With m.vsFunctionList
            .Redraw = False
            For Y = 0 To .Rows - 1
                'm.vsStr.Pattern = "^" & ScanWord
                m.vsStr.Text = .TextMatrix(Y, 1)
                m.vsStr.Pattern = "^" & ScanWord
                If m.vsStr.MatchCount > 0 Then Exit For
            Next Y
            If m.vsStr.MatchCount > 0 Then
                .Select Y, 1
            End If
            .Redraw = True
            'make sure our current row is visible
            .ShowCell .Row, 1
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cFunctionList.Search", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub ShowDescription()
On Error GoTo ErrSection:
    
    Dim pFunction As String
    Dim pFunctionID As Long
    Dim pFound As Boolean
    Dim pFunctionDescription As String
        
    If m.Lists.Item(CStr(m.vListID)).ListName <> "Of" Then
        'init description
        pFunctionDescription = ""
        'get selected row in functionlist
        pFunction = m.vsFunctionList.Text
        If Not pFunction = "" Then
            'retrieve the function and its description
            pFound = m.Functions.NameFound(CStr(pFunction), pFunctionID)
            If pFound Then
                pFunctionDescription = m.Functions.Item(CStr(pFunctionID)).Description
            Else
                pFunctionDescription = "Description for [" & pFunction & "] was not found."
            End If
        End If
        'set it
        m.rtDescription.Text = pFunctionDescription
    Else
        m.rtDescription.Text = ""
    End If

ErrExit:
    Exit Sub

ErrSection:

End Sub

Public Sub LoadFilteredFunctionList()
On Error GoTo ErrSection:
    
    Dim X               As Integer
    Dim astrFunctionsList As cGdArray
    Dim List1            As cList
    Dim ListFnd         As Boolean
    Dim bPort           As Boolean
    Dim bSys            As Boolean
    Dim bChart          As Boolean
    Dim bScan           As Boolean
    Dim Filtered        As Boolean
    Dim iCatID          As Integer
    
    'get the category ID
    iCatID = frmFunctionList!cbCategory.ItemData(frmFunctionList!cbCategory.ListIndex)
    
    ' Build Array of Functions from m.Functions collection
    Set astrFunctionsList = New cGdArray
    astrFunctionsList.Create eGDARRAY_Strings
    
    For X = 1 To m.Functions.Count
        With m.Functions.Item(X)
            'Filter functions by Usage
            bPort = GetBit(.Usage, 1)
            bSys = GetBit(.Usage, 2)
            bChart = GetBit(.Usage, 3)
            bScan = GetBit(.Usage, 4)
            Filtered = True
            If (bPort = True And m.bPort = True) Or _
               (bSys = True And m.bSys = True) Or _
               (bChart = True And m.bChart = True) Or _
               (bScan = True And m.bScan = True) Then
                Filtered = False
            End If
            
            If Not Filtered Then
                'if security is properly set AND
                'if the catID matches or its not set yet - load 'em
                If .SecurityLevel <> NOLIST And _
                (iCatID = .FunctionCategoryID Or iCatID = -1) Then
                    astrFunctionsList.Add .FunctionName
                End If
            End If
        End With
    Next X
    
    'Add Functions List to Lists Collection
    If ListFound("Functions") Then
        ListFnd = True
        Set List1 = m.Lists.Item("Functions")
    Else
        Set List1 = New cList
    End If
    With List1
        .ListID = 0
        .ListName = "Functions"
        .ListItemsHandle = astrFunctionsList.ArrayHandle
    End With
    If Not ListFnd Then m.Lists.Add "Functions", List1
    
ErrExit:
    Set List1 = Nothing
    Set astrFunctionsList = Nothing
    Exit Sub

ErrSection:
    Set List1 = Nothing
    Set astrFunctionsList = Nothing
    RaiseError "TSOCX.cFunctionList.LoadFilteredFunctionList", eGDRaiseError_Raise, g.strAppPath

End Sub

'Load Functions into Lists Collection
Public Sub Load()
On Error GoTo ErrSection:
    
    Dim X               As Integer
    Dim Cnt             As Integer
    Dim astrFunctionsList As cGdArray
    Dim List1            As cList
    Dim ListFnd         As Boolean
    Dim ID              As Long
    Dim bPort           As Boolean
    Dim bSys            As Boolean
    Dim bChart          As Boolean
    Dim bScan           As Boolean
    Dim Filtered        As Boolean
    
    'might not be the best place for this, but for category list to work
    ' we need the functions
    InitCategoryList
    LoadFilteredFunctionList
   
    'Add Symbols to Lists Collection marked as "Of" list
    For X = 1 To m.Lists.Count
        If m.Lists.Item(X).ListName = "Symbol Names" Then
            If ListFound("Of") Then
                ListFnd = True
                Set List1 = m.Lists.Item("Of")
            Else
                Set List1 = New cList
            End If
            With List1
                .ListID = 0
                .ListName = "Of"
                ''.ListItems = m.Lists.Item(X).ListItems
                .ListItemsHandle = m.Lists.Item(X).ListItemsHandle
            End With
            If Not ListFnd Then m.Lists.Add "Of", List1
        End If
    Next X
    
ErrExit:
    Set List1 = Nothing
    Set astrFunctionsList = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cFunctionList.Load", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Function ListFound(pName As String) As Boolean
On Error GoTo ErrSection:

    Dim X       As Long
    
    For X = 1 To m.Lists.Count
        With m.Lists.Item(X)
            If .ListName = pName Then
                ListFound = True
                Exit Function
            End If
        End With
    Next X

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cFunctionList.ListFound", eGDRaiseError_Raise, g.strAppPath

End Function

Private Sub InitFunctionsGrid()
On Error GoTo ErrSection:
    
    With m.vsFunctionList
        .Redraw = False
        .Editable = False
        .BorderStyle = flexBorderFlat
        .AutoSearch = flexSearchNone
        .ExplorerBar = flexExNone
        .ScrollBars = flexScrollBarVertical
        .SelectionMode = flexSelectionByRow
        .ScrollTrack = True
        .HighLight = flexHighlightAlways
        .FixedCols = 0
        .FixedRows = 0
        .AllowUserResizing = flexResizeNone
        .ExtendLastCol = True
        .GridLines = flexGridNone
        .WordWrap = False
        .Cols = 2
        .Rows = 0
        
        'Visible columns
        .ColDataType(0) = flexDTString      'English function name
        .ColWidth(0) = 0 '300
        .ColWidth(1) = 3120
        
        .Rows = 1
        .Redraw = True
        .Row = 0
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cFunctionList.InitFunctionsGrid", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub InitCategoryList()
On Error GoTo ErrSection:
    
    Dim pCategory As cFunctionCategories
    Dim i As Integer
    Dim ii As Long
    Dim aID As cGdArray
    Dim bPort           As Boolean
    Dim bSys            As Boolean
    Dim bChart          As Boolean
    Dim bScan           As Boolean
    Dim Filtered        As Boolean
    
    Set pCategory = New cFunctionCategories
    
    'lets clear the list
    frmFunctionList!cbCategory.Clear
    
    If pCategory.Count < 1 Then pCategory.Load 'only load once
    
    Set aID = New cGdArray
    
    ' Walk the function list and find all categories that have at least 1 function...
    For i = 1 To m.Functions.Count
        If aID(m.Functions.Item(i).FunctionCategoryID) = "" Then
            With m.Functions.Item(i)
                'Filter functions by Usage
                bPort = GetBit(.Usage, 1)
                bSys = GetBit(.Usage, 2)
                bChart = GetBit(.Usage, 3)
                bScan = GetBit(.Usage, 4)
                Filtered = True
                If (bPort = True And m.bPort = True) Or _
                   (bSys = True And m.bSys = True) Or _
                   (bChart = True And m.bChart = True) Or _
                   (bScan = True And m.bScan = True) Then
                    Filtered = False
                End If
    
                If Not Filtered Then
                    If .SecurityLevel <> NOLIST Then
                        For ii = 1 To pCategory.Count
                            If pCategory.Item(ii).FunctionCategoryID = .FunctionCategoryID Then
                                aID(.FunctionCategoryID) = pCategory.Item(ii).FunctionCategory
                                Exit For
                            End If
                        Next
                    End If
                End If 'filtered
            End With 'm.Functions.Item(i)
        End If
    Next
    
    ' Populate the category combo on the frmFunctionList (first is <All>)
    frmFunctionList!cbCategory.AddItem C_ALL
    frmFunctionList!cbCategory.ItemData(frmFunctionList!cbCategory.NewIndex) = -1
    
    For i = 0 To aID.Size - 1
        If aID(i) <> "" Then
            frmFunctionList!cbCategory.AddItem aID(i)
            frmFunctionList!cbCategory.ItemData(frmFunctionList!cbCategory.NewIndex) = i
        End If
    Next
    
    ' Default to C_ALL
    frmFunctionList!cbCategory = C_ALL
    
ErrExit:
    Set aID = Nothing
    Exit Sub

ErrSection:
    Set aID = Nothing
    RaiseError "TSOCX.cFunctionList.InitCategoryList", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub Class_Initialize()
    
    m.bShowNewFunction = True
    Set m.StrH = New cStrh
    Set m.vsFunctionList = frmFunctionList!vsFunctionList
    Set m.rtDescription = frmFunctionList!rtbDescription
    InitFunctionsGrid
    Set m.vsStr = frmFunctionList!vsStr
    m.vsStr.CaseSensitive = False

End Sub

Private Sub Class_Terminate()
    
    Set m.vsFunctionList = Nothing
    Set m.vsStr = Nothing
    Set m.Lists = Nothing
    Set m.rtDescription = Nothing

End Sub
