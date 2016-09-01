Attribute VB_Name = "mCommon"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mCommon.bas
'' Description: Common routines for the library manager
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/03/2013   DAJ         Move Strategy Baskets into database
'' 05/01/2013   DAJ         Shadow Trading
'' 07/23/2013   DAJ         Added the StrategyBasketHasFilter function
'' 04/01/2014   DAJ         Removed SetEditorCaption since it is now in mGenesis
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Text
Option Explicit

Public Const ALT_GRID_ROW_COLOR = &HC8F0FF

Public Type gGlobals
    dbNav As Database
    WrkJet As Workspace
    CommonBridge As Object
    bChanged As Boolean
    strPassword As String
    strIniFile As String
    strAppPath As String
    
    bReload As Boolean                  ' Does the library need to be reloaded?

    CalledFrom As eCalledFrom
    strMdbName As String
    strCommonDLL As String
    strTradeSenseOCX As String
    
    lHonestDate As Long
    lLCD As Long
    bShowShadow As Boolean              ' Show the Shadow Trading controls?

    ListImages As ListImages
    frmOwner As Form
    
    Help As Object
End Type
Global g As gGlobals

Public Enum eGDDependencyFilter
    eGDDependencyFilter_AllLibraries = 0
    eGDDependencyFilter_NonBuiltIn
    eGDDependencyFilter_NonBuiltInNonUser
    eGDDependencyFilter_UserLibraryOnly
End Enum

Global Const kUserLibrary = 8

Public Sub ShowMsg()

    If Err.Number < 0 Then
        MsgBox Err.Description, vbInformation, "Message"
    Else
        MsgBox "An unexpected error occurred.  Please report the following: " & Chr(13) & Chr(10) & _
            "Source:  " & Err.Source & Chr(13) & Chr(10) & _
            "Message: " & Err.Description & Chr(13) & Chr(10), vbCritical, "Error Message"
    End If
    Screen.MousePointer = vbDefault

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImplementationTypeDesc
'' Description: Return a Function Type Description
'' Inputs:      Type ID of the Function
'' Returns:     Description of the Function Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ImplementationTypeDesc(ByVal pID As Byte) As String
On Error GoTo ErrSection:

    Select Case pID
        Case 1: ImplementationTypeDesc = "Compiled"
        Case 2: ImplementationTypeDesc = "TradeSense"
        Case 3: ImplementationTypeDesc = "Internal"
        Case 4: ImplementationTypeDesc = "Compiled Action"
    End Select
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "Common.ImplementationTypeDesc", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SecurityDesc
'' Description: Return a Security Description
'' Inputs:      Security Level
'' Returns:     Description of the Security Level
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SecurityDesc(pID As Variant) As String
On Error GoTo ErrSection:

    Select Case pID
        Case 0: SecurityDesc = "Can Edit/Can View"
        Case 1: SecurityDesc = "No Edit/Can View"
        Case 2: SecurityDesc = "No Edit/No View"
        Case 3: SecurityDesc = "No Access"
    End Select
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "Common.SecurityDesc", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Picture16
'' Description: Return the appropriate icon for the form
'' Inputs:      Name of the Picture
'' Returns:     Picture Object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Picture16(ByVal strPicture As String) As Object
On Error Resume Next
    
    Set Picture16 = g.ListImages(strPicture).Picture
    If Picture16 Is Nothing Then
        Set Picture16 = g.ListImages("kBlank").Picture
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckedCell
'' Description: Returns whether the cell of the grid is checked or not
'' Inputs:      Grid, Row, and Column of cell to check
'' Returns:     True if checked, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get CheckedCell(fg As VSFlexGrid, ByVal lRow As Long, ByVal lCol As Long) As Boolean
On Error Resume Next
    
    Dim lValue As Long                  ' Value of the given cell
    
    lValue = fg.Cell(flexcpChecked, lRow, lCol)
    If lValue = flexChecked Then
        CheckedCell = True
    Else
        CheckedCell = False
    End If
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckedCell
'' Description: Sets a grid cell as being checked or not
'' Inputs:      Grid, Row, and Column of cell to check, Value to set it to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let CheckedCell(fg As VSFlexGrid, ByVal lRow As Long, ByVal lCol As Long, ByVal bIsChecked As Boolean)
On Error Resume Next

    If bIsChecked Then
        fg.Cell(flexcpChecked, lRow, lCol) = flexChecked
    Else
        fg.Cell(flexcpChecked, lRow, lCol) = flexUnchecked
    End If

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowHeight
'' Description: Return what the row height should be given the text of the
''              cell and the font size
'' Inputs:      Form, Font, and String to figure height for
'' Returns:     Height it should be
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RowHeight(frm As Form, NewFont As StdFont, ByVal strText As String) As Long
On Error GoTo ErrSection:

    Dim OldFont As StdFont

    Set OldFont = frm.Font
    Set frm.Font = NewFont
    RowHeight = frm.TextHeight(strText)
    Set frm.Font = OldFont

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "Common.RowHeight", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Function

Public Function ChangeGridFont(fg As VSFlexGrid, Optional ByVal bResizeColumns As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the Redraw
    Dim lAutoSizeMode As Long           ' Current status of the Auto Size Mode

    If CommonDialogFont(frmLibrary.CommonDialog1, fg.Font) Then
        With fg
            lRedraw = .Redraw
            lAutoSizeMode = .AutoSizeMode
            
            .Redraw = flexRDNone
            .Font = .Font '(this is required to trigger the grid to reset itself!)
            
            If bResizeColumns Then
                .AutoSizeMode = flexAutoSizeColWidth
                .AutoSize 0, .Cols - 1, , 75
            End If
            
            .AutoSizeMode = lAutoSizeMode
            .Redraw = lRedraw
        End With
        ChangeGridFont = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.ChangeGridFont", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DepLibs
'' Description: Retrieve a list of libraries that this library depends on
'' Inputs:      Library ID to check
'' Returns:     Array of Dependant Libraries
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DepLibs(ByVal lLibraryID As Long, bUserLibrary As Boolean) As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As New cGdArray      ' Array of libraries to return
    Dim QryDef As QueryDef              ' Query definition from the database
    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position to input the library

    astrReturn.Create eGDARRAY_Strings
    bUserLibrary = False

    ' Loop through "qryDep..." queries to determine dependent libraries
    For lIndex = 0 To g.dbNav.QueryDefs.Count - 1
        If Left(g.dbNav.QueryDefs.Item(lIndex).Name, 6) = "qryDep" Then
            Set QryDef = g.dbNav.QueryDefs.Item(lIndex)
            QryDef.Parameters(0).Value = lLibraryID
            Set rs = QryDef.OpenRecordset(, dbOpenSnapshot)
            
            Do Until rs.EOF
                If astrReturn.BinarySearch(rs!LibraryName, lPos) = False Then
                    Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] WHERE [LibraryName]='" & rs!LibraryName & "';", dbOpenDynaset)
                    If Not (rs2.BOF And rs2.EOF) Then
                        If rs2!BuiltIn = False Then
                            astrReturn.Add rs!LibraryName, lPos
                        End If
                        If rs2!LibraryID = kUserLibrary Then
                            bUserLibrary = True
                        End If
                    End If
                End If
                
                rs.MoveNext
            Loop
        End If
    Next lIndex
    
    Set DepLibs = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mCommon.DepLibs", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FunctionDependencies
'' Description: Figure out the dependant functions for the Function ID passed in
'' Inputs:      Array of Dependencies, Function ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FunctionDependencies(astrDepends As cGdArray, ByVal lID As Long, ByVal nFilter As eGDDependencyFilter)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Select Case nFilter
        Case eGDDependencyFilter_AllLibraries
            Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions_1.FunctionName, tblFunctions_1.FunctionID, tblLibrarys.LibraryID, tblLibrarys.LibraryName " & _
                    "FROM ((tblFunctions INNER JOIN tblFunctionRefs ON tblFunctions.FunctionID = tblFunctionRefs.FunctionID) INNER JOIN tblFunctions AS tblFunctions_1 ON tblFunctionRefs.FunctionIDRef = tblFunctions_1.FunctionID) INNER JOIN tblLibrarys ON tblFunctions_1.LibraryID = tblLibrarys.LibraryID " & _
                    "WHERE (((tblFunctions.FunctionID)=" & Str(lID) & "));", dbOpenDynaset)
        
        Case eGDDependencyFilter_NonBuiltIn
            Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions_1.FunctionName, tblFunctions_1.FunctionID, tblLibrarys.LibraryID, tblLibrarys.LibraryName, tblLibrarys.BuiltIn " & _
                    "FROM ((tblFunctions INNER JOIN tblFunctionRefs ON tblFunctions.FunctionID = tblFunctionRefs.FunctionID) INNER JOIN tblFunctions AS tblFunctions_1 ON tblFunctionRefs.FunctionIDRef = tblFunctions_1.FunctionID) INNER JOIN tblLibrarys ON tblFunctions_1.LibraryID = tblLibrarys.LibraryID " & _
                    "WHERE (((tblFunctions.FunctionID)=" & Str(lID) & ") AND ((tblLibrarys.BuiltIn)=0));", dbOpenDynaset)
                    
        Case eGDDependencyFilter_NonBuiltInNonUser
            Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions_1.FunctionName, tblFunctions_1.FunctionID, tblLibrarys.LibraryID, tblLibrarys.LibraryName, tblLibrarys.BuiltIn " & _
                    "FROM ((tblFunctions INNER JOIN tblFunctionRefs ON tblFunctions.FunctionID = tblFunctionRefs.FunctionID) INNER JOIN tblFunctions AS tblFunctions_1 ON tblFunctionRefs.FunctionIDRef = tblFunctions_1.FunctionID) INNER JOIN tblLibrarys ON tblFunctions_1.LibraryID = tblLibrarys.LibraryID " & _
                    "WHERE (((tblFunctions.FunctionID)=" & Str(lID) & ") AND ((tblLibrarys.BuiltIn)=0) AND ((tblLibrarys.LibraryID)<>" & Str(kUserLibrary) & "));", dbOpenDynaset)
        
        Case eGDDependencyFilter_UserLibraryOnly
            Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions_1.FunctionName, tblFunctions_1.FunctionID, tblLibrarys.LibraryID, tblLibrarys.LibraryName " & _
                    "FROM ((tblFunctions INNER JOIN tblFunctionRefs ON tblFunctions.FunctionID = tblFunctionRefs.FunctionID) INNER JOIN tblFunctions AS tblFunctions_1 ON tblFunctionRefs.FunctionIDRef = tblFunctions_1.FunctionID) INNER JOIN tblLibrarys ON tblFunctions_1.LibraryID = tblLibrarys.LibraryID " & _
                    "WHERE (((tblFunctions.FunctionID)=" & Str(lID) & ") AND ((tblLibrarys.LibraryID)=" & Str(kUserLibrary) & "));", dbOpenDynaset)
    End Select
    
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            astrDepends.Add Str(rs!FunctionID) & vbTab & rs!FunctionName & vbTab & "Function" & vbTab & Str(rs!LibraryID) & vbTab & rs!LibraryName
            FunctionDependencies astrDepends, rs!FunctionID, nFilter
            
            rs.MoveNext
        Loop
    End If
            
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "mCommon.FunctionDependencies", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RuleDependencies
'' Description: Figure out the dependant functions for the Rule ID passed in
'' Inputs:      Array of Dependencies, Rule ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RuleDependencies(astrDepends As cGdArray, ByVal lID As Long, ByVal nFilter As eGDDependencyFilter)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Select Case nFilter
        Case eGDDependencyFilter_AllLibraries
            Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.FunctionName, tblFunctions.FunctionID, tblLibrarys.LibraryID, tblLibrarys.LibraryName " & _
                    "FROM tblRules INNER JOIN ((tblLibrarys INNER JOIN tblFunctions ON tblLibrarys.LibraryID = tblFunctions.LibraryID) INNER JOIN tblFunctionRules ON tblFunctions.FunctionID = tblFunctionRules.FunctionIDRef) ON tblRules.RuleID = tblFunctionRules.RuleID " & _
                    "WHERE (((tblRules.RuleID)=" & Str(lID) & "));", dbOpenDynaset)
        
        Case eGDDependencyFilter_NonBuiltIn
            Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.FunctionName, tblFunctions.FunctionID, tblLibrarys.LibraryID, tblLibrarys.LibraryName, tblLibrarys.BuiltIn " & _
                    "FROM tblRules INNER JOIN ((tblLibrarys INNER JOIN tblFunctions ON tblLibrarys.LibraryID = tblFunctions.LibraryID) INNER JOIN tblFunctionRules ON tblFunctions.FunctionID = tblFunctionRules.FunctionIDRef) ON tblRules.RuleID = tblFunctionRules.RuleID " & _
                    "WHERE (((tblRules.RuleID)=" & Str(lID) & ") AND ((tblLibrarys.BuiltIn)=0));", dbOpenDynaset)
        
        Case eGDDependencyFilter_NonBuiltInNonUser
            Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.FunctionName, tblFunctions.FunctionID, tblLibrarys.LibraryID, tblLibrarys.LibraryName, tblLibrarys.BuiltIn " & _
                    "FROM tblRules INNER JOIN ((tblLibrarys INNER JOIN tblFunctions ON tblLibrarys.LibraryID = tblFunctions.LibraryID) INNER JOIN tblFunctionRules ON tblFunctions.FunctionID = tblFunctionRules.FunctionIDRef) ON tblRules.RuleID = tblFunctionRules.RuleID " & _
                    "WHERE (((tblRules.RuleID)=" & Str(lID) & ") AND ((tblLibrarys.BuiltIn)=0) AND ((tblLibrarys.LibraryID)<>" & Str(kUserLibrary) & "));", dbOpenDynaset)
        
        Case eGDDependencyFilter_UserLibraryOnly
            Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.FunctionName, tblFunctions.FunctionID, tblLibrarys.LibraryID, tblLibrarys.LibraryName " & _
                    "FROM tblRules INNER JOIN ((tblLibrarys INNER JOIN tblFunctions ON tblLibrarys.LibraryID = tblFunctions.LibraryID) INNER JOIN tblFunctionRules ON tblFunctions.FunctionID = tblFunctionRules.FunctionIDRef) ON tblRules.RuleID = tblFunctionRules.RuleID " & _
                    "WHERE (((tblRules.RuleID)=" & Str(lID) & ") AND ((tblLibrarys.LibraryID)=" & Str(kUserLibrary) & "));", dbOpenDynaset)
    End Select
    
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            astrDepends.Add Str(rs!FunctionID) & vbTab & rs!FunctionName & vbTab & "Function" & vbTab & Str(rs!LibraryID) & vbTab & rs!LibraryName
            FunctionDependencies astrDepends, rs!FunctionID, nFilter
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "mCommon.RuleDependencies", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SystemDependencies
'' Description: Figure out the dependant functions/rules for the given System ID
'' Inputs:      Array of Dependencies, System ID, Include Rules?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SystemDependencies(astrDepends As cGdArray, ByVal lID As Long, ByVal bIncludeRules As Boolean, ByVal nFilter As eGDDependencyFilter)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbNav.OpenRecordset("SELECT tblRules.Name, tblRules.RuleID, tblLibrarys.LibraryID, tblLibrarys.LibraryName " & _
            "FROM tblLibrarys INNER JOIN (tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID) ON tblLibrarys.LibraryID = tblRules.LibraryID " & _
            "WHERE (((tblSystemRules.SystemNumber)=" & Str(lID) & "));", dbOpenDynaset)
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            If bIncludeRules Then
                If (rs!LibraryID = kUserLibrary) Or (nFilter <> eGDDependencyFilter_UserLibraryOnly) Then
                    astrDepends.Add Str(rs!RuleID) & vbTab & rs!Name & vbTab & "Rule" & vbTab & Str(rs!LibraryID) & vbTab & rs!LibraryName
                End If
            End If
            RuleDependencies astrDepends, rs!RuleID, nFilter
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "mCommon.SystemDependencies", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BasketDependencies
'' Description: Figure out the dependant strategies for the given basket
'' Inputs:      Array of Dependencies, Strategy Basket ID, Include System Rules?, Filter
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BasketDependencies(astrDepends As cGdArray, ByVal lID As Long, ByVal bIncludeRules As Boolean, ByVal nFilter As eGDDependencyFilter)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbNav.OpenRecordset("SELECT tblSystems.SystemName, tblSystems.SystemNumber, tblLibrarys.LibraryID, tblLibrarys.LibraryName " & _
            "FROM tblStrategyBaskets INNER JOIN (tblStrategyBasketItems INNER JOIN (tblSystems INNER JOIN tblLibrarys ON tblSystems.LibraryID=tblLibrarys.LibraryID) ON tblStrategyBasketItems.SystemNumber=tblSystems.SystemNumber) ON tblStrategyBaskets.StrategyBasketID=tblStrategyBasketItems.StrategyBasketID " & _
            "WHERE (((tblStrategyBaskets.StrategyBasketID)=" & Str(lID) & "));", dbOpenDynaset)
    Do While Not rs.EOF
        If (rs!LibraryID = kUserLibrary) Or (nFilter <> eGDDependencyFilter_UserLibraryOnly) Then
            astrDepends.Add Str(rs!SystemNumber) & vbTab & rs!SystemName & vbTab & "System" & vbTab & Str(rs!LibraryID) & vbTab & rs!LibraryName
        End If
        SystemDependencies astrDepends, rs!SystemNumber, bIncludeRules, nFilter
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mCommon.BasketDependencies", , g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LibraryDependencies
'' Description: Figure out the dependant functions/rules for the given Library
'' Inputs:      Array of Dependencies, Library ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LibraryDependencies(astrDepends As cGdArray, ByVal lID As Long, ByVal nFilter As eGDDependencyFilter)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    ' Walk through strategy baskets that are in the library...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets] WHERE [LibraryID]=" & Str(lID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        BasketDependencies astrDepends, rs!StrategyBasketID, False, nFilter
        rs.MoveNext
    Loop
    
    ' Walk through rules in strategies that are in the library...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] WHERE [LibraryID]=" & Str(lID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        SystemDependencies astrDepends, rs!SystemNumber, False, nFilter
        rs.MoveNext
    Loop
    
    ' Walk through any building block rules in the library...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] WHERE [SystemNumber]=0 AND [LibraryID]=" & Str(lID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        RuleDependencies astrDepends, rs!RuleID, nFilter
        rs.MoveNext
    Loop
    
    ' Walk through any functions in the library...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] WHERE [LibraryID]=" & Str(lID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        FunctionDependencies astrDepends, rs!FunctionID, nFilter
        rs.MoveNext
    Loop

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "mCommon.LibraryDependencies", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UserLibraryItems
'' Description: Determine if library depends on any items in the user library
'' Inputs:      Array of Dependencies, Library ID
'' Returns:     True if Items in User Library, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UserLibraryItems(astrDepends As cGdArray, ByVal lID As Long) As Boolean
On Error GoTo ErrSection:

    astrDepends.Create eGDARRAY_Strings
    
    LibraryDependencies astrDepends, lID, eGDDependencyFilter_UserLibraryOnly
    UserLibraryItems = (astrDepends.Size > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mCommon.UserLibraryItems", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StrategyBasketHasFilter
'' Description: Determine if the given strategy basket contains a filter
'' Inputs:      Strategy Basket ID
'' Returns:     True if has filter, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StrategyBasketHasFilter(ByVal lStrategyBasketID As Long) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBasketItems] " & _
                "WHERE ([StrategyBasketID]=" & Str(lStrategyBasketID) & ") AND ([SymbolGroupID] LIKE 'FIL:*');", dbOpenDynaset)
    
    StrategyBasketHasFilter = Not (rs.BOF And rs.EOF)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mCommon.StrategyBasketHasFilter", eGDRaiseError_Raise, g.strAppPath
    
End Function
