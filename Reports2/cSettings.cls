VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
'==================================================================
'Summary
'-------
'This class manages data entry type grids.  Each row in the grid
'represents a data item.  Column 1 contains labels describing the data
'(i.e. System Name, System Type), Column 2 contains the data which can be
'changed by the user.  A preview box is shown below the data entry grid.
'As you move through data items in the grid, the detailed description
'of the current data item is updated.
'
'Settings File
'-------------
'The template data for building the grid comes from a serialized gdTable called
'gdSettings.dat.  This file is built off of "tblSettings".  If a developer
'changes settings to a report or maintenance form, run the global
'module: "BuildSettingsFile" in order to update "gdSettings.dat".  This
'file must be distributed to users.
'
'Grid column descriptions
'------------------------
' mC_Label      Data item name (ie Rule Name, Rule type, etc.)
' mC_Value      Data entry column.
' mC_Name       Data name (if grid loaded from a table, this would be
'                   the field name from the table, i.e.  SystemName)
' mC_Desc       Data item prevew (shows in the preview box).
' mC_KEYS       Contains key value (if data item combo box type is used)
' mC_Type       Data item type
'                - String       Any characters
'                - Number       Numeric values
'                - Boolean      YES or NO
'                - DateTime     (ie 1/1/2000 8:00am)
'                - ListAdd      Display list, allowed additions
'                - List         Display list, must select from list
'                - TableLookup  Loaded from table (done in calling app).
' mC_DecPos     ** not used anymore **
' mC_Format     Data formatting string (ie "#,###.00", "##0.0%")
' mC_ValFrom    Valid low end value, only used for Number types.
' mC_ValTo      Valid high end value, only used for Number types.
' mC_Length     Highest length allowed for String type
' mC_CanEdit    1=Allow editing, white background, 0=no edits gray background
' mC_Required   1=Data entry is required.
' mC_Show       1=Show in grid, 0=Hide row.
' mC_ShowAdd    1=Show in in grid in Add mode
' mC_Default    Default value initially and most recent value entered.
' mC_DefaultKey Default key value initially and most recent value entered.
' mC_Group      Future used, idea was to use for grouping rows by category
'               like in property window in VB.
'
'Property/Methods
'----------------
' FirstRow      First row number of the application.
' LastRow       last row number of the application.
' Col           Returns Col nbr from grid based on field name.
' Text          Returns all settings concatenated together.  Use in
'                   sub headings of reports to show report options.
' Load          Loads a grid with settings data.  Based on the
'                   application (AppID) and parms passed, a table
'                   will load the grid or defaults are used (reports).
'==================================================================

Private Type mPrivate
    PreviewRtf As RichTextBox
    vsGrid As VSFlexGrid
    lAppID As Long
    strPrimaryTableName As String
    strPrimaryKeyName As String
    dPrimaryKeyValue As Double
    bNewItem As Boolean
    DB As Database
    SettingsFile As cSettingsFile
    
    lFirstRow As Long
    lLastRow As Long
End Type
Private m As mPrivate

'Settings Grid
Private Const mC_Label = 0
Private Const mC_Value = 1
Private Const mC_Name = 2
Private Const mC_Desc = 3
Private Const mC_Keys = 4
Private Const mC_Type = 5
Private Const mC_DecPos = 6
Private Const mC_Format = 7
Private Const mC_ValFrom = 8
Private Const mC_ValTo = 9
Private Const mC_Length = 10
Private Const mC_CanEdit = 11
Private Const mC_Required = 12
Private Const mC_Show = 13
Private Const mC_ShowAdd = 14
Private Const mC_Default = 15
Private Const mC_Group = 16
Private Const mC_Global = 17
Private Const mC_Save = 18
Private Const mC_DefaultKey = 18
Private Const mC_Cols = 19

Public Enum ensg_Settings
     ensg_Label = 0
     ensg_Value = 1
     ensg_Name = 2
     ensg_Desc = 3
     ensg_Keys = 4
     ensg_Type = 5
     ensg_DecPos = 6
     ensg_Format = 7
     ensg_ValFrom = 8
     ensg_ValTo = 9
     ensg_Length = 10
     ensg_CanEdit = 11
     ensg_Required = 12
     ensg_Show = 13
     ensg_ShowAdd = 14
     ensg_Default = 15
     ensg_Group = 16
     ensg_Global = 17
     ensg_Save = 18
     ensg_DefaultKey = 19
End Enum

Property Get FirstRow() As Long
    FirstRow = m.lFirstRow
End Property
Property Get LastRow() As Long
    LastRow = m.lLastRow
End Property

'Returns a given grid column number
Public Function Col(ByVal pName As ensg_Settings) As Long
    Col = pName
End Function

'Set the following properties to use this class
Property Get DB() As Database
    Set DB = m.DB
End Property
Property Let DB(pData As Database)
    Set m.DB = pData
End Property
Property Get SettingsFile() As cSettingsFile
    Set SettingsFile = m.SettingsFile
End Property
Property Let SettingsFile(pData As cSettingsFile)
    Set m.SettingsFile = pData
End Property
Property Get PreviewRtf() As RichTextBox
    Set PreviewRtf = m.PreviewRtf
End Property
Property Let PreviewRtf(pData As RichTextBox)
    Set m.PreviewRtf = pData
End Property
Property Get vsGrid() As VSFlexGrid
    Set vsGrid = m.vsGrid
End Property
Property Let vsGrid(pData As VSFlexGrid)
    Set m.vsGrid = pData
End Property
Property Get NewItem() As Boolean
    NewItem = m.bNewItem
End Property
Property Let NewItem(pData As Boolean)
    m.bNewItem = pData
End Property

'Returns the number of visible rows
Public Function VisibleRows() As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRows As Long                   ' Number of rows that are not hidden
    
    With m.vsGrid
        For lIndex = 1 To .Rows - 1
            If Not .RowHidden(lIndex) Then
                lRows = lRows + 1
            End If
        Next lIndex
    End With
    VisibleRows = lRows
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSettings.VisibleRows", eGDRaiseError_Raise, g.strAppPath

End Function

'This routine concatenates together all displable settings from the grid.
'This is shown at the top of report in the Studies module.
Public Function Text() As String
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim strReturn As String             ' String to return
    
    With m.vsGrid
        For lIndex = 1 To .Rows - 1
            If Not .RowHidden(lIndex) Then
                If .TextMatrix(lIndex, mC_Name) <> "ReportDesc" And _
                   .TextMatrix(lIndex, mC_Name) <> "ReportName" Then
                    strReturn = strReturn & .TextMatrix(lIndex, mC_Label) & " " & _
                            .TextMatrix(lIndex, mC_Value) & ", "
                End If
            End If
        Next lIndex
    End With
    
    strReturn = Trim(strReturn)
    If Right(strReturn, 1) = "," Then
        Text = Left(strReturn, Len(strReturn) - 1)
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSettings.Text", eGDRaiseError_Raise, g.strAppPath

End Function

Public Sub Load(ByVal lAppID As Long, Optional ByVal strPrimaryTableName As String = "", _
    Optional ByVal strPrimaryKeyName As String = "", _
    Optional ByVal dPrimaryKeyValue As Double = 0)
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim rs As Recordset                 ' Recordset into the database
    Dim lRow As Long                    ' Index into a for loop
    
    m.lAppID = lAppID
    m.strPrimaryTableName = strPrimaryTableName
    m.strPrimaryKeyName = strPrimaryKeyName
    m.dPrimaryKeyValue = dPrimaryKeyValue
    
    InitGrid
    
    With m.vsGrid
        .Redraw = flexRDNone
        
        ' Create recordset for data from table
        If Len(strPrimaryTableName) > 0 Then
            Set rs = m.DB.OpenRecordset("SELECT * FROM [" & m.strPrimaryTableName & "] " & _
                    "WHERE [" & m.strPrimaryKeyName & "]=" & Str(m.dPrimaryKeyValue) & " " & _
                    "ORDER BY [" & m.strPrimaryKeyName & "];", dbOpenSnapshot)
            If Not rs.EOF Then
                ' Loop through fields from table and load grid in appropriate spots
                For lIndex = 0 To rs.Fields.Count - 1
                    For lRow = 1 To .Rows - 1
                        If rs.Fields.Item(lIndex).Name = .TextMatrix(lRow, mC_Name) Then
                            If .TextMatrix(lRow, mC_Type) = "TableLookup" Then
                                ShowData .TextMatrix(lRow, mC_Name), _
                                     rs.Fields.Item(lIndex).Value, _
                                     rs.Fields.Item(m.strPrimaryKeyName).Value
                            Else
                                ShowData .TextMatrix(lRow, mC_Name), _
                                     rs.Fields.Item(lIndex).Value, ""
                            End If
                            
                            Exit For
                        End If
                    Next lRow
                Next lIndex
            End If
        Else
            ' Load grids from Settings gdTable and use default value
            'For lIndex = 1 To m.vsGrid.Rows - 1
            '    ShowData .TextMatrix(lIndex, mC_Name), .TextMatrix(lIndex, mC_Default)
            'Next lIndex
        End If
        
        
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
        .Col = 1
        
        'Find first visible row to show
        For lIndex = 1 To .Rows - 1
            If Not .RowHidden(lIndex) Then
                .Row = lIndex
                Exit For
            End If
        Next lIndex
        
        If .Row < 0 Then
            .Row = 0
        End If
        RowColChange
    End With
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSettings.Load", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim bHide As Boolean                ' Whether or not to hide the row
    Dim bFirst As Boolean               ' Is this the first row?
    
    With m.vsGrid
        .Redraw = flexRDNone
        
        .Clear
        .AllowBigSelection = False
        .AllowSelection = False
        .Appearance = flex3DLight
        .BackColorFixed = &HD1E7ED
        .FocusRect = flexFocusSolid
        .HighLight = flexHighlightWithFocus
        .TabBehavior = flexTabCells
        .Editable = True
        .FontName = "MS Sans Serif"
        .FontSize = 8
        .ExtendLastCol = True
        .ExplorerBar = flexExNone
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .Ellipsis = flexEllipsisEnd
        .Cols = mC_Cols
        .Rows = 1
        .FixedCols = 1
        .FixedRows = 1
        
        .TextMatrix(0, mC_Label) = "Item"
        .TextMatrix(0, mC_Value) = "Value"
        
        bFirst = True
        
        ' Load field label names and description
        For lIndex = 1 To m.SettingsFile.NumRecords - 1
            If m.lAppID = m.SettingsFile.Num(ensgd_AppID, lIndex) Then
                
                ' Save first record
                If bFirst Then
                    m.lFirstRow = lIndex
                    bFirst = False
                End If
                
                bHide = False
                If m.bNewItem Then
                    If m.SettingsFile.Num(ensgd_ShowAdd, lIndex) = 0 Then
                        bHide = True
                    End If
                Else
                    If m.SettingsFile.Num(ensgd_ShowEdit, lIndex) = 0 Then
                        bHide = True
                    End If
                End If
                
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, mC_Name) = m.SettingsFile.Item(ensgd_SettingName, lIndex)
                .TextMatrix(.Rows - 1, mC_Label) = NullChk(m.SettingsFile.Item(ensgd_LabelName, lIndex), "")
                .TextMatrix(.Rows - 1, mC_Desc) = NullChk(m.SettingsFile.Item(ensgd_Desc, lIndex), "")
                .TextMatrix(.Rows - 1, mC_Type) = NullChk(m.SettingsFile.Item(ensgd_Type, lIndex), "")
                .TextMatrix(.Rows - 1, mC_Format) = NullChk(m.SettingsFile.Item(ensgd_Formatting, lIndex), "")
                .TextMatrix(.Rows - 1, mC_DecPos) = m.SettingsFile.Item(ensgd_DecPos, lIndex)
                .TextMatrix(.Rows - 1, mC_ValFrom) = m.SettingsFile.Item(ensgd_ValFrom, lIndex)
                .TextMatrix(.Rows - 1, mC_ValTo) = m.SettingsFile.Item(ensgd_ValTo, lIndex)
                .TextMatrix(.Rows - 1, mC_Length) = m.SettingsFile.Item(ensgd_Length, lIndex)
                .TextMatrix(.Rows - 1, mC_CanEdit) = m.SettingsFile.Item(ensgd_CanEdit, lIndex)
                .TextMatrix(.Rows - 1, mC_Required) = m.SettingsFile.Item(ensgd_Required, lIndex)
                .TextMatrix(.Rows - 1, mC_Show) = m.SettingsFile.Item(ensgd_ShowEdit, lIndex)
                .TextMatrix(.Rows - 1, mC_ShowAdd) = m.SettingsFile.Item(ensgd_ShowAdd, lIndex)
                
                'If a default exists, move it to grid
                If IsNull(m.SettingsFile.Item(ensgd_Default, lIndex)) Then
                    .TextMatrix(.Rows - 1, mC_Default) = ""
                Else
                    .TextMatrix(.Rows - 1, mC_Default) = m.SettingsFile.Item(ensgd_Default, lIndex)
                    .TextMatrix(.Rows - 1, mC_DefaultKey) = m.SettingsFile.Item(ensgd_DefaultKey, lIndex)
                End If
                .TextMatrix(.Rows - 1, mC_Value) = NullChk(m.SettingsFile.Item(ensgd_Value, lIndex), "")
                .TextMatrix(.Rows - 1, mC_Group) = NullChk(m.SettingsFile.Item(ensgd_Group, lIndex), "")
                .TextMatrix(.Rows - 1, mC_Global) = Str(NullChk(m.SettingsFile.Num(ensgd_Global, lIndex), 0))
                .TextMatrix(.Rows - 1, mC_Save) = Str(NullChk(m.SettingsFile.Num(ensgd_Save, lIndex), 0))
                
                'Shade cells that can't be changed to gray background
                If m.SettingsFile.Item(ensgd_CanEdit, lIndex) = 1 Then
                    .Cell(flexcpBackColor, .Rows - 1, mC_Value) = vbWhite
                Else
                    .Cell(flexcpBackColor, .Rows - 1, mC_Value) = &HE0E0E0
                End If
                
                'Show default value or Hide row if flagged
                If bHide Then
                    .RowHidden(.Rows - 1) = True
                Else
                    ShowData m.SettingsFile.Item(ensgd_SettingName, lIndex), _
                             m.SettingsFile.Item(ensgd_Value, lIndex), _
                             m.SettingsFile.Item(ensgd_DefaultKey, lIndex)
                End If
                
            Else
                If m.SettingsFile.Num(ensgd_AppID, lIndex) > m.lAppID Then
                    m.lLastRow = lIndex - 1
                    Exit For
                End If
            End If
        Next lIndex
        If m.lLastRow = 0 Then m.lLastRow = m.SettingsFile.NumRecords - 1
        
        ' Hide all columns but label and value...
        For lIndex = 0 To mC_Cols - 1
            .ColHidden(lIndex) = True
        Next lIndex
        .ColHidden(mC_Label) = False
        .ColHidden(mC_Value) = False
        
        .ColAlignment(mC_Value) = flexAlignLeftCenter
        
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
    End With
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSettings.InitGrid", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub ShowData(ByVal strFieldName As String, ByVal vFieldValue As Variant, Optional ByVal vKeyValue As Variant)
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Whether or not the item was found
    Dim bContinue As Boolean            ' Whether or not to continue
    
    With m.vsGrid
        For lIndex = 1 To .Rows - 1
            If .Cell(flexcpText, lIndex, mC_Name) = strFieldName Then
                If m.bNewItem Then
                    If .Cell(flexcpValue, lIndex, mC_ShowAdd) = 1 Then
                        bContinue = True
                    Else
                        bContinue = False
                    End If
                Else
                   If .Cell(flexcpValue, lIndex, mC_Show) = 1 Then
                        bContinue = True
                   Else
                        bContinue = False
                   End If
                End If
                bFound = True
                
                If bContinue Then
                    Select Case .Cell(flexcpText, lIndex, mC_Type)
                        
                        Case "Number"
                            .TextMatrix(lIndex, mC_Value) = _
                                Format(vFieldValue, .Cell(flexcpText, lIndex, mC_Format))
                            
                        Case "Pct"
                            If vFieldValue = "" Or IsNull(vFieldValue) Then
                                .TextMatrix(lIndex, mC_Value) = "0"
                            Else
                                .TextMatrix(lIndex, mC_Value) = _
                                    Format(vFieldValue / 100, .Cell(flexcpText, lIndex, mC_Format))
                            End If
                            
                        Case "Boolean"
                            If vFieldValue = 0 Or vFieldValue = "False" Or vFieldValue = "No" Then
                                .TextMatrix(lIndex, mC_Value) = "No"
                            Else
                                .TextMatrix(lIndex, mC_Value) = "Yes"
                            End If
                        
                        Case "DateTime"
                            If vFieldValue <> "12:00:00 AM" And vFieldValue > 0 Then
                                ' Only display the date, not the time 10/21/2002 DAJ
                                .TextMatrix(lIndex, mC_Value) = DateFormat(vFieldValue) '& " " & Format(vFieldValue, "hh:mm AM/PM")
                            Else
                                .TextMatrix(lIndex, mC_Value) = ""
                            End If
                            
                        Case "ListAdd"
                            If IsNull(vFieldValue) Then
                                .TextMatrix(lIndex, mC_Value) = ""
                            Else
                                .TextMatrix(lIndex, mC_Value) = vFieldValue
                            End If
                            
                        'Note: The keyvalue for Lists is stored in the table.
                        'move this to the KEYS column.  Value column must be
                        'set by the caller.
                        Case "List"
                            If IsNull(vFieldValue) Then
                                .TextMatrix(lIndex, mC_Value) = ""
                            Else
                                If InStr(1, vFieldValue, ";") = 0 Then
                                    .TextMatrix(lIndex, mC_Value) = vFieldValue
                                End If
                                .TextMatrix(lIndex, mC_Keys) = vFieldValue
                            End If
                            
                        Case "TableLookup"
                            If IsNull(vFieldValue) Then
                                .TextMatrix(lIndex, mC_Value) = ""
                                .TextMatrix(lIndex, mC_Keys) = ""
                            Else
                                .TextMatrix(lIndex, mC_Value) = vFieldValue
                                .TextMatrix(lIndex, mC_Keys) = vKeyValue
                            End If
                                                    
                        Case Else
                            If IsNull(vFieldValue) Then
                                .TextMatrix(lIndex, mC_Value) = ""
                            Else
                                .TextMatrix(lIndex, mC_Value) = vFieldValue
                            End If
                            
                    End Select
                    Exit For
                End If
            End If
            
        Next lIndex
    End With

    If Not bFound Then
        Err.Raise gUserErr, , "Field: " & strFieldName & " not found in tblSettings."
    End If
                
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSettings.ShowData", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    With m.vsGrid
        .Redraw = flexRDNone
        
        'Numeric formatting
        'MT:  Add logic to save the date in the Keys column and
        'then make sure formatting kicks in.
        Select Case .Cell(flexcpText, Row, mC_Type)
            Case "Pct", "Number"
                ' If the format is a percent style format and the user did not
                ' type a percent sign at the end of his value, append one so
                ' that the value of the cell does not end up multiplied by 100
                ' DAJ: 8/14/2002
                If Right(.Cell(flexcpText, Row, mC_Format), 1) = "%" Then
                    If Right(.Cell(flexcpText, Row, Col), 1) <> "%" Then
                        .Cell(flexcpText, Row, Col) = .Cell(flexcpText, Row, Col) & "%"
                    End If
                End If
                .TextMatrix(Row, Col) = Format(.Cell(flexcpValue, Row, Col), _
                           .Cell(flexcpText, Row, mC_Format))
            
            Case "DateTime"
                ' If the user has not defined a time, append a 00:00 to the from
                ' date and a 23:59 to the to date
                ' DAJ: 10/8/2002
                ' We are only going to show the date portion in the grid and
                ' handle this in the Item function
                ' DAJ: 10/21/2002
                ''If Not InStr(.TextMatrix(Row, Col), " ") Then
                ''    If .TextMatrix(Row, mC_Name) = "ToDate" Then
                ''        .TextMatrix(Row, Col) = .TextMatrix(Row, Col) & " 11:59 PM"
                ''    ElseIf .TextMatrix(Row, mC_Name) = "FromDate" Then
                ''        .TextMatrix(Row, Col) = .TextMatrix(Row, Col) & " 12:00 AM"
                ''    End If
                ''End If
        
        End Select
        
        If .TextMatrix(Row, mC_Name) = "RiskPct" And .TextMatrix(Row, mC_Type) = "ListAdd" Then
            If ValOfText(.TextMatrix(Row, Col)) <> 0 Then
                If Right(.TextMatrix(Row, Col), 1) <> "%" Then
                    .TextMatrix(Row, Col) = .TextMatrix(Row, Col) & "%"
                End If
                .TextMatrix(Row, Col) = Format(.Cell(flexcpValue, Row, Col), "#,##0.00%")
            End If
        End If
        
        If .TextMatrix(Row, mC_Name) = "Delta" Then
            If ValOfText(.TextMatrix(Row, Col)) <> 0 Then
                .TextMatrix(Row, Col) = Format(ValOfText(.TextMatrix(Row, Col)), "#,##0")
            ElseIf .TextMatrix(Row, Col) <> "Drawdown / 2" Then
                .TextMatrix(Row, Col) = "0"
            End If
        End If
        
        .Redraw = flexRDDirect
    End With
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettings.AfterEdit", eGDRaiseError_Raise, g.strAppPath

End Sub

'Searches each row for a value in the grid.  Specify a colume to search
'using Col.  If strFieldName is a number, look for the row instead of by name
Public Function Item(ByVal strFieldName As String, Optional ByVal Col As ensg_Settings = ensg_Value) As Variant
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim Row As Long
    
    ' Determine row of data
    If ValOfText(strFieldName) > 0 Then
        Row = ValOfText(strFieldName)
    Else
        For lIndex = 1 To m.vsGrid.Rows - 1
            If strFieldName = m.vsGrid.Cell(flexcpText, lIndex, mC_Name) Then
                Row = lIndex
                Exit For
            End If
        Next lIndex
    End If
    
    ' Item not found (Row=0)
    If Row = 0 Then

        Select Case Col
            Case ensg_Label, ensg_Name, ensg_Desc, ensg_Format, ensg_Type
                Item = ""
            Case Else
                Item = 0
        End Select
        
        Exit Function
    Else
        ' If requesting the name return it
        Select Case Col
        
            Case ensg_Label, ensg_Name, ensg_Desc, ensg_Format, ensg_Type
                Item = m.vsGrid.Cell(flexcpText, Row, Col)
        
            Case ensg_Value
                Select Case m.vsGrid.Cell(flexcpText, Row, mC_Type)
                    Case "Number", "Pct"
                        Item = m.vsGrid.Cell(flexcpValue, Row, Col)
                    Case "String", "Button"
                        Item = m.vsGrid.Cell(flexcpText, Row, Col)
                    Case "ListAdd", "List", "TableLookup"
                        If Col = ensg_Keys Then
                            If m.vsGrid.Cell(flexcpText, Row, Col) = "" Then
                                Item = 0
                            Else
                                Item = m.vsGrid.Cell(flexcpText, Row, Col)
                            End If
                        Else
                            Item = m.vsGrid.Cell(flexcpText, Row, Col)
                        End If
                    Case "Boolean"
                        If m.vsGrid.TextMatrix(Row, mC_Value) = "Yes" Then
                            Item = -1
                        Else
                            Item = 0
                        End If
                    Case "DateTime"
                        If Len(m.vsGrid.TextMatrix(Row, mC_Value)) > 0 Then
                            If m.vsGrid.TextMatrix(Row, mC_Name) = "ToDate" Then
                                Item = CDate(m.vsGrid.TextMatrix(Row, Col) & " 11:59PM")
                            Else
                                Item = CDate(m.vsGrid.TextMatrix(Row, Col))
                            End If
                        End If
                End Select
                
            Case Else
                Item = m.vsGrid.Cell(flexcpValue, Row, Col)
        End Select
        
    End If
                
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cSettings.Item", eGDRaiseError_Raise, g.strAppPath
    
End Function

Public Function SetItem(ByVal strFieldName As String, ByVal strValue As String, Optional ByVal Col As ensg_Settings = ensg_Value) As Variant
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.vsGrid.Rows - 1
        If strFieldName = m.vsGrid.Cell(flexcpText, lIndex, mC_Name) Then
            m.vsGrid.TextMatrix(lIndex, Col) = strValue
            Exit For
        End If
    Next lIndex
          
    ' Call showdata to ensure correct format
    ShowData strFieldName, strValue
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSettings.SetItem", eGDRaiseError_Raise, g.strAppPath

End Function

Public Function Row(ByVal strDataItemName As String) As Long
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.vsGrid.Rows - 1
        If strDataItemName = m.vsGrid.Cell(flexcpText, lIndex, mC_Name) Then
            Row = lIndex
            Exit For
        End If
    Next lIndex
                
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSettings.Row", eGDRaiseError_Raise, g.strAppPath

End Function

Public Sub BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim strTemp As String
    
    If Col <> mC_Value Then
        Cancel = True
        Exit Sub
    End If
    
    With m.vsGrid
        .EditMask = ""
        .ComboList = ""
        
        'List box formatting
        Select Case .Cell(flexcpText, Row, mC_Type)
            Case "List", "ListAdd"
                .ComboList = .Cell(flexcpText, Row, mC_Format)
                
                'Customization Example to use in calling event
                'If .Cell(flexcpText, Row, mC_Name) = "Rating" Then
                '    .ComboList = "|#1;LEO" & vbTab & "Leading Edge Group: Could operate in multiple mini-markets.  Want to keep in system." & _
                '        "|" & "#2;PLEO" & vbTab & "Potential LEO: Eligible for growth.  Plan needs to be created to move up to HP1." & _
                '        "|" & "#3;SP" & vbTab & "Steady Performer: Self Explanatory." & _
                '        "|" & "#4;NG" & vbTab & "No Growth: Can't handle more stores or have no desire to handle more or want to leave system." & _
                '        "|" & "#5;UP" & vbTab & "Under Performer: Do not contribute, have a negative impact, don't Group to our standards."
                'End If
           
            Case "TableLookup"
                If .Cell(flexcpText, Row, mC_Type) = "TableLookup" Then
                
                    'Sample for calling event
                    'If .Cell(flexcpText, Row, mC_Name) = "FutureGroup" Then
                        'For lIndex = 1 To gVar.Groups.Count
                        '    strTemp = strTemp & gVar.Groups.Item(lIndex).GroupID & vbTab & _
                        '        gVar.Groups.Item(lIndex).GroupName & "|"
                        'Next lIndex
                        '.ComboList = left(strTemp, Len(strTemp) - 1)
                    'End If
                End If

            Case "Button", "DateTime"
                .ComboList = "..."
                
            Case "Boolean"
                .ComboList = "Yes|No"
                
        End Select
    End With
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettings.BeforeEdit", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Sub CellButtonClick(ByVal Row As Long, ByVal Col As Long, RptForm As Form)
On Error GoTo ErrSection:

    Dim pt As POINTAPI
    Dim dDate As Double

    With m.vsGrid
        pt.X = .ColPos(Col) / Screen.TwipsPerPixelX
        ''pt.Y = (.RowPos(Row) + .RowHeight(Row)) / Screen.TwipsPerPixelY
        pt.Y = (.RowPos(Row) - ((frmEditDate.Height - .RowHeight(Row)) / 2)) / Screen.TwipsPerPixelY
        ClientToScreen .hWnd, pt
        
        'If .Cell(flexcpText, Row, mC_Name) = "SymbolsPath1" Or _
        '   .Cell(flexcpText, Row, mC_Name) = "SymbolsPath2" Then
        '    With CD
        '        .InitDir = gAppsettings.AppPath
        '        .DialogTitle = "Select a symbols group"
        '        .Filter = "Text (*.txt)|*.txt"
        '        .ShowOpen
        '    End With
        '    If Len(CD.FileName) = 0 Then Exit Sub
        '    .TextMatrix(Row, mC_Value) = CD.FileName
        '    SetToolbarBtn "ID_Save", True
        'End If
        If .TextMatrix(Row, mC_Type) = "DateTime" Then
            pt.X = pt.X * Screen.TwipsPerPixelX
            pt.Y = pt.Y * Screen.TwipsPerPixelY
            dDate = CDate(.TextMatrix(Row, Col))
            dDate = frmEditDate.ShowMe(pt.X, pt.Y, dDate, RptForm)
            .TextMatrix(Row, Col) = DateFormat(dDate)
        End If
    End With
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettings.CellButtonClick", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Sub ChangeEdit(Cancel As Boolean)
On Error GoTo ErrSection:

    'Continue if list or combo box
    Cancel = False
    With m.vsGrid
        Select Case .TextMatrix(.Row, mC_Type)
            Case "List", "TableLookup", "Boolean", "ListAdd"
                
                ' need to handle typing in a number (for allowed fields)
                Select Case UCase(.TextMatrix(.Row, mC_Name))
                Case "RISKPCT", "DELTA", "RISKBASEDON"
                    If ValOfText(.EditText) = 0 Then
                        .TextMatrix(.Row, mC_Value) = .EditText
                    Else
                        Cancel = True
                    End If
                Case Else
                    .TextMatrix(.Row, mC_Value) = .EditText
                End Select
            
            Case Else
        
        End Select
    End With
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettings.ChangeEdit", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub RowColChange()
On Error GoTo ErrSection:
    
    Static slPrevRow As Long
    
    'Set background of label to dark blue with white letters
    'Reset previous label to background of white with black letters
    With m.vsGrid
        If slPrevRow <> 0 And slPrevRow <= .Rows - 1 And Not .RowHidden(.Row) Then
            .Cell(flexcpBackColor, slPrevRow, 0, slPrevRow, 0) = vbDefault  'Gray
            .Cell(flexcpForeColor, slPrevRow, 0, slPrevRow, 0) = &H0&   'Black
        End If
        .Cell(flexcpBackColor, .Row, 0, .Row, 0) = &H800000     'Blue
        .Cell(flexcpForeColor, .Row, 0, .Row, 0) = &HFFFFFF     'White
        
        slPrevRow = .Row
    End With
    
    'When showing lists, the preview box is not used.
    If m.PreviewRtf Is Nothing Then Exit Sub
    
    'Show field explanation
    With m.PreviewRtf
        .Text = m.vsGrid.TextMatrix(m.vsGrid.Row, mC_Label) & Chr(13) & Chr(10)
        .SelStart = 0
        .SelLength = Len(m.vsGrid.TextMatrix(m.vsGrid.RowSel, mC_Label))
        .SelBold = True
        .Text = .Text & m.vsGrid.TextMatrix(m.vsGrid.Row, mC_Desc)
        .SelStart = Len(m.vsGrid.TextMatrix(m.vsGrid.Row, mC_Label)) + 2
        .SelLength = Len(m.vsGrid.TextMatrix(m.vsGrid.Row, mC_Desc))
        .SelBold = False
        .SelLength = 0
        .SelStart = 0
    End With
                
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSettings.RowColChange", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    With m.vsGrid
        
        If .Cell(flexcpValue, Row, mC_CanEdit) = 0 Then
            Cancel = True
            Err.Raise gUserErr, , "This value cannot be changed"
        End If
        
        Select Case .Cell(flexcpText, Row, mC_Type)
            Case "String", "Button", "ListAdd"
                
                'String is too long...
                If Len(.EditText) > .Cell(flexcpValue, Row, mC_Length) Then
                    Cancel = True
                    Err.Raise gUserErr, , "Length must be less then or equal to " & .Cell(flexcpText, Row, mC_Length)
                End If
                
                If .TextMatrix(Row, mC_Name) = "Delta" Then
                    If ValOfText(.EditText) < .Cell(flexcpValue, Row, mC_ValFrom) Or _
                       ValOfText(.EditText) > .Cell(flexcpValue, Row, mC_ValTo) Then
                        Cancel = True
                        Err.Raise gUserErr, , "Value must be between: " & _
                            Format(.Cell(flexcpValue, Row, mC_ValFrom), "#,##0") & _
                                   " and " & _
                            Format(.Cell(flexcpValue, Row, mC_ValTo), "#,##0") & "."
                    End If
                End If
                
            Case "Pct", "Number"
                
                'value not valid
                If ValOfText(.EditText) < .Cell(flexcpValue, Row, mC_ValFrom) Or _
                   ValOfText(.EditText) > .Cell(flexcpValue, Row, mC_ValTo) Then
                    Cancel = True
                    Err.Raise gUserErr, , "Value must be between: " & _
                        Format(.Cell(flexcpValue, Row, mC_ValFrom), _
                               .Cell(flexcpText, Row, mC_Format)) & _
                               " and " & _
                        Format(.Cell(flexcpValue, Row, mC_ValTo), _
                               .Cell(flexcpText, Row, mC_Format)) & "."
                End If
                
            Case "List", "TableLookup"
                .TextMatrix(Row, mC_Keys) = .ComboData
                
            Case "DateTime"
                Dim wDate As String
                Dim wTime As String
                Dim wSep  As Long
                wDate = Left(.EditText, 10)
                wSep = InStr(1, .EditText, " ")
                If wSep > 0 Then
                    wTime = Right(.EditText, Len(.EditText) - wSep)
                End If
                wDate = Trim(wDate & " " & wTime)
                If Not IsDate(wDate) Then
                    Cancel = True
                    Err.Raise gUserErr, , "Please enter a valid date"
                End If
            
        End Select
        
    End With
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettings.ValidateEdit", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'Call this in the "Save" routine of caller
Public Sub CheckRequiredFields()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    
    With m.vsGrid
        For lIndex = 1 To .Rows - 1
            If .Cell(flexcpValue, lIndex, mC_Required) = 1 Then
                If Len(.TextMatrix(lIndex, mC_Value)) = 0 Then
                    MoveFocus m.vsGrid
                    m.vsGrid.Row = lIndex
                    m.vsGrid.Col = mC_Value
                    Err.Raise gUserErr, , "Please enter a value for " & .Cell(flexcpText, lIndex, mC_Label)
                End If
            End If
        Next lIndex
    End With
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettings.CheckRequiredFields", eGDRaiseError_Raise, g.strAppPath

End Sub

'Call this to save the settings from the grid back to the gdTable
Public Sub Save()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    
    With m.vsGrid
        For lIndex = 1 To .Rows - 1
            If .TextMatrix(lIndex, mC_Type) = "TableLookup" Then
                m.SettingsFile.Save .TextMatrix(lIndex, mC_Name), .TextMatrix(lIndex, mC_Value), .TextMatrix(lIndex, mC_Keys), m.lFirstRow, m.lLastRow
            ElseIf .TextMatrix(lIndex, mC_Global) = "0" Then
                m.SettingsFile.SetItem m.lAppID, .TextMatrix(lIndex, mC_Name), .TextMatrix(lIndex, mC_Value), ensgd_Value
            Else
                m.SettingsFile.SetItem 0&, .TextMatrix(lIndex, mC_Name), .TextMatrix(lIndex, mC_Value), ensgd_Value
            End If
        Next lIndex
    End With
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettings.Save", eGDRaiseError_Raise, g.strAppPath

End Sub

'Standard for checking for renaming of function, portfolio, etc.
'Pass in Field name ("PortfolioName"), it's value (mPortfolio.PortfolioName),
'Flag whether this is new in the caller (m.bNewItem).
Public Sub CheckForRenaming(ByVal strDescFieldName As String, ByVal strDescFieldValue As String, _
        ByVal strMsgBoxItem As String, ByVal strAction As String)
On Error GoTo ErrSection:
    
    Dim rs  As Recordset
    
    ' If the portfolio name changes, allow user to make a copy
    strAction = ""
    If Len(strDescFieldValue) > 0 Then
        If m.vsGrid.Cell(flexcpText, Row(strDescFieldName), mC_Value) <> strDescFieldValue Or m.bNewItem Then
            ' Make sure single quotes not in name
            If InStr(1, m.vsGrid.Cell(flexcpText, 1, mC_Value), "'") > 0 Then
                MoveFocus m.vsGrid
                m.vsGrid.Row = 1
                m.vsGrid.Col = mC_Value
                Err.Raise gUserErr, , "Single quotes not allowed in Name"
            End If
            
            Set rs = m.DB.OpenRecordset("SELECT * FROM [" & m.strPrimaryTableName & "] " & _
                "WHERE [" & strDescFieldName & "]='" & m.vsGrid.Cell(flexcpText, 1, mC_Value) & "' " & _
                "ORDER BY [" & strDescFieldName & "];", dbOpenSnapshot)
            If rs.EOF Then
                If m.bNewItem Then
                    'pKeyValue = 0
                Else
                    If RenameOrCopy(strMsgBoxItem) = "C" Then
                        strAction = "C"
                    Else
                        strAction = "R"
                    End If
                End If
            Else
                Err.Raise gUserErr, , m.vsGrid.Cell(flexcpText, 1, mC_Value) & " already exists"
            End If
        End If
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cSettings.CheckForRenaming", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub Class_Terminate()
    Set m.SettingsFile = Nothing
End Sub
