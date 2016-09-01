VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeConsoleGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeConsoleGrid.cls
'' Description: Wrapper class for a generic trade console grid
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/15/2010   DAJ         Fixed the grid column information persistence
'' 03/16/2010   DAJ         Added DumpRowInGrid for debugging
'' 06/28/2011   DAJ         Setup clickable cells like hyperlinks
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    fgGrid As VSFlexGrid                ' Grid control
    tmrRealtime As Timer                ' Timer object
    GridColumns As cGridColumns         ' Grid column information
End Type
Private m As mPrivate

Public Property Get Display() As String
    Display = m.GridColumns.ToString
End Property
Public Property Let Display(ByVal strDisplay As String)
    m.GridColumns.FromString strDisplay
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the object
'' Inputs:      Flex Grid Object, Source Name, Display
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(fgGrid As VSFlexGrid, tmrRealtime As Timer, ByVal strDisplay As String)
On Error GoTo ErrSection:

    Set m.fgGrid = fgGrid
    SetColDataToColNum
    m.GridColumns.Init fgGrid
    
    Set m.tmrRealtime = tmrRealtime
    
    Display = strDisplay

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleGrid.Init"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetColDataToColNum
'' Description: Set the column data to the original column number before the
''              column gets moved
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetColDataToColNum()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With m.fgGrid
        For lIndex = 0 To .Cols - 1
            .ColData(lIndex) = lIndex
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleGrid.SetColDataToColNum"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColNumFromColData
'' Description: Determine where the given column is really sitting right now
'' Inputs:      Column Number
'' Returns:     Current Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ColNumFromColData(ByVal Col As Long) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = -1&
    If ValidCol(Col) Then
        For lIndex = 0 To m.fgGrid.Cols - 1
            If m.fgGrid.ColData(lIndex) = Col Then
                lReturn = lIndex
                Exit For
            End If
        Next lIndex
    End If
    
    ColNumFromColData = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeConsoleGrid.ColNumFromColData"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidCol
'' Description: Determine if the given column is valid in the grid
'' Inputs:      Column Number
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidCol(ByVal Col As Long) As Boolean
On Error GoTo ErrSection:

    ValidCol = ((Col >= m.fgGrid.FixedCols) And (Col < m.fgGrid.Cols))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeConsoleGrid.ValidCol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidRow
'' Description: Determine if the given Row is valid in the grid
'' Inputs:      Row Number
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidRow(ByVal Row As Long) As Boolean
On Error GoTo ErrSection:

    ValidRow = ((Row >= m.fgGrid.FixedRows) And (Row < m.fgGrid.Rows))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeConsoleGrid.ValidRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectRow
'' Description: Select the row in the grid if it is valid
'' Inputs:      Row Number
'' Returns:     True if Selected, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SelectRow(ByVal Row As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    
    m.fgGrid.Row = Row
    If ValidRow(Row) Then
        m.fgGrid.RowSel = Row
        bReturn = True
    End If
    
    SelectRow = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeConsoleGrid.SelectRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentRowValid
'' Description: Determine if the current row in the grid is valid
'' Inputs:      None
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentRowValid()
On Error GoTo ErrSection:

    CurrentRowValid = ValidRow(m.fgGrid.Row)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeConsoleGrid.CurrentRowValid"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeCell
'' Description: To change text and forecolor of grid cell
'' Inputs:      Row and Column to change, New Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeCell(ByVal Row As Long, ByVal Col As Long, ByVal strNewText As String)
On Error GoTo ErrSection:

    Dim nForeColor As Long              ' Foreground color to color the cell
    Dim dTickCount As Double            ' Current tick count
    
    With m.fgGrid
        nForeColor = frmQuotes.UnchColor
        
        If .TextMatrix(Row, Col) <> strNewText Then
            .TextMatrix(Row, Col) = strNewText
            
            If m.tmrRealtime.Enabled Then
                nForeColor = frmQuotes.UpdateColor
                .Cell(flexcpForeColor, Row, 0) = frmQuotes.UpdateColor
                .Cell(flexcpData, Row, Col) = gdTickCount
            End If
        ElseIf m.tmrRealtime.Enabled Then
            dTickCount = .Cell(flexcpData, Row, Col)
            dTickCount = gdTickCount - dTickCount
            If dTickCount >= 0 And dTickCount <= 1000 Then
                nForeColor = frmQuotes.UpdateColor
            End If
        End If
        
        .Cell(flexcpForeColor, Row, Col) = nForeColor
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleGrid.ChangeCell"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorCell
'' Description: Color the given cell red or green depending on sign
'' Inputs:      Grid, Row and Column to color
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ColorCell(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim dValue As Double                ' Value of the cell
    Dim lColor As Long

    dValue = ValOfText(m.fgGrid.TextMatrix(Row, Col))
    
    If dValue < 0 Then
        m.fgGrid.Cell(flexcpForeColor, Row, Col) = vbRed
    ElseIf dValue = 0 Then
        If g.nColorTheme = kDarkThemeColor Then
            m.fgGrid.Cell(flexcpForeColor, Row, Col) = vbWhite
        Else
            m.fgGrid.Cell(flexcpForeColor, Row, Col) = vbBlack
        End If
    Else
        lColor = QBColor(2)
        If g.nColorTheme = kDarkThemeColor Then
            If IsBlueRange(lColor) Then
                lColor = vbCyan
            ElseIf IsGreenRange(lColor, True) Then
                lColor = vbGreen
            End If
        End If
        m.fgGrid.Cell(flexcpForeColor, Row, Col) = lColor
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleGrid.ColorCell"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyDisplayString
'' Description: Apply the display string from the grid
'' Inputs:      First Non-Frozen Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ApplyDisplayString(ByVal lMinColumn As Long)
On Error GoTo ErrSection:

    m.GridColumns.ApplyDisplayString lMinColumn

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleGrid.ApplyDisplayString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildDisplayString
'' Description: Build the display string from the grid
'' Inputs:      Exclude Columns, Force Sum Calculation?
'' Returns:     Display String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildDisplayString(ByVal strExcludeColumns As String, Optional ByVal bForceSumCalc As Boolean = False)
On Error GoTo ErrSection:

    m.GridColumns.BuildDisplayString strExcludeColumns, bForceSumCalc

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleGrid.BuildDisplayString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoSizeGrid
'' Description: Do an auto size on the grid unless the user has changed one
''              or more of the column widths manually
'' Inputs:      Force Resize
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AutoSizeGrid(Optional ByVal bForceResize As Boolean = False)
On Error GoTo ErrSection:

    m.GridColumns.AutoSizeGrid bForceResize
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleGrid.AutoSizeGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the print preview form for this grid
'' Inputs:      Arguments passed in from PrintMe, Title
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant, ByVal strTitle As String)
On Error GoTo ErrSection:

    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader

        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .TextAlign = taCenterMiddle
        .Text = strTitle
        .TextAlign = taLeftMiddle
        .Font.Bold = False
        
        .Paragraph = ""
        
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable m.fgGrid
        Else
            .RenderControl = m.fgGrid.hWnd
        End If
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleGrid.GenerateReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpRowInGrid
'' Description: Create a string for the given row in the grid
'' Inputs:      Row
'' Returns:     Delimited String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DumpRowInGrid(ByVal lRow As Long) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrFields As cGdArray          ' Array of fields
    
    Set astrFields = New cGdArray
    astrFields.Create eGDARRAY_Strings
    
    With m.fgGrid
        For lIndex = 0 To .Cols - 1
            astrFields.Add "'" & .TextMatrix(lRow, lIndex) & "'"
        Next lIndex
    End With
    
    DumpRowInGrid = astrFields.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetMouseCursor
'' Description: Set the mouse cursor appropriately for the given form
'' Inputs:      Form
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetMouseCursor(frm As Form)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row that the mouse cursor is currently in
    Dim lMouseCol As Long               ' Column that the mouse cursor is currently in
    
    lMouseRow = m.fgGrid.MouseRow
    lMouseCol = m.fgGrid.MouseCol

    If Screen.MousePointer <> vbHourglass Then
        If (frm.MousePointer = vbDefault) And (ValidRow(lMouseRow) = True) And (ValidCol(lMouseCol) = True) Then
            If m.fgGrid.Cell(flexcpFontUnderline, lMouseRow, lMouseCol) = True Then
                frm.MousePointer = vbCustom
                frm.MouseIcon = Picture16(ToolbarIcon("kHand"))
            End If
        ElseIf frm.MousePointer = vbCustom Then
            If (ValidRow(lMouseRow) = False) Or (ValidCol(lMouseCol) = False) Then
                frm.MousePointer = vbDefault
            ElseIf m.fgGrid.Cell(flexcpFontUnderline, lMouseRow, lMouseCol) = False Then
                frm.MousePointer = vbDefault
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleGrid.SetMouseCursor"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.GridColumns = New cGridColumns

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleGrid.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate the class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.GridColumns = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleGrid.Class_Terminate"
    
End Sub
