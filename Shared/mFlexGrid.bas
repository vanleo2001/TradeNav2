Attribute VB_Name = "mFlexGrid"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mFlexGrid.bas
'' Description: Common functions to use with flex grids
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/07/2014   DAJ         Created
'' 03/14/2014   DAJ         Added ValidGridRow and ValidGridCol functions
'' 03/20/2014   DAJ         Tweaked ValidGridRow and ValidGridCol; Added EditCell
'' 08/19/2014   DAJ         Added PointFromCell function
'' 10/27/2014   DAJ         Added ToggleCell
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Public Const ALT_GRID_ROW_COLOR As Long = &HC0F0F8
'Public Const ALT_GRID_ROW_COLOR As Long = &HC8F0FF
'Public Const ALT_GRID_ROW_COLOR As Long = &HDDFAF9

Public ALT_GRID_ROW_COLOR As Long

Public Property Get CheckedCell(fg As VSFlexGrid, ByVal nRow&, ByVal nCol&) As Boolean
On Error Resume Next

    Dim i&
    
    i = fg.Cell(flexcpChecked, nRow, nCol)
    If i = flexChecked Then
        CheckedCell = True
    Else
        CheckedCell = False
    End If

End Property

Public Property Let CheckedCell(fg As VSFlexGrid, ByVal nRow&, ByVal nCol&, ByVal bIsChecked As Boolean)
On Error Resume Next

    If bIsChecked Then
        fg.Cell(flexcpChecked, nRow, nCol) = flexChecked
    Else
        fg.Cell(flexcpChecked, nRow, nCol) = flexUnchecked
    End If

End Property

Public Function CheckedMatrix(fg As VSFlexGrid, ByVal nRow&, ByVal nCol&) As Boolean
On Error Resume Next

    If fg.Cell(flexcpChecked, nRow, nCol) = flexChecked Then
        CheckedMatrix = True
    Else
        CheckedMatrix = False
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetGridLevel
'' Description: Set the grid level on the given grid
'' Inputs:      Grid, Level
'' Returns:     New Level
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SetGridLevel(fgGrid As VSFlexGrid, ByVal lLevel As Long) As Long
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lRow As Long                    ' Row in the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim lLowestLevel As Long            ' Lowest level
    Dim lHighestLevel As Long           ' Highest level
    
    With fgGrid
        If .Rows > .FixedRows Then
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            lLowestLevel = Abs(kNullData)
            lHighestLevel = kNullData
            For lIndex = .FixedRows To .Rows - 1
                If .RowOutlineLevel(lIndex) > lHighestLevel Then
                    lHighestLevel = .RowOutlineLevel(lIndex)
                End If
                If .RowOutlineLevel(lIndex) < lLowestLevel Then
                    lLowestLevel = .RowOutlineLevel(lIndex)
                End If
            Next lIndex
            
            If lLevel > lHighestLevel Then
                lLevel = lHighestLevel
            End If
            If lLevel < lLowestLevel Then
                lLevel = lLowestLevel
            End If
            
            For lIndex = .FixedRows To .Rows - 1
                If .RowOutlineLevel(lIndex) < lLevel Then
                    .IsCollapsed(lIndex) = flexOutlineExpanded
                Else
                    .IsCollapsed(lIndex) = flexOutlineCollapsed
                End If
            Next lIndex
            
            .Redraw = nRedraw
        End If
    End With
    
    SetGridLevel = lLevel

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mFlexGrid.SetGridLevel"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetBackColors
'' Description: Set the background color of the rows appropriately
'' Inputs:      Grid that is currently active
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetBackColors(Grid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim bAlt As Boolean                 ' Is this an alternate row?
    Dim lRow As Long                    ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the redraw
    
    With Grid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        For lRow = .FixedRows To .Rows - 1
            If .RowHidden(lRow) = False Then
                If Not bAlt Then
                    .Cell(flexcpBackColor, lRow, 0, lRow, .Cols - 1) = .BackColor
                Else
                    .Cell(flexcpBackColor, lRow, 0, lRow, .Cols - 1) = .BackColorAlternate
                End If
                bAlt = Not bAlt
            End If
        Next lRow
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mFlexGrid.SetBackColors", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidGridRow
'' Description: Determine if the given row is valid in the given grid
'' Inputs:      Grid, Row, Fixed is Valid?
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidGridRow(Grid As VSFlexGrid, Optional ByVal Row As Long = kNullData, Optional ByVal bFixedIsValid As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    With Grid
        If Row = kNullData Then
            Row = .Row
        End If
        
        If bFixedIsValid Then
            bReturn = (Row >= 0) And (Row < .Rows)
        Else
            bReturn = (Row >= .FixedRows) And (Row < .Rows)
        End If
    End With
    
    ValidGridRow = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mFlexGrid.ValidGridRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidGridCol
'' Description: Determine if the given column is valid in the given grid
'' Inputs:      Grid, Col, Fixed is Valid?
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidGridCol(Grid As VSFlexGrid, Optional ByVal Col As Long = kNullData, Optional ByVal bFixedIsValid As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    With Grid
        If Col = kNullData Then
            Col = .Col
        End If
        
        If bFixedIsValid Then
            bReturn = (Col >= 0) And (Col < .Cols)
        Else
            bReturn = (Col >= .FixedCols) And (Col < .Cols)
        End If
    End With
    
    ValidGridCol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mFlexGrid.ValidGridCol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditCell
'' Description: Edit the current or given cell
'' Inputs:      Row, Column, Send F4?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EditCell(Grid As VSFlexGrid, Optional ByVal lRow As Long = kNullData, Optional ByVal lCol As Long = kNullData, Optional ByVal bSendF4 As Boolean = True)
On Error GoTo ErrSection:

    With Grid
        If lRow <> kNullData Then
            .Row = lRow
        End If
        If lCol <> kNullData Then
            .Col = lCol
        End If
        
        .EditCell
        
        If bSendF4 Then
            SendKeys "{F4}"
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mFlexGrid.EditCell"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PointFromCell
'' Description: Determine the point for the given cell information
'' Inputs:      Grid, Row, Column
'' Returns:     Point
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PointFromCell(Grid As VSFlexGrid, ByVal lRow As Long, ByVal lCol As Long) As POINTAPI
On Error GoTo ErrSection:

    Dim ptReturn As POINTAPI            ' Point to return
    
    With Grid
        ptReturn.X = .ColPos(lCol) / Screen.TwipsPerPixelX
        ptReturn.Y = (.RowPos(lRow) + .RowHeight(lRow)) / Screen.TwipsPerPixelY
        
        ClientToScreen .hWnd, ptReturn
        
        ptReturn.X = ptReturn.X * Screen.TwipsPerPixelX
        ptReturn.Y = ptReturn.Y * Screen.TwipsPerPixelY
    End With
    
    PointFromCell = ptReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mFlexGrid.PointFromCell"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToggleCell
'' Description: Toggle the given cell in the grid
'' Inputs:      Grid, Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ToggleCell(fg As VSFlexGrid, ByVal nRow As Long, ByVal nCol As Long)
On Error GoTo ErrSection:

    CheckedCell(fg, nRow, nCol) = Not CheckedCell(fg, nRow, nCol)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mFlexGrid.ToggleCell"
    
End Sub
