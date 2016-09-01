VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGridColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cGridColumns.cls
'' Description: Collection of grid column objects
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/25/2012   DAJ         Option to reposition columns on ToString
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kGridColumnsVersion = 1&

Private Type mPrivate
    GridColumns As cGdTree              ' Collection of grid column objects
    
    fgGrid As VSFlexGrid                ' Grid object
    lVersion As Long                    ' Version number of the string
    lSumColumnWidth As Long             ' Sum of the column widths
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cGridColumn
On Error GoTo ErrSection:

    Dim GridColumn As cGridColumn       ' Return value for the function
    
    Set GridColumn = Nothing
    If m.GridColumns.Exists(vIndexOrKey) Then
        Set GridColumn = m.GridColumns(vIndexOrKey)
    End If
    
    Set Item = GridColumn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cGridColumns.Item.Get"
    
End Property

Public Property Get Count() As Long
    Count = m.GridColumns.Count
End Property

Public Property Get Version() As Long
    Version = m.lVersion
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add a GridColumn object to the collection
'' Inputs:      Grid Column
'' Returns:     Index
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(GridColumn As cGridColumn) As Long
On Error GoTo ErrSection:

    Add = m.GridColumns.Add(GridColumn, GridColumn.Name)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cGridColumns.Add"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the members
'' Inputs:      Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(Grid As VSFlexGrid)
On Error GoTo ErrSection:

    Set m.fgGrid = Grid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGridColumns.Init"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Fill in the collection from a settings string
'' Inputs:      Settings String
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strSettings As String)
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of fields from the settings string
    Dim lIndex As Long                  ' Index into a for loop
    Dim GridColumn As cGridColumn       ' Grid column object
    Dim strColumns As String            ' String of column information
    
    If InStr(strSettings, "|") = 0 Then
        m.lVersion = 0&
        strColumns = strSettings
    Else
        m.lVersion = CLng(Val(Parse(strSettings, "|", 1)))
        strColumns = Parse(strSettings, "|", 2)
    End If
    
    m.GridColumns.Clear
    
    astrFields.SplitFields strColumns, ","
    For lIndex = 0 To astrFields.Size - 1
        Set GridColumn = New cGridColumn
        GridColumn.FromString astrFields(lIndex)
        Add GridColumn
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGridColumns.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Create a settings string from the collection
'' Inputs:      Set Positions?
'' Returns:     Settings String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString(Optional ByVal bSetPositions As Boolean = True) As String
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of fields from the settings string
    Dim lIndex As Long                  ' Index into a for loop
    Dim GridColumn As cGridColumn       ' Grid column object
        
    If bSetPositions Then
        ' Make sure to set the positions here to make sure that they are correct...
        SetPositions
    End If
    
    ' Walk through the collection and get the string version of each object...
    astrFields.Create eGDARRAY_Strings, m.GridColumns.Count
    For lIndex = 1 To m.GridColumns.Count
        Set GridColumn = Item(lIndex)
        If Not GridColumn Is Nothing Then
            astrFields(lIndex - 1) = GridColumn.ToString
        End If
    Next lIndex
    
    ToString = Str(kGridColumnsVersion) & "|" & astrFields.JoinFields(",")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cGridColumns.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPositions
'' Description: Set the positions for each of the objects in the collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetPositions()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim GridColumn As cGridColumn       ' Grid column object
    
    For lIndex = 1 To m.GridColumns.Count
        Set GridColumn = Item(lIndex)
        If Not GridColumn Is Nothing Then
            GridColumn.Position = lIndex - 1
            Set m.GridColumns(lIndex) = GridColumn
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGridColumns.SetPositions"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoSizeGrid
'' Description: Do an auto size on the grid unless the user has changed one
''              or more of the column widths manually
'' Inputs:      Force Resize?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AutoSizeGrid(Optional ByVal bForceResize As Boolean = False)
On Error GoTo ErrSection:

    If Not m.fgGrid Is Nothing Then
        If (m.lSumColumnWidth = 0&) Or (bForceResize = True) Then
            m.fgGrid.AutoSize 0, m.fgGrid.Cols - 1, False, 75
            m.lSumColumnWidth = 0&
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGridColumns.AutoSizeGrid"
    
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

    Dim GridColumn As cGridColumn       ' Grid column object
    Dim lCol As Long                    ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSumColumnWidth As Long         ' Sum of the column widths
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw

    If Not m.fgGrid Is Nothing Then
        With m.fgGrid
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            AutoSizeGrid True
            
            For lIndex = 1 To Count
                Set GridColumn = Item(lIndex)
                If Not GridColumn Is Nothing Then
                    For lCol = 0 To .Cols - 1
                        If .TextMatrix(0, lCol) = GridColumn.Name Then
                            If GridColumn.Width > 0 Then
                                .ColWidth(lCol) = GridColumn.Width
                                lSumColumnWidth = lSumColumnWidth + GridColumn.Width
                            End If
                            .ColHidden(lCol) = Not GridColumn.Visible
                            If (GridColumn.Position < lMinColumn) And (m.lVersion = 0&) Then
                                .ColPosition(lCol) = lMinColumn
                                lMinColumn = lMinColumn + 1
                            Else
                                .ColPosition(lCol) = GridColumn.Position
                            End If
                            
                            Exit For
                        End If
                    Next lCol
                End If
            Next lIndex
            
            .Redraw = nRedraw
        End With
        
        m.lSumColumnWidth = lSumColumnWidth
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGridColumns.ApplyDisplayString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildDisplayString
'' Description: Build the display string from the grid
'' Inputs:      Exclude Columns, Force Sum Calculation?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildDisplayString(ByVal strExcludeColumns As String, Optional ByVal bForceSumCalc As Boolean = False)
On Error GoTo ErrSection:

    Dim GridColumn As cGridColumn       ' Grid column object
    Dim lCol As Long                    ' Index into a for loop
    Dim lSumColumnWidth As Long         ' Sum of the column widths
    
    If Not m.fgGrid Is Nothing Then
        lSumColumnWidth = 0&
        m.GridColumns.Clear
        
        With m.fgGrid
            For lCol = 0 To .Cols - 1
                Set GridColumn = New cGridColumn
                
                GridColumn.Visible = Not .ColHidden(lCol)
                GridColumn.Name = .TextMatrix(0, lCol)
                GridColumn.Position = lCol
                
                If (m.lSumColumnWidth = 0&) And (bForceSumCalc = False) Then
                    GridColumn.Width = 0&
                Else
                    GridColumn.Width = .ColWidth(lCol)
                    lSumColumnWidth = lSumColumnWidth + .ColWidth(lCol)
                End If
                
                GridColumn.ShowInSettings = (InStr(strExcludeColumns, "," & Str(lCol) & ",") = 0)
                
                Add GridColumn
            Next lCol
        End With
        
        m.lSumColumnWidth = lSumColumnWidth
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGridColumns.BuildDisplayString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.GridColumns = New cGdTree
    Set m.fgGrid = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGridColumns.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.GridColumns = Nothing
    Set m.fgGrid = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGridColumns.Class_Terminate"
    
End Sub
