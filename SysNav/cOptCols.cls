VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptCols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cOptCols.cls
'' Description: Class for handling optimization columns
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
'' 06/22/2011   DAJ         Changed "# Days" to "# Bars"
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Text
Option Explicit

Private Type mPrivate
    OptCols As Collection
    lNextFieldNbr As Long
    strOptColumn As String
    strOptOrder As String
    iInputStartingPos As Integer
    iAltStartingPos As Integer
End Type
Private m As mPrivate

Property Let OptColumn(pData As String)
    m.strOptColumn = pData
End Property
Property Let OptOrder(pData As String)
    m.strOptOrder = pData
End Property
Property Get OptColumn() As String
    OptColumn = m.strOptColumn
End Property
Property Get OptOrder() As String
    OptOrder = m.strOptOrder
End Property

Property Get InputStartingPos() As Integer
    InputStartingPos = m.iInputStartingPos
End Property
Property Get AltStartingPos() As Integer
    AltStartingPos = m.iAltStartingPos
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Loads the statistic columns from the stats text file
''              to the m.OptCols collection
'' Inputs:      Hide Test# Column?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load(Optional ByVal bHideTestNumCol As Boolean = False)
On Error GoTo ErrSection:
    
    Dim strLine As String               ' Line from the input file
    Dim astrCols As Variant             ' Temporary array of the columns
    Dim fh As Integer                   ' Handle to the input file
    
    Add bHideTestNumCol, "Test#", "Test Number", "#,##0", "Desc", "N", "N", True
    
    fh = FreeFile
    Open App.Path & "\FixedStats.txt" For Input As #fh
        
    'First row contains optimization column
    Line Input #fh, strLine
    astrCols = Split(strLine, vbTab)
    m.strOptColumn = astrCols(0)
    m.strOptOrder = astrCols(1)
    
    Do Until EOF(fh)
        Line Input #fh, strLine
        astrCols = Split(strLine, vbTab)
        
        ' DAJ 08/04/2003: Need the following check to stay in because of a bug created
        ' when we changed the term System to Strategy...
        If UCase(astrCols(1)) <> "SYSTEM" Then
            Add astrCols(0), astrCols(1), astrCols(2), astrCols(3), astrCols(4), astrCols(5), astrCols(6), False
        End If
        
        If EOF(fh) Then Exit Do
    Loop
    
ErrExit:
    Close #fh
    Exit Sub

ErrSection:
    Close #fh
    RaiseError "cOptCols.Load", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveCols
'' Description: Saves from the fields collection and uses the current order
''              of the grid columns
'' Inputs:      Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveCols(pvsGrid As VSFlexGrid)
On Error GoTo ErrSection:
    
    Dim lCol As Long                    ' Column in the grid
    Dim astrColInfo(6) As Variant       ' Array of column information
    Dim strTextLine As Variant          ' Line to output to the file
    Dim OptCol As cOptCol               ' Optimization Column
    Dim strColName As String            ' Column Name
    Dim fh As Integer                   ' Handle to the output file
    
    fh = FreeFile
    Open AddSlash(App.Path) & "FixedStats.txt" For Output As #fh
    
    ' Save optimization column in row 1 of text file
    astrColInfo(0) = m.strOptColumn
    astrColInfo(1) = m.strOptOrder
    strTextLine = Join(astrColInfo(), vbTab)
    Print #fh, strTextLine
    
    ' Start with column after "Test#" (column 0)
    For lCol = 1 To pvsGrid.Cols - 1
        strColName = pvsGrid.TextMatrix(0, lCol)
        If Not Left(strColName, 2) = "I-" And _
           Not Left(strColName, 2) = "R-" And _
           Not strColName = "Strategy" And _
           Not strColName = "Symbol" And _
           Not strColName = "Period" And _
           Not strColName = "FromDate" And _
           Not strColName = "ToDate" And _
           Not strColName = "ToEnd" And _
           Not strColName = "Custom Inputs" And _
           Not strColName = "Filenames" And _
           Not strColName = "# Bars" Then
            Set OptCol = Item(strColName)
            astrColInfo(0) = OptCol.Hide
            astrColInfo(1) = OptCol.FieldName
            astrColInfo(2) = OptCol.FieldDesc
            astrColInfo(3) = OptCol.FieldFormat
            astrColInfo(4) = OptCol.FieldOrder
            astrColInfo(5) = OptCol.Operator
            astrColInfo(6) = OptCol.FilterValue
            strTextLine = Join(astrColInfo(), vbTab)
            Print #fh, strTextLine
        End If
    Next lCol
    
ErrExit:
    Set OptCol = Nothing
    Close #fh
    Exit Sub
    
ErrSection:
    Set OptCol = Nothing
    Close #fh
    RaiseError "cOptCols.SaveCols", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Saves from the fields collection to the text file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrOptInfo(6) As Variant       ' Array of opt col information
    Dim strTextLine As String           ' Line to output to the file
    Dim fh As Integer                   ' Handle to the output file
    
    fh = FreeFile
    Open AddSlash(App.Path) & "FixedStats.txt" For Output As #fh
    
    'Save optimization stat
    astrOptInfo(0) = m.strOptColumn
    astrOptInfo(1) = m.strOptOrder
    strTextLine = Join(astrOptInfo(), vbTab)
    Print #fh, strTextLine
    
    'Start with item after "Test#" (item 2)
    For lIndex = 2 To Count
        With m.OptCols.Item(lIndex)
            astrOptInfo(0) = .Hide
            astrOptInfo(1) = .FieldName
            astrOptInfo(2) = .FieldDesc
            astrOptInfo(3) = .FieldFormat
            astrOptInfo(4) = .FieldOrder
            astrOptInfo(5) = .Operator
            
            ' DAJ 2/22/2002: If the field is a percent and the user typed
            ' in a number greater than one, divide it by 100 to get the
            ' percent value.
            If InStr(.FieldFormat, "%") And .FilterValue > 1 Then
                .FilterValue = .FilterValue / 100
            End If
            astrOptInfo(6) = .FilterValue
        End With
        strTextLine = Join(astrOptInfo(), vbTab)
        Print #fh, strTextLine
    Next lIndex
    
ErrExit:
    Close #fh
    Exit Sub
    
ErrSection:
    Close #fh
    RaiseError "cOptCols.Save", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add an optimization column to the collection
'' Inputs:      Optimization Column Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add(pHide As Variant, pFieldName As Variant, _
    pFieldDesc As Variant, pFormat As Variant, pFieldOrder As Variant, _
    pOperator As Variant, pFilterValue As Variant, pOptimized As Boolean)
On Error GoTo ErrSection:
    
    Dim OptCol As cOptCol
    
    Set OptCol = New cOptCol
    With OptCol
        .Hide = pHide
        .FieldName = pFieldName
        .FieldDesc = pFieldDesc
        m.lNextFieldNbr = m.lNextFieldNbr + 1
        .FieldNumber = m.lNextFieldNbr
        .FieldFormat = pFormat
        .FieldOrder = pFieldOrder
        If pOperator <> "N" Then .Operator = pOperator
        If pFilterValue <> "N" Then .FilterValue = pFilterValue
        .Optimized = pOptimized
        m.OptCols.Add OptCol, pFieldName
    End With
    Set OptCol = Nothing
    
    'Store the 1st user input...
    If m.iInputStartingPos = 0 Then
        If (UCase(Left(pFieldName, 2)) = "I-") Or (UCase(pFieldName) = "# BARS") Then
            m.iInputStartingPos = Count
        End If
    End If
    
    'Store the 1st Alt rule...
    If m.iAltStartingPos = 0 Then
        If Left(pFieldName, 2) = "R-" Then
            m.iAltStartingPos = Count
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    Set OptCol = Nothing
    RaiseError "cOptCols.Add", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Found
'' Description: Is the given ID in the collection?
'' Inputs:      ID to find
'' Returns:     True if Found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Found(pID As Long) As Boolean
On Error GoTo ErrSection:

    Dim OptCol As cOptCol
    
    Found = True
    Set OptCol = Item(CStr(pID))

ErrExit:
    Exit Property

ErrSection:
    If Err.Number = 91 Or Err.Number = 5 Then
        Found = False
        Resume Next
    Else
        RaiseError "cOptCols.Found", eGDRaiseError_Raise
    End If

End Property

Property Get FiltersExist() As Boolean
    
    Dim X       As Integer
    
    FiltersExist = False
    For X = 1 To Count
        If Item(X).Operator <> "" Then
            FiltersExist = True
            Exit Property
        End If
    Next X

End Property

Property Get Item(ByVal pKey As Variant) As cOptCol
    If pKey = "Game" Then pKey = 1
    Set Item = m.OptCols(pKey)
End Property

Property Get Count() As Long
    Count = m.OptCols.Count
End Property

Private Sub Class_Initialize()
    Set m.OptCols = New Collection
End Sub

Private Sub Class_Terminate()
    Set m.OptCols = Nothing
End Sub
