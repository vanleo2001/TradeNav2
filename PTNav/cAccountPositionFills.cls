VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccountPositionFills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAccountPositionFills.cls
'' Description: Wrapper class and handler for an account position fills
''
'' Author:      Genesis Financial Data Services
''              425 Windchime Pl
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDFields
    eGDField_FillID = 0
    eGDField_FillDate
    eGDField_Flat
    eGDField_Reverse
    eGDField_NumFields
End Enum

Private Type mPrivate
    lAccountPositionID As Long          ' Account Position ID for the fills
    tblAccountPositionFills As cGdTable ' Table of account position fills
    alSortedIndex As cGdArray           ' Sorted index for the fill date field
    Fills As cGdTree                    ' Collection of fill objects
    lLastFlat As Long                   ' Index of last flat in the table
    lLastReverse As Long                ' Index of last reverse in the table
End Type
Private m As mPrivate

Private Function FieldNum(ByVal lField As eGDFields) As Long
    FieldNum = lField
End Function

Public Property Get AccountPositionID() As Long
    AccountPositionID = m.lAccountPositionID
End Property
Public Property Let AccountPositionID(ByVal lAccountPositionID As Long)
    m.lAccountPositionID = lAccountPositionID
End Property

Private Property Get TableNum(ByVal nField As eGDFields, ByVal lRecord As Long) As Double
    TableNum = m.tblAccountPositionFills.Num(nField, lRecord)
End Property
Private Property Let TableNum(ByVal nField As eGDFields, ByVal lRecord As Long, ByVal dNumber As Double)
    m.tblAccountPositionFills.Num(nField, lRecord) = dNumber
End Property

Private Property Get SortedNum(ByVal nField As eGDFields, ByVal lSortedRecord As Long) As Double
    SortedNum = m.tblAccountPositionFills.Num(nField, m.alSortedIndex(lSortedRecord))
End Property
Private Property Let SortedNum(ByVal nField As eGDFields, ByVal lSortedRecord As Long, ByVal dNumber As Double)
    m.tblAccountPositionFills.Num(nField, m.alSortedIndex(lSortedRecord)) = dNumber
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the account position fills for the given account position
'' Inputs:      Account Position ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load(ByVal lAccountPositionID As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    m.lAccountPositionID = lAccountPositionID
    
    m.tblAccountPositionFills.Clear
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionFills] " & _
                "WHERE [AccountPositionID]=" & Str(lAccountPositionID) & " " & _
                "ORDER BY [FillDate];", dbOpenDynaset)
    Do While Not rs.EOF
        m.tblAccountPositionFills.AddRecord Str(rs!FillID) & vbTab & Str(rs!FillDate) & vbTab & Str(rs!Flat) & vbTab & Str(rs!Reverse), , vbTab
        rs.MoveNext
    Loop
    
    RebuildSortedIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositionFills.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LastFill
'' Description: Get the last fill for the account position
'' Inputs:      None
'' Returns:     Last Fill (or Nothing if no fills)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastFill() As cPtFill
On Error GoTo ErrSection:

    Dim Fill As New cPtFill             ' Fill object to return

    If m.alSortedIndex.Size = 0 Then
        Set LastFill = Nothing
    Else
        If Fill.Load(SortedNum(eGDField_FillID, m.alSortedIndex.Size - 1)) Then
            Set LastFill = Fill
        Else
            Set LastFill = Nothing
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPositionFills.LastFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateAll
'' Description: Recalculate all of the values from the beginning
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateAll()
On Error GoTo ErrSection:

    Recalculate 0&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositionFills.RecalculateAll"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateFromDate
'' Description: Recalculate the account position from a certain date/time
'' Inputs:      Date/Time
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateFromDate(ByVal dDateTime As Double, Optional ByVal bReload As Boolean = True)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in array to start
    Dim lIndex As Long                  ' Index into a for loop
    Dim lStartingFillID As Long         ' Fill ID to start the recalculation

    ' May need to reload the table since fills could have been added or removed...
    If bReload Then Load m.lAccountPositionID

    lStartingFillID = -1&

    m.tblAccountPositionFills.SearchAsIndex m.alSortedIndex, FieldNum(eGDField_FillDate), dDateTime, lPos
    For lIndex = lPos To 0 Step -1
        If (SortedNum(eGDField_Flat, lIndex) <> 0) Or (SortedNum(eGDField_Reverse, lIndex) <> 0) Then
            lStartingFillID = SortedNum(eGDField_FillID, lIndex)
            Exit For
        End If
    Next lIndex
    
    Recalculate lStartingFillID
    
frmTest2.AddList "RecalculateFromDate(" & DateFormat(dDateTime, MM_DD_YY, HH_MM_SS, AMPM_UPPER) & ", " & Str(bReload) & "): " & Str(m.tblAccountPositionFills.NumRecords) & ", " & Str(m.alSortedIndex.Size)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositionFills.RecalculateFromDate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFill
'' Description: Update the fill in the account position
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateFill(Fill As cPtFill, Optional ByVal bFillChanged As Boolean = True)
On Error GoTo ErrSection:

    Dim lTableIndex As Long             ' Index into the table for the fill
    Dim lPos As Long                    ' Position to insert fill
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRecalcStart As Long            ' Fill ID for the place to start the recalc

    lTableIndex = IndexForFillID(Fill.FillID)
    If (lTableIndex = 0) Or (bFillChanged = True) Then
        If lTableIndex >= 0 Then
            If Fill.FillDate <> SortedNum(eGDField_FillDate, lTableIndex) Then
                SortedNum(eGDField_FillDate, lTableIndex) = Fill.FillDate
                Save m.alSortedIndex(lTableIndex)
                
                RebuildSortedIndex
                
                lTableIndex = IndexForFillID(Fill.FillID)
            End If
        Else
            m.tblAccountPositionFills.AddRecord Str(Fill.FillID) & vbTab & Str(Fill.FillDate) & vbTab & "0" & vbTab & "0"
            Save m.tblAccountPositionFills.NumRecords - 1
            
            RebuildSortedIndex
            
            lTableIndex = IndexForFillID(Fill.FillID)
        End If
        
        If lTableIndex = 0& Then
            lRecalcStart = 0&
        Else
            For lIndex = lTableIndex - 1 To 0 Step -1
                If SortedNum(eGDField_Flat, lIndex) <> 0 Then
                    lRecalcStart = SortedNum(eGDField_FillID, lIndex)
                    Exit For
                ElseIf SortedNum(eGDField_Reverse, lIndex) <> 0 Then
                    lRecalcStart = SortedNum(eGDField_FillID, lIndex)
                    Exit For
                End If
            Next lIndex
        End If
        
        Recalculate lRecalcStart
        Fill.Load Fill.FillID
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositionFills.UpdateFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFill
'' Description: Remove the fill from the account position
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveFill(Fill As cPtFill)
On Error GoTo ErrSection:

    Dim lTableIndex As Long             ' Index into the table for the fill
    Dim lRecalcStart As Long            ' Fill ID for the place to start the recalc
    Dim lIndex As Long                  ' Index into a for loop

    lTableIndex = IndexForFillID(Fill.FillID)
    If lTableIndex >= 0 Then
        If (lTableIndex = 0&) Then
            lRecalcStart = 0&
        ElseIf (lTableIndex = (m.alSortedIndex.Size - 1)) Then
            lRecalcStart = -1&
        Else
            For lIndex = lTableIndex To 0 Step -1
                If SortedNum(eGDField_Flat, lIndex) <> 0 Then
                    lRecalcStart = SortedNum(eGDField_FillID, lIndex)
                    Exit For
                ElseIf SortedNum(eGDField_Reverse, lIndex) <> 0 Then
                    lRecalcStart = SortedNum(eGDField_FillID, lIndex)
                    Exit For
                End If
            Next lIndex
        End If
        
        m.tblAccountPositionFills.RemoveRecords m.alSortedIndex(lTableIndex)
        RebuildSortedIndex
        
        If lRecalcStart >= 0 Then
            Recalculate lRecalcStart
            Fill.Load Fill.FillID
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositionFills.RemoveFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IndexForFillID
'' Description: Find the index in the sorted array for the given Fill ID
'' Inputs:      Fill ID
'' Returns:     Index in sorted array (or -1 if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IndexForFillID(ByVal lFillID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value from the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = -1&
    
    If lFillID > 0& Then
        For lIndex = 0 To m.alSortedIndex.Size - 1
            If SortedNum(eGDField_FillID, lIndex) = lFillID Then
                lReturn = lIndex
                Exit For
            End If
        Next lIndex
    Else
        lReturn = 0&
    End If
    
    IndexForFillID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPositionFills.IndexForFillID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFills
'' Description: Load the fills starting with the given Fill ID
'' Inputs:      Starting Fill ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFills(ByVal lStartingFillID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lStartingIndex As Long          ' Index in the array to start
    Dim Fill As cPtFill                 ' Fill object to add to the collection

    ' Clear out the fills collection...
    m.Fills.Clear
    
    ' Determine the index of the fill to start with...
    lStartingIndex = IndexForFillID(lStartingFillID)
    
    ' Load fills from the starting index to the end...
    If lStartingIndex >= 0 Then
        For lIndex = lStartingIndex To m.alSortedIndex.Size - 1
            Set Fill = New cPtFill
            If Fill.Load(SortedNum(eGDField_FillID, lIndex)) Then
                m.Fills.Add Fill, Str(Fill.FillID)
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositionFills.LoadFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the entry at the given index to the database
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save(ByVal lIndex As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionFills] " & _
                "WHERE [AccountPositionID]=" & Str(m.lAccountPositionID) & " " & _
                "AND [FillID]=" & Str(TableNum(eGDField_FillID, lIndex)) & ";", dbOpenDynaset)
    If (rs.BOF And rs.EOF) Then
        rs.AddNew
    Else
        rs.Edit
    End If
    
    rs!AccountPositionID = m.lAccountPositionID
    rs!FillID = TableNum(eGDField_FillID, lIndex)
    rs!FillDate = TableNum(eGDField_FillDate, lIndex)
    rs!Flat = TableNum(eGDField_Flat, lIndex)
    rs!Reverse = TableNum(eGDField_Reverse, lIndex)
    
    rs.Update
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositionFills.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Recalculate
'' Description: Recalculate all of the values from the starting fill
'' Inputs:      Starting Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Recalculate(ByVal lStartingFillID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object
    Dim Fill2 As cPtFill                ' Fill object
    Dim lCurrentPosition As Long        ' Current position after the fill
    Dim bIsEntry As Boolean             ' Is this fill an entry?
    Dim bIsReverse As Boolean           ' Is this fill a reversal?
    Dim dAverageEntry As Double         ' Average entry price after this fill
    Dim lSessionDate As Long            ' Session date for the fill
    Dim lSessionQuantity As Long        ' Session quantity after the fill
    Dim dSessionProfit As Double        ' Session profit after the fill
    Dim Bars As New cGdBars             ' Bars object
    Dim nAcctType As eTT_AccountType    ' Account type for the fill
    Dim Entries As New cGdTree          ' Collection of entries
    Dim lClosedQuantity As Long         ' Quantity of the fill that is closed
    Dim dTickValuePerMove As Double     ' Tick value per move
    Dim lPrevQuantity As Long           ' Quantity previous to this fill
    Dim lStartingIndex As Long          ' Index for the starting fill ID

Dim lNumSaves As Long
Dim lNumSaves2 As Long

gdResetProfiles 950, 955
gdStartProfile 950

    ' Load up all of the fills...
gdStartProfile 953
    LoadFills lStartingFillID
gdStopProfile 953
    
    If m.Fills.Count > 0 Then
        SetBarProperties Bars, m.Fills(1).SymbolOrSymbolID
        nAcctType = g.Broker.AccountTypeForID(m.Fills(1).AccountID)
        If Bars.Prop(eBARS_TickMove) = 0 Then
            dTickValuePerMove = 0#
        Else
            dTickValuePerMove = Bars.Prop(eBARS_TickValue) / Bars.Prop(eBARS_TickMove)
        End If
        lStartingIndex = IndexForFillID(lStartingFillID)
    
        ' Walk through the fills recalculating the session and profit fields for each...
        For lIndex = lStartingIndex To m.alSortedIndex.Size - 1
            Set Fill = m.Fills(Str(SortedNum(eGDField_FillID, lIndex)))
            If Not Fill Is Nothing Then
                With Fill
                    If (lIndex = lStartingIndex) And (lStartingFillID > 0) Then
                        If .CurrentPosition = 0 Then
                            bIsEntry = False
                            bIsReverse = False
                        Else
                            bIsEntry = True
                            bIsReverse = True
                            .NumOpen = Abs(.CurrentPosition)
gdStartProfile 951
                            .Save
gdStopProfile 951
lNumSaves = lNumSaves + 1
                        End If
                        
                        lCurrentPosition = .CurrentPosition
                        lSessionDate = .SessionDate
                        lSessionQuantity = .SessionQuantity
                        dSessionProfit = .SessionProfit
                        dAverageEntry = .AverageEntry
                    Else
                        .NumOpen = 0&
                    
                        If .Buy Then
                            If lCurrentPosition >= 0 Then
                                .NumOpen = .Quantity
                                bIsEntry = True
                                bIsReverse = False
                            ElseIf lCurrentPosition + .Quantity > 0 Then
                                .NumOpen = .Quantity - Abs(lCurrentPosition)
                                bIsEntry = True
                                bIsReverse = True
                            Else
                                bIsEntry = False
                                bIsReverse = False
                            End If
                            
                            If bIsReverse And (lIndex = lStartingIndex) Then
                                lCurrentPosition = .CurrentPosition
                                .NumOpen = Abs(.CurrentPosition)
                            Else
                                lCurrentPosition = lCurrentPosition + .Quantity
                            End If
                        Else
                            If lCurrentPosition <= 0 Then
                                .NumOpen = .Quantity
                                bIsEntry = True
                                bIsReverse = False
                            ElseIf lCurrentPosition - .Quantity < 0 Then
                                .NumOpen = .Quantity - Abs(lCurrentPosition)
                                bIsEntry = True
                                bIsReverse = True
                            Else
                                bIsEntry = False
                                bIsReverse = False
                            End If
                            
                            If bIsReverse And (lIndex = lStartingIndex) Then
                                lCurrentPosition = .CurrentPosition
                                .NumOpen = Abs(.CurrentPosition)
                            Else
                                lCurrentPosition = lCurrentPosition - .Quantity
                            End If
                        End If
                        
                        If bIsEntry And (Not bIsReverse) Then
                            .ClosedProfit = 0#
                        ElseIf Not (bIsReverse And (lIndex = lStartingIndex)) Then
                            lClosedQuantity = .Quantity - .NumOpen
                            
                            .ClosedProfit = 0#
                            For lIndex2 = Entries.Count To 1 Step -1
                                Set Fill2 = Entries(lIndex2)
                                If Fill2.NumOpen >= lClosedQuantity Then
                                    If Fill2.Buy Then
                                        .ClosedProfit = .ClosedProfit + ((.Price - Fill2.Price) * dTickValuePerMove * lClosedQuantity)
                                    Else
                                        .ClosedProfit = .ClosedProfit + ((Fill2.Price - .Price) * dTickValuePerMove * lClosedQuantity)
                                    End If
                                    Fill2.NumOpen = Fill2.NumOpen - lClosedQuantity
gdStartProfile 951
                                    Fill2.Save
gdStopProfile 951
lNumSaves = lNumSaves + 1
                                    lClosedQuantity = 0&
                                Else
                                    If Fill2.Buy Then
                                        .ClosedProfit = .ClosedProfit + ((.Price - Fill2.Price) * dTickValuePerMove * Fill2.NumOpen)
                                    Else
                                        .ClosedProfit = .ClosedProfit + ((Fill2.Price - .Price) * dTickValuePerMove * Fill2.NumOpen)
                                    End If
                                    lClosedQuantity = lClosedQuantity - Fill2.NumOpen
                                    Fill2.NumOpen = 0&
gdStartProfile 951
                                    Fill2.Save
gdStopProfile 951
lNumSaves = lNumSaves + 1
                                End If
                                
                                If Fill2.NumOpen = 0 Then Entries.Remove lIndex2
                                If lClosedQuantity = 0& Then Exit For
                            Next lIndex2
                        End If
                        
                        dAverageEntry = 0#
                        lPrevQuantity = 0&
                        For lIndex2 = 1 To Entries.Count
                            Set Fill2 = Entries(lIndex2)
                            With Fill2
                                If .NumOpen > 0 Then
                                    dAverageEntry = ((dAverageEntry * lPrevQuantity) + (.Price * .NumOpen)) / (lPrevQuantity + .NumOpen)
                                    lPrevQuantity = lPrevQuantity + .NumOpen
                                End If
                            End With
                        Next lIndex2
                        If .NumOpen > 0 Then
                            .AverageEntry = ((dAverageEntry * lPrevQuantity) + (.Price * .NumOpen)) / (lPrevQuantity + .NumOpen)
                        Else
                            .AverageEntry = dAverageEntry
                        End If
                    
                        .SessionDate = Bars.SessionDateForTime(ConvertBrokerDate(Fill.FillDate, nAcctType, Bars.Prop(eBARS_Symbol), False))
                        If .SessionDate = lSessionDate Then
                            lSessionQuantity = lSessionQuantity + .Quantity
                            dSessionProfit = dSessionProfit + .ClosedProfit
                        Else
                            lSessionDate = .SessionDate
                            lSessionQuantity = .Quantity
                            dSessionProfit = .ClosedProfit
                        End If
        
                        .CurrentPosition = lCurrentPosition
                        .SessionQuantity = lSessionQuantity
                        .SessionProfit = dSessionProfit
                        
gdStartProfile 951
                        .Save
gdStopProfile 951
lNumSaves = lNumSaves + 1
                    End If
                    
                    If bIsEntry Then Entries.Add Fill, Str(Fill.FillID)
                End With
            
                If lCurrentPosition = 0 Then
                    SortedNum(eGDField_Flat, lIndex) = 1
                    SortedNum(eGDField_Reverse, lIndex) = 0
                ElseIf bIsReverse Then
                    SortedNum(eGDField_Flat, lIndex) = 0
                    SortedNum(eGDField_Reverse, lIndex) = 1
                Else
                    SortedNum(eGDField_Flat, lIndex) = 0
                    SortedNum(eGDField_Reverse, lIndex) = 0
                End If
gdStartProfile 952
                Save m.alSortedIndex(lIndex)
gdStopProfile 952
lNumSaves2 = lNumSaves2 + 1
            End If
        Next lIndex
    End If

frmTest2.AddList vbTab & "Recalculate: " & Str(m.alSortedIndex.Size) & vbTab & Str(lNumSaves) & vbTab & Str(lNumSaves2)

gdStopProfile 950

For lIndex = 950 To 953
    frmTest2.AddList vbTab & gdGetProfiles(lIndex, lIndex)
Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositionFills.Recalculate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RebuildSortedIndex
'' Description: Rebuild the sorted index for the table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RebuildSortedIndex()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    Set m.alSortedIndex = m.tblAccountPositionFills.CreateSortedIndex(FieldNum(eGDField_FillDate), , FieldNum(eGDField_FillID))
    
    For lIndex = 0 To m.alSortedIndex.Size - 1
        If SortedNum(eGDField_Flat, lIndex) <> 0 Then m.lLastFlat = lIndex
        If SortedNum(eGDField_Reverse, lIndex) <> 0 Then m.lLastReverse = lIndex
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositionFills.RebuildSortedIndex"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the class object is created
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.tblAccountPositionFills = New cGdTable
    With m.tblAccountPositionFills
        .CreateField eGDARRAY_Longs, FieldNum(eGDField_FillID), "FillID"
        .CreateField eGDARRAY_Doubles, FieldNum(eGDField_FillDate), "FillDate"
        .CreateField eGDARRAY_TinyInts, FieldNum(eGDField_Flat), "Flat"
        .CreateField eGDARRAY_TinyInts, FieldNum(eGDField_Reverse), "Reverse"
    End With
    
    Set m.alSortedIndex = New cGdArray
    Set m.Fills = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositionFills.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when the class object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.tblAccountPositionFills = Nothing
    Set m.Fills = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositionFills.Class_Terminate"
    
End Sub
