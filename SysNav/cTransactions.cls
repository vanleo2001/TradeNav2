VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTransactions.cls
'' Description: Split out a trades array into transactions and perform
''              calculations on the transactions
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    tblTransactions As cGdTable         ' Table of transactions
    hTransactions As Long               ' Handle for the transactions table
    
    EquityFilter As cEquityFilter       ' Equity filter information
    nTakeNextTrade As eGDTakeNextTradeValue ' Take the next trade?
End Type
Private m As mPrivate

Public Property Get TblNum(ByVal nField As eGDTransactionField, Optional ByVal lRecord As Long = -99999) As Double
    If NumRecords = 0 Then
        If nField = eGDTransactionField_Link Then
            TblNum = -1
        Else
            TblNum = 0
        End If
    Else
        If lRecord = -99999 Then lRecord = NumRecords - 1
        TblNum = gdGetTableNum(m.hTransactions, nField, lRecord)
    End If
End Property
Public Property Let TblNum(ByVal nField As eGDTransactionField, Optional ByVal lRecord As Long = -99999, ByVal dValue As Double)
    If lRecord = -99999 Then lRecord = NumRecords - 1
    gdSetTableNum m.hTransactions, nField, lRecord, dValue
End Property

Public Property Get TblStr(ByVal nField As eGDTransactionField, Optional ByVal lRecord As Long = -99999) As String
    If NumRecords = 0 Then
        TblStr = ""
    Else
        If lRecord = -99999 Then lRecord = NumRecords - 1
        TblStr = gdGetTableString(m.hTransactions, nField, lRecord)
    End If
End Property
Public Property Let TblStr(ByVal nField As eGDTransactionField, Optional ByVal lRecord As Long = -99999, ByVal strValue As String)
    If lRecord = -99999 Then lRecord = NumRecords - 1
    gdSetTableStr m.hTransactions, nField, lRecord, strValue
End Property

Public Property Get Item(ByVal nField As eGDTransactionField, ByVal lRecord As Long) As Variant
    Item = m.tblTransactions.Item(nField, lRecord)
End Property
Public Property Let Item(ByVal nField As eGDTransactionField, ByVal lRecord As Long, ByVal vValue As Variant)
    m.tblTransactions.Item(nField, lRecord) = vValue
End Property

Public Property Get EquityFilter() As cEquityFilter
    Set EquityFilter = m.EquityFilter
End Property
Public Property Let EquityFilter(ByVal EquityFilter As cEquityFilter)
    Set m.EquityFilter = EquityFilter
End Property

Public Property Get NumRecords() As Long
    NumRecords = m.tblTransactions.NumRecords
End Property
Public Property Let NumRecords(ByVal lNumRecords As Long)
    m.tblTransactions.NumRecords = lNumRecords
End Property

Public Property Get TakeNextTrade() As eGDTakeNextTradeValue
    TakeNextTrade = m.nTakeNextTrade
End Property

Public Function FieldArray(ByVal nField As eGDTransactionField, Optional ByVal bGetCopy As Boolean = True) As cGdArray
    Set FieldArray = m.tblTransactions.FieldArray(nField, bGetCopy)
End Function

Public Sub RemoveRecords(ByVal iFromItem As Long, Optional ByVal nNumToDelete As Long = 1)
    m.tblTransactions.RemoveRecords iFromItem, nNumToDelete
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromTradesArray
'' Description: Split out individual transactions from a trade-by-trade array
'' Inputs:      Trades Array
'' Returns:     None
''
'' Trade line:  Direction, Entry Time, Entry Price, Entry Rule, Exit Time, Exit Price,
''              Exit Rule, Profit, Max Profit, Max Loss, Bars in Trade, Entry Rule ID,
''              Exit Rule ID, TDOW, TDOM, TDOY, -1 if Low before High or 1 otherwise,
''              Quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromTradesArray(ByVal astrTrades As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim astrLine As cGdArray            ' Trade line split out into fields
    Dim lEntryRecord As Long            ' Entry record
    Dim lRec As Long                    ' Index into a for loop
    Dim lLinkRecord As Long             ' Linked record
    Dim lQuantity As Long               ' Quantity for the line
    Dim dTime As Double                 ' Time of the transaction
    Dim dPrice As Double                ' Price of the transaction
    Dim strRule As String               ' Rule for the transaction
    Dim dEntryTime As Double            ' Entry time for the trade

    Set astrLine = New cGdArray
    astrLine.Create eGDARRAY_Strings

    m.tblTransactions.NumRecords = 0&
    
    For lIndex = 1 To astrTrades.Size - 1
        astrLine.SplitFields astrTrades(lIndex), vbTab
        
        lQuantity = CLng(Val(astrLine(17)))
        If lQuantity = 0& Then lQuantity = 1&
        
        Select Case UCase(astrLine(0))
            Case "L"
                lQuantity = lQuantity
            Case "S"
                lQuantity = lQuantity * -1&
            Case Else
                lQuantity = 0&
        End Select
        
        If (lQuantity <> 0&) And (astrLine.Size >= 7) Then
            dTime = RoundToSecond(Val(astrLine(1)))
            dEntryTime = dTime
            dPrice = Val(astrLine(2))
            strRule = Trim(astrLine(3))
            
            ' Verify valid date and rule name exists...
            If (dTime >= 10000) And (Len(strRule) > 0) Then
                ' Handle the Entry Data...
                
                ' Insert before any later trades...
                m.tblTransactions.FieldArray(eGDTransactionField_Date).BinarySearch dTime, lRec
                ' Before we insert the record, we need to make sure that we update
                ' any of the links that will need to change (any records after the
                ' inserted record are now one record later)...
                For lIndex2 = m.tblTransactions.NumRecords - 1 To lRec Step -1
                    lLinkRecord = TblNum(eGDTransactionField_Link, lIndex2)
                    If lLinkRecord >= lRec Then
                        ' TLB 8/3/2009 (#4910): if linked record is after the insertion point,
                        ' then just bump up this link# for this record
                        TblNum(eGDTransactionField_Link, lIndex2) = TblNum(eGDTransactionField_Link, lIndex2) + 1
                    ElseIf lLinkRecord >= 0 Then
                        ' else if linked record is before the insertion point,
                        ' then just bump up the linked record's link# (to where this record will be)
                        TblNum(eGDTransactionField_Link, lLinkRecord) = TblNum(eGDTransactionField_Link, lLinkRecord) + 1
                    End If
                Next lIndex2
                ' now insert the record
                m.tblTransactions.AddRecord "", lRec
                
                If lQuantity > 0& Then
                    TblStr(eGDTransactionField_Action, lRec) = "EL"
                Else
                    TblStr(eGDTransactionField_Action, lRec) = "ES"
                End If
                
                TblNum(eGDTransactionField_Date, lRec) = dTime
                TblNum(eGDTransactionField_Quantity, lRec) = lQuantity
                TblNum(eGDTransactionField_Price, lRec) = dPrice
                TblStr(eGDTransactionField_Rule, lRec) = strRule
                TblNum(eGDTransactionField_Link, lRec) = -1&
                lEntryRecord = lRec
                
                ' Handle the Exit Data...
                
                dTime = RoundToSecond(Val(astrLine(4)))
                ' for sorting purposes: if exit same bar as entered, bump up 1 second, else bump down 1 second
                ' (because if exit date same as entry date then this exit happened after the entries on this bar,
                ' otherwise this exit occurred before any other entries on this bar)
                If dTime = dEntryTime Then
                    dTime = dTime + 0.000001
                Else
                    dTime = dTime - 0.000001
                End If
                dPrice = Val(astrLine(5))
                strRule = Trim(astrLine(6))
                
                ' Verify valid date and rule name exists
                If (dTime >= 10000) And (Len(strRule) > 0) Then
                    ' Insert before any later trades...
                    m.tblTransactions.FieldArray(eGDTransactionField_Date).BinarySearch dTime, lRec
                    ' Before we insert the record, we need to make sure that we update
                    ' any of the links that will need to change (any records after the
                    ' inserted record are now one record later)...
                    For lIndex2 = m.tblTransactions.NumRecords - 1 To lRec Step -1
                        lLinkRecord = TblNum(eGDTransactionField_Link, lIndex2)
                        If lLinkRecord >= lRec Then
                            ' TLB 8/3/2009 (#4910): if linked record is after the insertion point,
                            ' then just bump up this link# for this record
                            TblNum(eGDTransactionField_Link, lIndex2) = TblNum(eGDTransactionField_Link, lIndex2) + 1
                        ElseIf lLinkRecord >= 0 Then
                            ' else if linked record is before the insertion point,
                            ' then just bump up the linked record's link# (to where this record will be)
                            TblNum(eGDTransactionField_Link, lLinkRecord) = TblNum(eGDTransactionField_Link, lLinkRecord) + 1
                        End If
                    Next lIndex2
                    ' now insert the record
                    m.tblTransactions.AddRecord "", lRec

                    ' set fields in table
                    If lQuantity > 0 Then
                        TblStr(eGDTransactionField_Action, lRec) = "XL"
                    Else
                        TblStr(eGDTransactionField_Action, lRec) = "XS"
                    End If
                    
                    TblNum(eGDTransactionField_Date, lRec) = dTime
                    TblNum(eGDTransactionField_Quantity, lRec) = lQuantity * -1&
                    TblNum(eGDTransactionField_Price, lRec) = dPrice
                    TblStr(eGDTransactionField_Rule, lRec) = strRule
                    TblNum(eGDTransactionField_Profit, lRec) = Val(astrLine(7))
                    
                    ' Link entry and exit...
                    If lRec <= lEntryRecord Then
                        ' If we have inserted the exit before the entry for some reason, then
                        ' we need to add one to the entry record because it is now one record
                        ' further than it was before...
                        TblNum(eGDTransactionField_Link, lRec) = lEntryRecord + 1
                    Else
                        TblNum(eGDTransactionField_Link, lRec) = lEntryRecord
                    End If
                    TblNum(eGDTransactionField_Link, lEntryRecord) = lRec
                End If
            End If
        End If
    Next lIndex
   
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransactions.FromTradesArray"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStats
'' Description: Calculate statistics based on the transactions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CalculateStats(Optional ByVal dFromPositionBeforeDate As Double = 0#)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lStart As Long                  ' Starting point for the main for loop
    Dim lQuantity As Long               ' Quantity
    Dim lPosition As Long               ' Position
    Dim dProfit As Double               ' Profit
    Dim dNetProfit As Double            ' Net Profit
    Dim lEntry As Long                  ' Index of the entry
    Dim lEntryQuantity As Long          ' Entry quantity
    Dim lEntryQuantityUsed As Long      ' Quantity of entry used
    Dim dPrice As Double                ' Price
    Dim dSumEntry As Double             ' Sum of the entry
    Dim lExitQuantity As Long           ' Exit quantity
    Dim strAction As String             ' Action for this record

    lStart = 0&

    If m.tblTransactions.NumRecords > 0& Then
        If dFromPositionBeforeDate > 0# Then
            For lIndex = m.tblTransactions.NumRecords - 1 To 0 Step -1
                If TblNum(eGDTransactionField_Position, lIndex) = 0 Then
                    If TblNum(eGDTransactionField_Date, lIndex) < dFromPositionBeforeDate Then
                        lStart = lIndex + 1
                        Exit For
                    End If
                End If
            Next lIndex
        End If
        
        For lIndex = lStart To m.tblTransactions.NumRecords - 1
            TblNum(eGDTransactionField_Date, lIndex) = RoundToSecond(TblNum(eGDTransactionField_Date, lIndex))
            
            lQuantity = TblNum(eGDTransactionField_Quantity, lIndex)
            dProfit = TblNum(eGDTransactionField_Profit, lIndex)
            strAction = TblStr(eGDTransactionField_Action, lIndex)
            
            Select Case UCase(strAction)
                Case "EL", "ES"
                    If lPosition = 0& Then
                        lEntry = lIndex
                        lEntryQuantityUsed = 0&
                    End If
                    dPrice = TblNum(eGDTransactionField_Price, lIndex)
                    dSumEntry = dSumEntry + dPrice * Abs(lQuantity)
                    If dProfit = kNullData Then
                        TblNum(eGDTransactionField_Profit, lIndex) = 0#
                    End If
                
                Case "XL", "XS"
                    If lQuantity = lPosition * -1& Then
                        ' If exiting entire position, clear out entry prices...
                        dSumEntry = 0#
                    Else
                        ' Drop off oldest entry prices...
                        lExitQuantity = Abs(lQuantity)
                        Do While (lEntry < lIndex) And (lExitQuantity > 0&)
                            If Left(strAction, 1) = "E" Then
                                ' Get unexited part of this entry...
                                dPrice = TblNum(eGDTransactionField_Price, lEntry)
                                lEntryQuantity = Abs(TblNum(eGDTransactionField_Quantity, lEntry)) - lEntryQuantityUsed
                                If lExitQuantity < lEntryQuantity Then
                                    ' If only exiting part of this entry...
                                    dSumEntry = dSumEntry - dPrice * lExitQuantity
                                    lEntryQuantityUsed = lEntryQuantityUsed + lExitQuantity
                                    Exit Do
                                End If
                                
                                dSumEntry = dSumEntry - dPrice * lEntryQuantity
                                lEntryQuantityUsed = 0
                                lExitQuantity = lExitQuantity - lEntryQuantity
                            End If
                            lEntry = lEntry + 1&
                        Loop
                    End If
                    If dProfit = kNullData Then
                        TblNum(eGDTransactionField_Profit, lIndex) = 0#
                    End If
            
            End Select
        
            lPosition = lPosition + lQuantity
            
            If lPosition = 0 Then
                TblNum(eGDTransactionField_AvgEntry, lIndex) = 0
            Else
                TblNum(eGDTransactionField_AvgEntry, lIndex) = dSumEntry / Abs(lPosition)
            End If
            
            If dProfit = kNullData Then
                dProfit = 0#
            End If
            
            If lIndex = 0& Then
                TblNum(eGDTransactionField_UnfilteredEquity, lIndex) = dProfit
            Else
                TblNum(eGDTransactionField_UnfilteredEquity, lIndex) = TblNum(eGDTransactionField_UnfilteredEquity, lIndex - 1) + dProfit
            End If
            TblNum(eGDTransactionField_Position, lIndex) = lPosition
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransactions.CalculateStats"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateEquityFilter
'' Description: Calculate the moving average and filtered equity
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CalculateEquityFilter()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim dRollingSum As Double           ' Rolling sum value
    Dim dValue As Double                ' Value to store into the array
    Dim lNum As Long                    ' Number of periods for moving average
    Dim dPct As Double                  ' Percentage for a exponential MA
    Dim lCounter As Long                ' Counter variable
    Dim adEquity As New cGdArray        ' Rolling queue of equities
    Dim dPrevMa As Double               ' Previous moving average value
    Dim lExit As Long                   ' Index for the exit in the table
    Dim dFilteredEquity As Double       ' Filtered equity
    
    lCounter = 0&
    dRollingSum = 0#
    dValue = 0#
    dFilteredEquity = 0#
    lNum = m.EquityFilter.MovingAveragePeriod
    dPct = 2# / (lNum + 1)
    adEquity.Create eGDARRAY_Doubles, lNum + 1
    m.nTakeNextTrade = eGDTakeNextTrade_NotEnoughData
    
    For lIndex = 0 To m.tblTransactions.NumRecords - 1
        Select Case UCase(TblStr(eGDTransactionField_Action, lIndex))
            Case "EL", "ES"
                lExit = TblNum(eGDTransactionField_Link, lIndex)
                
                TblNum(eGDTransactionField_EquityMovAvg, lIndex) = dValue
                TblNum(eGDTransactionField_FilteredEquity, lIndex) = dFilteredEquity
                
                If m.nTakeNextTrade = eGDTakeNextTrade_No Then
                    TblNum(eGDTransactionField_Skip, lIndex) = 1
                    TblNum(eGDTransactionField_Skip, lExit) = 1
                Else
                    TblNum(eGDTransactionField_Skip, lIndex) = 0
                    TblNum(eGDTransactionField_Skip, lExit) = 0
                End If
                
            Case "XL", "XS"
                lCounter = lCounter + 1&
                
                Select Case UCase(m.EquityFilter.MovingAverageType)
                    Case "SIMPLE"
                        dPrevMa = dValue
                        
                        dRollingSum = dRollingSum + TblNum(eGDTransactionField_UnfilteredEquity, lIndex)
                        If lCounter > lNum Then
                            dRollingSum = dRollingSum - adEquity(1)
                            adEquity.Remove 1
                            adEquity(lNum) = TblNum(eGDTransactionField_UnfilteredEquity, lIndex)
                        Else
                            adEquity(lCounter) = TblNum(eGDTransactionField_UnfilteredEquity, lIndex)
                        End If
                        
                        If lCounter >= lNum Then
                            dValue = dRollingSum / CDbl(lNum)
                        Else
                            dValue = -999999#
                        End If
                    
                    Case "EXPONENTIAL"
                        dValue = (TblNum(eGDTransactionField_UnfilteredEquity, lIndex) * dPct) + (dValue * (1 - dPct))
                        
                    Case Else
                        dValue = 0#
                
                End Select
                
                If TblNum(eGDTransactionField_Skip, lIndex) = 0 Then
                    dFilteredEquity = dFilteredEquity + TblNum(eGDTransactionField_Profit, lIndex)
                End If
                
                If m.EquityFilter.EquityFilterOn = False Then
                    m.nTakeNextTrade = eGDTakeNextTrade_NoEquityFilter
                ElseIf dValue = -999999# Then
                    m.nTakeNextTrade = eGDTakeNextTrade_NotEnoughData
                Else
                    Select Case m.EquityFilter.EquityFilterMode
                        Case eGDEquityFilterMode_BelowMa
                            If (TblNum(eGDTransactionField_UnfilteredEquity, lIndex) >= dValue) Then
                                m.nTakeNextTrade = eGDTakeNextTrade_Yes
                            Else
                                m.nTakeNextTrade = eGDTakeNextTrade_No
                            End If
                        
                        Case eGDEquityFilterMode_MaDown
                            If (dValue >= dPrevMa) Then
                                m.nTakeNextTrade = eGDTakeNextTrade_Yes
                            Else
                                m.nTakeNextTrade = eGDTakeNextTrade_No
                            End If
                            
                    End Select
                End If
                
                TblNum(eGDTransactionField_EquityMovAvg, lIndex) = dValue
                TblNum(eGDTransactionField_FilteredEquity, lIndex) = dFilteredEquity
                
        End Select
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransactions.CalculateEquityFilter"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToFile
'' Description: Dump the transactions table to a file
'' Inputs:      Filename
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ToFile(ByVal strFileName As String)
On Error GoTo ErrSection:

    FileFromString strFileName, m.tblTransactions.ToString(vbCrLf, vbTab)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransactions.ToFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.tblTransactions = New cGdTable

    m.tblTransactions.CreateField eGDARRAY_Doubles, eGDTransactionField_Date, "Date"
    m.tblTransactions.CreateField eGDARRAY_Strings, eGDTransactionField_Action, "Action"
    m.tblTransactions.CreateField eGDARRAY_Longs, eGDTransactionField_Quantity, "Quantity"
    m.tblTransactions.CreateField eGDARRAY_Doubles, eGDTransactionField_Price, "Price"
    m.tblTransactions.CreateField eGDARRAY_Strings, eGDTransactionField_Rule, "Rule"
    m.tblTransactions.CreateField eGDARRAY_Doubles, eGDTransactionField_AvgEntry, "AvgEntry"
    m.tblTransactions.CreateField eGDARRAY_Longs, eGDTransactionField_Link, "Link"
    
    m.tblTransactions.CreateField eGDARRAY_Longs, eGDTransactionField_Position, "Position"
    m.tblTransactions.CreateField eGDARRAY_Doubles, eGDTransactionField_Profit, "Profit"
    m.tblTransactions.CreateField eGDARRAY_Doubles, eGDTransactionField_UnfilteredEquity, "UnfilteredEquity"
    m.tblTransactions.CreateField eGDARRAY_Doubles, eGDTransactionField_EquityMovAvg, "EquityMovAvg"
    m.tblTransactions.CreateField eGDARRAY_TinyInts, eGDTransactionField_Skip, "Skip"
    m.tblTransactions.CreateField eGDARRAY_Doubles, eGDTransactionField_FilteredEquity, "FilteredEquity"
    
    m.hTransactions = m.tblTransactions.TableHandle
    
    Set m.EquityFilter = New cEquityFilter
    m.nTakeNextTrade = eGDTakeNextTrade_NoEquityFilter

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransactions.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.tblTransactions = Nothing
    Set m.EquityFilter = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransactions.Class_Terminate"
    
End Sub
