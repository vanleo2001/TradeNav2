VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLateCalculatingFills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cLateCalculatingFills.cls
'' Description: Collection of fills necessary for late calculating object
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 05/16/2011   DAJ         In FromAcctPos, check for LastExit to be Nothing
'' 07/13/2011   DAJ         Calculate stats with every fill in FromFills
'' 10/24/2013   DAJ         Pass account number to g.Profit.Profit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    EntryFills As cGdTree               ' Collection of entry fills
    LastExit As cPtFill                 ' Fill that was the last exit
    
    lPosition As Long                   ' Current position
    dAverageEntry As Double             ' Average entry price
    dClosedProfit As Double             ' Closed profit
End Type
Private m As mPrivate

Private Property Get Entry(ByVal vIndexOrKey As Variant) As cPtFill
On Error GoTo ErrSection:

    Dim ReturnFill As cPtFill           ' Fill to return
    
    Set ReturnFill = Nothing
    If m.EntryFills.Exists(vIndexOrKey) Then
        Set ReturnFill = m.EntryFills(vIndexOrKey)
    End If

    Set Entry = ReturnFill

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cLateCalculatingFills.Entry.Get"
    
End Property

Public Property Get FirstEntry() As cPtFill
    Set FirstEntry = Entry(1)
End Property

Public Property Get LastEntry() As cPtFill
    Set LastEntry = Entry(m.EntryFills.Count)
End Property

Public Property Get LastExit() As cPtFill
    Set LastExit = m.LastExit
End Property

Public Property Get Position() As Long
    Position = m.lPosition
End Property

Public Property Get AverageEntry() As Double
    AverageEntry = m.dAverageEntry
End Property

Public Property Get ClosedProfit() As Double
    ClosedProfit = m.dClosedProfit
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromAcctPos
'' Description: Set class members based on the given account position object
'' Inputs:      Account Position object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromAcctPos(ByVal AcctPos As cAccountPosition)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    m.EntryFills.Clear
    If Not AcctPos Is Nothing Then
        For lIndex = 1 To AcctPos.Entries.Count
            m.EntryFills.Add AcctPos.Entries(lIndex), Str(AcctPos.Entries(lIndex).FillID)
        Next lIndex
        
        If AcctPos.LastExit Is Nothing Then
            m.dClosedProfit = 0#
        Else
            m.dClosedProfit = AcctPos.LastExit.ClosedProfitCategory
        End If
    End If
    
    CalcStats

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculatingFills.FromAcctPos"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromFills
'' Description: Set class members based on the given collection of fills
'' Inputs:      Fills
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromFills(ByVal Fills As cGdTree)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    m.EntryFills.Clear
    If Not Fills Is Nothing Then
        For lIndex = 1 To Fills.Count
            ' 07/13/2011 DAJ: I was passing a False to this function, but that causes
            ' the position not to be calculated and therefore all fills in the collection
            ' are put into the entries array...
            FillCallback Fills(lIndex)
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculatingFills.FromFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillCallback
'' Description: If we get a fill callback, recalc things
'' Inputs:      Fill, Calc Stats?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FillCallback(ByVal Fill As cPtFill, Optional ByVal bCalcStats As Boolean = True)
On Error GoTo ErrSection:

    Dim dClosedProfit As Double         ' Closed profit for the trade
    Dim lEntryIndex As Long             ' Index into the entry fills collection
    Dim EntryFill As cPtFill            ' Fill to close
    Dim ExitFill As cPtFill             ' Exit fill
    Dim nFillMatch As eTT_FillMatchMode ' Fill match mode

    If FillIsEntry(Fill) Then
        m.EntryFills.Add Fill.MakeCopy, Str(Fill.FillID)
        dClosedProfit = 0#
    Else
        Set ExitFill = Fill.MakeCopy
        nFillMatch = g.Broker.Account(ExitFill.AccountID).FillMatchMode
        
        Do
            If nFillMatch = eTT_FillMatchMode_Fifo Then
                lEntryIndex = 1
            Else
                lEntryIndex = m.EntryFills.Count
            End If
            
            Set EntryFill = m.EntryFills(lEntryIndex)
            If EntryFill.Quantity <= ExitFill.Quantity Then
                If ExitFill.Buy Then
                    dClosedProfit = dClosedProfit + g.Profit.Profit(ExitFill.Symbol, EntryFill.Price - ExitFill.Price, EntryFill.Quantity, , , , EntryFill.AccountID)
                Else
                    dClosedProfit = dClosedProfit + g.Profit.Profit(ExitFill.Symbol, ExitFill.Price - EntryFill.Price, EntryFill.Quantity, , , , EntryFill.AccountID)
                End If
                
                m.EntryFills.Remove lEntryIndex
                ExitFill.Quantity = ExitFill.Quantity - EntryFill.Quantity
                
                If (m.EntryFills.Count = 0) And (ExitFill.Quantity > 0) Then
                    m.EntryFills.Add ExitFill, Str(ExitFill.FillID)
                    Exit Do
                End If
            ElseIf EntryFill.Quantity > ExitFill.Quantity Then
                If ExitFill.Buy Then
                    dClosedProfit = dClosedProfit + g.Profit.Profit(ExitFill.Symbol, EntryFill.Price - ExitFill.Price, ExitFill.Quantity, , , , EntryFill.AccountID)
                Else
                    dClosedProfit = dClosedProfit + g.Profit.Profit(ExitFill.Symbol, ExitFill.Price - EntryFill.Price, ExitFill.Quantity, , , , EntryFill.AccountID)
                End If
                
                EntryFill.Quantity = EntryFill.Quantity - ExitFill.Quantity
                Set m.EntryFills(lEntryIndex) = EntryFill
                ExitFill.Quantity = 0
            End If
        Loop While ExitFill.Quantity > 0
        
        Fill.ClosedProfitCategory = dClosedProfit
        m.dClosedProfit = dClosedProfit
        
        Set m.LastExit = Fill.MakeCopy
    End If
    
    If bCalcStats Then
        CalcStats
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculatingFills.FillCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OpenProfit
'' Description: Calculate the open profit based on the entries and the given price
'' Inputs:      Last Price
'' Returns:     Open Profit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenProfit(ByVal dLastPrice As Double) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Entry fill
    Dim dSumPrices As Double            ' Sum of the entry prices
    Dim dSumQuantity As Double          ' Sum of the quantities
    
    dReturn = 0#
    For lIndex = 1 To m.EntryFills.Count
        Set Fill = Entry(lIndex)
        If Not Fill Is Nothing Then
            If Fill.Buy Then
                dReturn = dReturn + g.Profit.Profit(Fill.Symbol, dLastPrice - Fill.Price, Fill.Quantity, , , , Fill.AccountID)
            Else
                dReturn = dReturn + g.Profit.Profit(Fill.Symbol, Fill.Price - dLastPrice, Fill.Quantity, , , , Fill.AccountID)
            End If
        End If
    Next lIndex
    
    OpenProfit = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLateCalculatingFills.OpenProfit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcStats
'' Description: Calculate the statistics
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcStats()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Entry fill
    Dim dSumPrices As Double            ' Sum of the entry prices
    Dim dSumQuantity As Double          ' Sum of the quantities
    
    dSumPrices = 0#
    dSumQuantity = 0#
    
    m.dAverageEntry = 0#
    m.lPosition = 0&
    
    For lIndex = 1 To m.EntryFills.Count
        Set Fill = Entry(lIndex)
        If Not Fill Is Nothing Then
            If Fill.Buy = True Then
                m.lPosition = m.lPosition + Fill.Quantity
            Else
                m.lPosition = m.lPosition - Fill.Quantity
            End If
            
            dSumPrices = dSumPrices + (Fill.Price * Fill.Quantity)
            dSumQuantity = dSumQuantity + Fill.Quantity
        End If
    Next lIndex
    
    If dSumQuantity > 0 Then
        m.dAverageEntry = dSumPrices / dSumQuantity
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculatingFills.CalcStats"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillIsEntry
'' Description: Determine if the given fill is an entry or not
'' Inputs:      Fill
'' Returns:     True if Entry, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FillIsEntry(ByVal Fill As cPtFill) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lPosition As Long               ' Current position
    
    bReturn = False
    lPosition = Position
    
    If lPosition = 0 Then
        bReturn = True
    ElseIf (lPosition > 0) And (Fill.Buy = True) Then
        bReturn = True
    ElseIf (lPosition < 0) And (Fill.Buy = False) Then
        bReturn = True
    End If
    
    FillIsEntry = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLateCalculatingFills.FillIsEntry"
    
End Function

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

    Set m.EntryFills = New cGdTree
    Set m.LastExit = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculatingFills.Class_Initialize"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.EntryFills = Nothing
    Set m.LastExit = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculatingFills.Class_Terminate"
    
End Sub
