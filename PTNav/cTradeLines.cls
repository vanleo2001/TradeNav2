VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeLines.cls
'' Description: Collection of Trade Line objects
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lAccountPositionID As Long          ' Account Position ID for the collection
    TradeLines As cGdTree               ' Collection of trade line objects
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cTradeLine
Attribute Item.VB_UserMemId = 0
    If m.TradeLines.Exists(vIndexOrKey) Then
        Set Item = m.TradeLines(vIndexOrKey)
    Else
        Set Item = Nothing
    End If
End Property

Public Property Let Item(ByVal vIndexOrKey As Variant, TradeLine As cTradeLine)
    Set m.TradeLines(vIndexOrKey) = TradeLine
End Property

Public Property Get Count() As Long
    Count = m.TradeLines.Count
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add the given trade line to the collection
'' Inputs:      Trade Line
'' Returns:     Index of the Item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(TradeLine As cTradeLine) As Long
On Error GoTo ErrSection:

    Add = m.TradeLines.Add(TradeLine, Str(TradeLine.TradeNumber))
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeLines.Add"
    
End Function

Public Function Remove(ByVal vIndexOrKey As Variant) As Boolean
    Remove = m.TradeLines.Remove(vIndexOrKey)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load up all of the trade lines for the given account position ID
'' Inputs:      Account Position ID, Symbol, Recordset
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load(ByVal lAccountPositionID As Long, ByVal vSymbolOrSymbolID As Variant, Optional ByVal rs As Recordset = Nothing)
On Error GoTo ErrSection:

    Dim TradeLine As cTradeLine         ' Trade line object
    Dim Bars As cGdBars                 ' Bars object
    Dim lAccountID As Long              ' Account ID
    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim rs2 As Recordset                ' Recordset into the database
    
    m.lAccountPositionID = lAccountPositionID
    
    ' If the recordset was not passed in, we need to load it from the database...
    If rs Is Nothing Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionTrades] " & _
                "WHERE [AccountPositionID]=" & Str(lAccountPositionID) & ";", dbOpenDynaset)
    End If

    ' Get the account ID from the account positions record...
    Set rs2 = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                "WHERE [AccountPositionID]=" & Str(m.lAccountPositionID) & ";", dbOpenDynaset)
    If Not (rs2.BOF And rs2.EOF) Then
        lAccountID = rs2!AccountID
        nBroker = g.Broker.AccountTypeForID(lAccountID)
    End If

    ' Set bar properties to pass in to each of the tradeline objects...
    Set Bars = New cGdBars
    SetBarProperties Bars, vSymbolOrSymbolID

    Do While Not rs.EOF
        If rs!AccountPositionID = lAccountPositionID Then
            Set TradeLine = New cTradeLine
            If TradeLine.Load(rs!AccountPositionID, rs!TradeNumber, vSymbolOrSymbolID, rs, lAccountID, nBroker, Bars) Then
                Add TradeLine
            End If
        End If
        
        rs.MoveNext
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLines.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadHistorical
'' Description: Load up all of the closed historical trade lines for the given
''              account position ID
'' Inputs:      Account Position ID, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadHistorical(ByVal lAccountPositionID As Long, ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim TradeLine As cTradeLine         ' Trade line object
    
    m.lAccountPositionID = lAccountPositionID
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionTrades] " & _
            "WHERE [AccountPositionID]=" & Str(lAccountPositionID) & " AND [IsSnapshot]=0;", dbOpenDynaset)
    Do While Not rs.EOF
        Set TradeLine = New cTradeLine
        If TradeLine.Load(rs!AccountPositionID, rs!TradeNumber, vSymbolOrSymbolID, rs) Then
            Add TradeLine
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLines.LoadHistorical"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSnapshot
'' Description: Load up all of the snapshot trade lines for the given
''              account position ID
'' Inputs:      Account Position ID, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadSnapshot(ByVal lAccountPositionID As Long, ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim TradeLine As cTradeLine         ' Trade line object
    
    m.lAccountPositionID = lAccountPositionID
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionTrades] " & _
            "WHERE [AccountPositionID]=" & Str(lAccountPositionID) & " AND [IsSnapshot]=-1;", dbOpenDynaset)
    Do While Not rs.EOF
        Set TradeLine = New cTradeLine
        If TradeLine.Load(rs!AccountPositionID, rs!TradeNumber, vSymbolOrSymbolID, rs) Then
            Add TradeLine
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLines.LoadHistorical"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the tradelines to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim rs As Recordset                 ' Recordset into the database
    
    ' First, delete all of the trade lines from the database in case
    ' some no longer exist...
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionTrades] " & _
            "WHERE [AccountPositionID]=" & Str(m.lAccountPositionID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
    
    ' Next, save each of the trade lines...
    For lIndex = 1 To m.TradeLines.Count
        Item(lIndex).Save
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLines.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Clear
'' Description: Clear out the collection
'' Inputs:      Delete from Database?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear(Optional ByVal bDeleteFromDatabase As Boolean = False)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    If bDeleteFromDatabase Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionTrades] " & _
                "WHERE [AccountPositionID]=" & Str(m.lAccountPositionID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            rs.Delete
            rs.MoveNext
        Loop
    End If
    m.TradeLines.Clear

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLines.Clear"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearSnapshot
'' Description: Clear out the snapshot trades from the collection
'' Inputs:      Delete from Database?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearSnapshot(Optional ByVal bDeleteFromDatabase As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim rs As Recordset                 ' Recordset into the database

    If bDeleteFromDatabase Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionTrades] " & _
                "WHERE [AccountPositionID]=" & Str(m.lAccountPositionID) & " AND [IsSnapshot]=-1;", dbOpenDynaset)
        Do While Not rs.EOF
            rs.Delete
            rs.MoveNext
        Loop
    End If
    
    For lIndex = Count To 1 Step -1
        If Item(lIndex).IsSnapshot Then
            m.TradeLines.Remove lIndex
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLines.ClearSnapshot"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Combine
'' Description: Combine trade lines that are the same
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Combine()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lStart As Long                  ' Starting spot
    Dim lEnd As Long                    ' Ending spot

    For lIndex = Count To 2 Step -1
        lEnd = lIndex
        
        lStart = 1
        For lIndex2 = lEnd - 1 To 1 Step -1
            If Item(lEnd).IsTheSame(Item(lIndex2)) = False Then
                lStart = lIndex2 + 1
                Exit For
            End If
        Next lIndex2
        
        Item(lStart).Quantity = lEnd - lStart + 1
        Item(lStart).ClosedProfit = Item(lStart).ClosedProfit * Item(lStart).Quantity
        
        For lIndex2 = lEnd To lStart + 1 Step -1
            m.TradeLines.Remove lIndex2
        Next lIndex2
        
        lIndex = lStart - 1
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLines.Combine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToArray
'' Description: Convert the collection to a trade-by-trade array
'' Inputs:      None
'' Returns:     Trade-by-trade array
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToArray() As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array to return from the function
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    For lIndex = 1 To m.TradeLines.Count
        astrReturn.Add Item(lIndex).ToString
    Next lIndex
    
    Set ToArray = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeLines.ToArray"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EarliestFillTime
'' Description: Determine the earliest fill time in the collection
'' Inputs:      None
'' Returns:     Earliest Fill Time
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EarliestFillTime() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value from the function
    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.TradeLines.Count
        If (Item(lIndex).EntryTime < dReturn) Or (dReturn = 0#) Then
            dReturn = Item(lIndex).EntryTime
        End If
        If (Item(lIndex).ExitTime < dReturn) Or (dReturn = 0#) Then
            dReturn = Item(lIndex).ExitTime
        End If
    Next lIndex
    
    EarliestFillTime = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeLines.EarliestFillTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LatestFillTime
'' Description: Determine the latest fill time in the collection
'' Inputs:      None
'' Returns:     Latest Fill Time
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LatestFillTime() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value from the function
    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.TradeLines.Count
        If (Item(lIndex).EntryTime > dReturn) Or (dReturn = 0#) Then
            dReturn = Item(lIndex).EntryTime
        End If
        If (Item(lIndex).ExitTime > dReturn) Or (dReturn = 0#) Then
            dReturn = Item(lIndex).ExitTime
        End If
    Next lIndex
    
    LatestFillTime = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeLines.LatestFillTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the orders
'' Inputs:      None
'' Returns:     Copy of the orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cTradeLines
On Error GoTo ErrSection:

    Dim aCopy As New cTradeLines        ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeLines.MakeCopy"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private data for copying purposes
'' Inputs:      Private members to copy from
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    m = mCopyFrom
    
    Set m.TradeLines = mCopyFrom.TradeLines.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLines.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when object gets initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.TradeLines = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLines.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when object gets terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:
    
    Set m.TradeLines = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLines.Class_Terminate"
    
End Sub
