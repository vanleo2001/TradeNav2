VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccountPositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAccountPositions.cls
'' Description: Collection handler for Account Position objects
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 03/06/2009   DAJ         If an item existed in the database in the create,
''                          but the symbol or symbol ID changed, resave it
'' 10/26/2010   DAJ         Added the last modified property
'' 07/15/2011   DAJ         Return ID from CreateForSymbol call
'' 07/22/2011   DAJ         Don't attempt to create if account doesn't exist
'' 10/21/2011   DAJ         Clear collection before load
'' 02/05/2014   DAJ         Fixes for deleting an account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    AccountPositions As cGdTree         ' Collection of account positions
    dLastModified As Double             ' Last modified date for the collection
End Type
Private m As mPrivate

Public Property Get Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.AccountPositions.Exists(vIndexOrKey)
End Property

Public Property Get Item(ByVal vIndexOrKey As Variant) As cAccountPosition
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    If Exists(vIndexOrKey) Then
        Set Item = m.AccountPositions(vIndexOrKey)
    Else
        Set Item = Nothing
    End If
End Property
Public Property Let Item(ByVal vIndexOrKey As Variant, AcctPos As cAccountPosition)
    If Exists(vIndexOrKey) Then
        Set m.AccountPositions(vIndexOrKey) = AcctPos
        UpdateLastModified
    Else
        Add AcctPos
    End If
End Property

Public Property Get Count()
    Count = m.AccountPositions.Count
End Property

Public Property Get LastModified() As Double
    LastModified = m.dLastModified
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add the given item to the correct place in the collection
'' Inputs:      Account Position
'' Returns:     Index where added
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(AcctPos As cAccountPosition) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position to add the element in
    Dim strKey$, strFind$
    Dim iLo&, iHi&, iMid&, bFound As Boolean
    
    ' TLB 3/22/2013: do a binary search to find where to insert this item
    iLo = 1         ' smallest item#
    iHi = Count     ' largest item#
    iMid = iLo
    bFound = False
    strFind = AcctPos.SortKey
    Do While iLo <= iHi And Not bFound
        iMid = iLo + (iHi - iLo) \ 2  ' so sum does not overflow
        strKey = Item(iMid).SortKey
        If strFind < strKey Then
            iHi = iMid - 1  ' in lower half
        ElseIf strFind > strKey Then
            iLo = iMid + 1  ' in upper half
            iMid = iLo  ' position to insert
        Else
            bFound = True
            ' now back up to the very first match
            Do While iMid > 1 ' bottom
                If Item(iMid - 1).SortKey <> strFind Then Exit Do
                iMid = iMid - 1
            Loop
        End If
    Loop
    lPos = iMid ' position of match or where to insert
    
    Add = m.AccountPositions.Add(AcctPos, AcctPos.Key, lPos, eTREE_Myself)
    
    UpdateLastModified
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPositions.Add"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Key
'' Description: Determine the key into the collection for the given information
'' Inputs:      Account, Symbol, Auto Trade ID
'' Returns:     Key
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Key(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long) As String
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID
    Dim lSymbolID As Long               ' Symbol ID
    Dim strSymbol As String             ' Symbol
    Dim strKey As String                ' Key into the collection
    
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    
    If lSymbolID <> 0& Then
        strKey = Str(lAccountID) & vbTab & Str(lSymbolID) & vbTab & Str(lAtID)
    Else
        strKey = Str(lAccountID) & vbTab & strSymbol & vbTab & Str(lAtID)
    End If
    
    Key = strKey
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPositions.Key"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load all of the account positions from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim AcctPos As cAccountPosition     ' Account Position object
    
    m.AccountPositions.Clear
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions];", dbOpenDynaset)
    Do While Not rs.EOF
        Set AcctPos = New cAccountPosition
        If AcctPos.Load(rs!AccountPositionID, rs) Then
            Add AcctPos
        End If
    
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadForBroker
'' Description: Load all of the account positions for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadForBroker(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim AcctPos As cAccountPosition     ' Account Position object
    
    m.AccountPositions.Clear
    
    Set rs = g.dbPaper.OpenRecordset("SELECT tblAccountPositions.* " & _
                "FROM tblAccountPositions INNER JOIN tblAccounts ON tblAccountPositions.AccountID=tblAccounts.AccountID " & _
                "WHERE tblAccounts.AccountType=" & Str(nBroker) & ";", dbOpenDynaset)

    Do While Not rs.EOF
        Set AcctPos = New cAccountPosition
        
        If AcctPos.Load(rs!AccountPositionID, rs) Then
            Add AcctPos
        End If
        
        rs.MoveNext
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.LoadForBroker"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadForAccount
'' Description: Load all of the account positions for the given account
'' Inputs:      Account, Recordset, Trade Lines Recordset
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadForAccount(ByVal vAccountIDorNumber As Variant, Optional ByVal rs As Recordset = Nothing, Optional ByVal rsTradeLines As Recordset = Nothing)
On Error GoTo ErrSection:

    Dim AcctPos As cAccountPosition     ' Account Position object
    Dim lAccountID As Long              ' Account ID
    
    lAccountID = g.Broker.GetAccountID(vAccountIDorNumber)
    m.AccountPositions.Clear
    
    If rs Is Nothing Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                    "WHERE [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
    End If
    Do While Not rs.EOF
        If rs!AccountID = lAccountID Then
            Set AcctPos = New cAccountPosition
            If AcctPos.Load(rs!AccountPositionID, rs, rsTradeLines) Then
                Add AcctPos
            End If
        End If
    
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.LoadForBroker"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateForSymbol
'' Description: Create an account position for the given information
'' Inputs:      Account, Symbol, Auto Trade ID
'' Returns:     Fill summary ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim strKey As String                ' Key into the collection
    Dim lAccountID As Long              ' Account ID
    Dim lSymbolID As Long               ' Symbol ID
    Dim strSymbol As String             ' Symbol
    Dim rs As Recordset                 ' Recordset into the database
    Dim AcctPos As cAccountPosition     ' Account Position object
    
    lReturn = 0&
    
    strKey = Key(vAccountNumberOrID, vSymbolOrSymbolID, lAtID)
    If Exists(strKey) = True Then
        lReturn = Item(strKey).AccountPositionID
    ElseIf g.Broker.AccountExists(vAccountNumberOrID) Then
        lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
        lSymbolID = GetSymbolID(vSymbolOrSymbolID)
        strSymbol = GetSymbol(vSymbolOrSymbolID)
    
        If lSymbolID <> 0& Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                        "WHERE [AccountID]=" & Str(lAccountID) & " AND [SymbolID]=" & Str(lSymbolID) & " AND [AutoTradeItemID]=" & Str(lAtID) & ";", dbOpenDynaset)
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                        "WHERE [AccountID]=" & Str(lAccountID) & " AND [Symbol]='" & strSymbol & "' AND [AutoTradeItemID]=" & Str(lAtID) & ";", dbOpenDynaset)
        End If
        If rs.BOF And rs.EOF Then
            Set AcctPos = New cAccountPosition
            With AcctPos
                .AccountID = lAccountID
                .SymbolOrSymbolID = vSymbolOrSymbolID
                .AutoTradeItemID = lAtID
                .Save
                
                lReturn = .AccountPositionID
            End With
            
            Add AcctPos
        Else
            ' DAJ 03/05/2009: Found out that if a symbol goes away such that it no longer has a
            ' symbol ID (gets deleted out of symbol universe), that it ends up giving an Error
            ' Code 5 in SetDirtyForSymbol.  Need to resave the record here if the symbol or
            ' symbol ID has changed...
            If (rs!SymbolID <> lSymbolID) Or (rs!Symbol <> strSymbol) Then
                rs.Edit
                rs!SymbolID = lSymbolID
                rs!Symbol = strSymbol
                rs.Update
            End If
            
            Set AcctPos = New cAccountPosition
            If AcctPos.Load(rs!AccountPositionID, rs) Then
                Add AcctPos
                lReturn = AcctPos.AccountPositionID
            End If
        End If
    End If
    
    CreateForSymbol = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPositions.CreateForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateFromDate
'' Description: Recalculate the account position from a certain date/time
'' Inputs:      Account ID, Symbol or SymbolID, Auto Trade Item ID, Date/Time
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateFromDate(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant, ByVal lAutoTradeItemID As Long, ByVal dDateTime As Double)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    
    strKey = Str(lAccountID) & vbTab & Str(vSymbolOrSymbolID) & vbTab & Str(lAutoTradeItemID)
    If Exists(strKey) Then
        Item(strKey).RecalculateFromDate dDateTime
        UpdateLastModified
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.RecalculateFromDate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateFromDateForAccount
'' Description: Recalculate account positions for an account from a certain date/time
'' Inputs:      Account ID, Date/Time
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateFromDateForAccount(ByVal lAccountID As Long, ByVal dDateTime As Double)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.AccountPositions.Count
        If Item(lIndex).AccountID = lAccountID Then
            Item(lIndex).RecalculateFromDate dDateTime
            UpdateLastModified
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.RecalculateFromDate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateHistory
'' Description: Recalculate the history information for the entire collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateHistory()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.AccountPositions.Count
        Item(lIndex).RecalculateHistory
        UpdateLastModified
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.RecalculateHistory"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateHistoryForSymbol
'' Description: Recalculate the history information for the given information
'' Inputs:      Account, Symbol, Auto Trade ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateHistoryForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    
    strKey = Key(vAccountNumberOrID, vSymbolOrSymbolID, lAtID)
    If Exists(strKey) = False Then
        CreateForSymbol vAccountNumberOrID, vSymbolOrSymbolID, lAtID
    End If
    Item(strKey).RecalculateHistory
    UpdateLastModified
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.RecalculateHistoryForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateHistoryForAccount
'' Description: Recalculate the history information for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateHistoryForAccount(ByVal vAccountNumberOrID As Variant)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lAccountID As Long              ' Account ID
    
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    
    For lIndex = 1 To Count
        If Item(lIndex).AccountID = lAccountID Then
            Item(lIndex).RecalculateHistory
            UpdateLastModified
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.RecalculateHistoryForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateSnapshot
'' Description: Recalculate the snapshot information for the entire collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateSnapshot()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.AccountPositions.Count
        Item(lIndex).RecalculateSnapshot
    Next lIndex
    UpdateLastModified
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.RecalculateSnapshot"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateSnapshotForSymbol
'' Description: Recalculate the snapshot information for the given information
'' Inputs:      Account, Symbol, Auto Trade ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateSnapshotForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    
    strKey = Key(vAccountNumberOrID, vSymbolOrSymbolID, lAtID)
    If Exists(strKey) = False Then
        CreateForSymbol vAccountNumberOrID, vSymbolOrSymbolID, lAtID
    End If
    Item(strKey).RecalculateSnapshot
    UpdateLastModified
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.RecalculateSnapshotForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateSnapshotForAccount
'' Description: Recalculate the Snapshot information for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateSnapshotForAccount(ByVal vAccountNumberOrID As Variant)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lAccountID As Long              ' Account ID
    
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    
    For lIndex = 1 To Count
        If Item(lIndex).AccountID = lAccountID Then
            Item(lIndex).RecalculateSnapshot
            UpdateLastModified
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.RecalculateSnapshotForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddSnapshotFillForSymbol
'' Description: Add the snapshot fill for the given information
'' Inputs:      Account, Symbol, Auto Trade ID, Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddSnapshotFillForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long, Fill As cPtFill)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    
    strKey = Key(vAccountNumberOrID, vSymbolOrSymbolID, lAtID)
    If Exists(strKey) = False Then
        CreateForSymbol vAccountNumberOrID, vSymbolOrSymbolID, lAtID
    End If
    Item(strKey).AddSnapshotFill Fill
    UpdateLastModified
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.AddSnapshotFillForSymbol"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetDirtyForSymbol
'' Description: Set the account position dirty for the given information
'' Inputs:      Account, Symbol, Auto Trade ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetDirtyForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    Dim strKey2 As String               ' Key into the collection
    
    strKey = Key(vAccountNumberOrID, vSymbolOrSymbolID, lAtID)
    
    ' If this fill summary doesn't exist yet, create it...
    If Exists(strKey) = False Then
        CreateForSymbol vAccountNumberOrID, vSymbolOrSymbolID, lAtID
        
        ' If we just created a fill summary for the total line, make sure that a
        ' fill summary exists for the manual line as well and set it dirty...
        If lAtID = -1& Then
            strKey2 = Key(vAccountNumberOrID, vSymbolOrSymbolID, 0&)
            If Exists(strKey2) = False Then
                CreateForSymbol vAccountNumberOrID, vSymbolOrSymbolID, 0&
                Item(strKey2).Dirty = True
            End If
        End If
    End If
    
    ' Set the fill summary as being dirty...
    Item(strKey).Dirty = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.SetDirtyForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFill
'' Description: Update the account position for which this fill belongs
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateFill(Fill As cPtFill, Optional ByVal bFillChanged As Boolean = True)
On Error GoTo ErrSection:

    Dim AcctPos As cAccountPosition     ' Account position to create
    
    If Exists(Fill.Key) Then
        Set AcctPos = Item(Fill.Key)
    Else
        Set AcctPos = New cAccountPosition
        AcctPos.AccountID = Fill.AccountID
        AcctPos.AutoTradeItemID = Fill.AutoTradingItemID
        AcctPos.SymbolOrSymbolID = Fill.SymbolOrSymbolID
        AcctPos.Save
        
        Add AcctPos
    End If

    AcctPos.UpdateFill Fill, bFillChanged
    UpdateLastModified

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.UpdateFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFill
'' Description: Remove the fill from the appropriate account position
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveFill(Fill As cPtFill)
On Error GoTo ErrSection:

    Dim AcctPos As cAccountPosition     ' Account position to create
    
    If Exists(Fill.Key) Then
        Set AcctPos = Item(Fill.Key)
        AcctPos.RemoveFill Fill
        UpdateLastModified
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.RemoveFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateFromTradeItem
'' Description: Create a new account position from the given trade item
'' Inputs:      Trade Item
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateFromTradeItem(ByVal TradeItem As cAutoTradeItem)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset from the database
    Dim strKey As String                ' Key for the given Trade Item
    Dim AcctPos As New cAccountPosition ' Account position to create
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountID]=" & Str(TradeItem.AccountID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        strKey = Str(TradeItem.AccountID) & vbTab & Str(TradeItem.SymbolOrSymbolID) & vbTab & Str(TradeItem.AutoTradeItemID)
        If Not Exists(strKey) Then
            With AcctPos
                .AccountID = TradeItem.AccountID
                .AutoTradeItemID = TradeItem.AutoTradeItemID
                .SymbolOrSymbolID = TradeItem.SymbolOrSymbolID
                .Save
            End With
            
            Add AcctPos
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAccountPositions.CreateFromTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshFromTradeItem
'' Description: Refresh the account position for the given trade item
'' Inputs:      Old Trade Item, New Trade Item
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshFromTradeItem(ByVal OldItem As cAutoTradeItem, ByVal NewItem As cAutoTradeItem)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key for the given Trade Item
    Dim AcctPos As New cAccountPosition ' Account position to create
    Dim rs As Recordset                 ' Recordset into the database
    
    strKey = Str(OldItem.AccountID) & vbTab & Str(OldItem.SymbolOrSymbolID) & vbTab & Str(OldItem.AutoTradeItemID)
    If Exists(strKey) Then
        With Item(strKey)
            .AccountID = NewItem.AccountID
            .AutoTradeItemID = NewItem.AutoTradeItemID
            .SymbolOrSymbolID = NewItem.SymbolOrSymbolID
            .Save
        End With
        UpdateLastModified
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.RefreshFromTradeItem"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFromTradeItem
'' Description: Remove the account position for the given trade item
'' Inputs:      Trade Item
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveFromTradeItem(ByVal TradeItem As cAutoTradeItem)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key for the given Trade Item
    Dim AcctPos As New cAccountPosition ' Account position to create
    Dim rs As Recordset                 ' Recordset into the database
    
    strKey = Str(TradeItem.AccountID) & vbTab & Str(TradeItem.SymbolOrSymbolID) & vbTab & Str(TradeItem.AutoTradeItemID)
    If Exists(strKey) Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                "WHERE [AccountPositionID]=" & Str(Item(strKey).AccountPositionID) & ";", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            rs.Delete
        End If
        
        m.AccountPositions.Remove strKey
        UpdateLastModified
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.RemoveFromTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateFromActiveExit
'' Description: Create a new account position from the given exit strategy
'' Inputs:      Exit Strategy
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateFromActiveExit(ByVal ActiveExit As cActiveExit)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset from the database
    Dim strKey As String                ' Key for the given Trade Item
    Dim AcctPos As New cAccountPosition ' Account position to create
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountID]=" & Str(ActiveExit.AccountID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        strKey = Str(ActiveExit.AccountID) & vbTab & Str(ActiveExit.SymbolOrSymbolID) & vbTab & "0"
        If Not Exists(strKey) Then
            With AcctPos
                .AccountID = ActiveExit.AccountID
                .AutoTradeItemID = 0&
                .SymbolOrSymbolID = ActiveExit.SymbolOrSymbolID
                .Save
            End With
            
            Add AcctPos
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAccountPositions.CreateFromActiveExit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveForAccount
'' Description: Remove any account positions for the given account
'' Inputs:      Account ID
'' Returns:     Number of items deleted
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemoveForAccount(ByVal lAccountID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim AcctPos As cAccountPosition     ' Account position object
    
    lReturn = 0&
    For lIndex = m.AccountPositions.Count To 1 Step -1
        Set AcctPos = m.AccountPositions(lIndex)
        If AcctPos.AccountID = lAccountID Then
            m.AccountPositions.Remove lIndex
            lReturn = lReturn + 1&
        End If
    Next lIndex
    
    RemoveForAccount = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPositions.RemoveForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Refresh
'' Description: Refresh the account position with the given account and symbol
'' Inputs:      Account ID, Symbol or Symbol ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Refresh(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key for the given Trade Item

    strKey = Str(lAccountID) & vbTab & Str(vSymbolOrSymbolID) & vbTab & "0"
    If Exists(strKey) Then
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.Refresh"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPosition
'' Description: Determine the current position for the given account and symbol
'' Inputs:      Account ID, Symbol or Symbol ID
'' Returns:     Current Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentPosition(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant) As Long
On Error GoTo ErrSection:

    Dim strKey As String                ' Key for the given Trade Item
    Dim lReturn As Long                 ' Return value from the function

    lReturn = 0&
    strKey = Str(lAccountID) & vbTab & Str(vSymbolOrSymbolID) & vbTab & "0"
    If Exists(strKey) Then
        lReturn = Item(strKey).CurrentPosition
    End If
    
    CurrentPosition = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPositions.CurrentPosition"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateLastModified
'' Description: Update the last modified property
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateLastModified()
On Error GoTo ErrSection:

    m.dLastModified = CurrentTime

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.UpdateLastModified"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the account positions
'' Inputs:      None
'' Returns:     Copy of the account positions
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cAccountPositions
On Error GoTo ErrSection:

    Dim aCopy As New cAccountPositions  ' Copy to return from the function
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPositions.MakeCopy"

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
    
    Set m.AccountPositions = mCopyFrom.AccountPositions.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.SetPrivateData"
    
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
    
    Set m.AccountPositions = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.Class_Initialize"

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

    Set m.AccountPositions = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPositions.Class_Terminate"
    
End Sub
