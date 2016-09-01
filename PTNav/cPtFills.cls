VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPtFills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPtFills.cls
'' Description: Collection of cPtFill objects
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 02/23/2009   DAJ         Don't move manual fill to history if no earliest session
'' 03/10/2009   DAJ         Removed Xpress manual fill forced to history
'' 03/27/2009   DAJ         Implemented DropOff time for TransAct refreshes
'' 04/30/2009   DAJ         Change snapshot manual fills to history if applicable
'' 07/21/2011   DAJ         Enhanced logging for removing snapshot
'' 08/24/2011   DAJ         Use separate sorted index array
'' 10/07/2011   DAJ         Handle a fill with a previous broker fill ID
'' 10/21/2011   DAJ         Clear collection before load
'' 02/23/2012   DAJ         Fix for when to delete snapshot fills for Rithmic
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 11/01/2013   DAJ         Fix for when to delete snapshot fills for Interactive Brokers
'' 11/22/2013   DAJ         Import historical fills for Turnkey
'' 02/05/2014   DAJ         Fixes for deleting an account
'' 10/24/2014   DAJ         Added LoadFillsForAutoTradeItem; Changed LoadFillsForSymbol
''                          when continuous; Fill Display
'' 12/10/2014   DAJ         Determine snapshot/history line per account
'' 04/14/2015   DAJ         Fix for when to delete snapshot fills for OEC
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Fills As cGdTree                    ' Collection of fills
    astrFillKeys As cGdArray            ' Index for sorting
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cPtFill
Attribute Item.VB_UserMemId = 0
    If m.Fills.Exists(vIndexOrKey) Then
        Set Item = m.Fills(vIndexOrKey)
    Else
        Set Item = Nothing
    End If
End Property
Public Property Let Item(ByVal vIndexOrKey As Variant, Fill As cPtFill)
    Set m.Fills.Item(vIndexOrKey) = Fill
    If m.Fills.Key(vIndexOrKey) <> Fill.BrokerID Then
        m.Fills.Key(vIndexOrKey) = Fill.BrokerID
    End If
End Property

Public Property Get Count() As Long
    Count = m.Fills.Count
End Property

Public Function Remove(ByVal vIndexOrKey As Variant) As Boolean
    Remove = m.Fills.Remove(vIndexOrKey)
End Function

Public Function Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.Fills.Exists(vIndexOrKey)
End Function

Public Sub Clear()
    m.Fills.Clear
    m.astrFillKeys.Clear
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add the given fill to the collection
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(Fill As cPtFill) As Long
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position to insert the fill at
    Dim strFillKey As String            ' Key for the fill
    
    strFillKey = Str(Fill.FillDate) & ";" & Str(Fill.FillID)
    m.astrFillKeys.BinarySearch strFillKey, lPos
    m.astrFillKeys.Add strFillKey, lPos

    Add = m.Fills.Add(Fill, Fill.BrokerID, lPos + 1, eTREE_Myself)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtFills.Add"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Append
'' Description: Append the given fills collection to this one
'' Inputs:      Fills
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Append(ByVal Fills As cPtFills)
On Error GoTo ErrSection:

    Dim lIndex As Long              ' Index into a for loop
    Dim Fill As cPtFill             ' Fill object
    
    If Not Fills Is Nothing Then
        For lIndex = 1 To Fills.Count
            Set Fill = Fills(lIndex)
            
            If Exists(Fill.BrokerID) = False Then
                Add Fill
            End If
        Next lIndex
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFills.Append"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IndexForFill
'' Description: Determine the index for the fill
'' Inputs:      Fill
'' Returns:     Index
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IndexForFill(Fill As cPtFill) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = m.Fills.Index(Fill.BrokerID)
    If (lReturn = 0&) And (Len(Fill.PreviousBrokerID) > 0) Then
        lReturn = m.Fills.Index(Fill.PreviousBrokerID)
    End If
    If lReturn = 0& Then lReturn = -1&
    
    IndexForFill = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtFills.IndexForFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSnapshotFillsForBroker
'' Description: Load the snapshot fills for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadSnapshotFillsForBroker(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As cPtFill                 ' Fill object
    
    Clear
    
    Set rs = g.dbPaper.OpenRecordset("SELECT tblFills.* " & _
                "FROM tblFills INNER JOIN tblAccounts ON tblFills.AccountID=tblAccounts.AccountID " & _
                "WHERE tblAccounts.AccountType=" & Str(nBroker) & " AND tblFills.IsSnapshot=-1 AND tblFills.IsSpread=0;", dbOpenDynaset)
    Do While Not rs.EOF
        Set Fill = New cPtFill
        If Fill.Load(rs!FillID, rs) Then
            Add Fill
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFills.LoadSnapshotFillsForBroker"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadHistoricalFillsForAccount
'' Description: Load the historical fills for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadHistoricalFillsForAccount(ByVal vAccountNumberOrID As Variant)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As cPtFill                 ' Fill object
    Dim lAccountID As Long              ' Account ID
            
    Clear
    
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    If lAccountID > 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT tblFills.* " & _
                    "FROM tblFills INNER JOIN tblAccounts ON tblFills.AccountID=tblAccounts.AccountID " & _
                    "WHERE tblAccounts.AccountID=" & Str(lAccountID) & " AND tblFills.IsSnapshot=0 AND tblFills.IsSpread=0;", dbOpenDynaset)
        Do While Not rs.EOF
            Set Fill = New cPtFill
            If Fill.Load(rs!FillID, rs) Then
                Add Fill
            End If
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFills.LoadHistoricalFillsForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFillsFromRecordset
'' Description: Load the Fills from the given recordset
'' Inputs:      Fills Recordset, Snapshot?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadFillsFromRecordset(rsFills As Recordset, ByVal bSnapshot As Boolean)
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' Fill object

    Clear
    If Not (rsFills.BOF And rsFills.EOF) Then
        rsFills.MoveFirst
        Do While Not rsFills.EOF
            If rsFills!IsSnapshot = bSnapshot Then
                Set Fill = New cPtFill
                If Fill.Load(rsFills!FillID, rsFills) Then
                    Add Fill
                End If
            End If
            
            rsFills.MoveNext
        Loop
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFills.LoadFillsFromRecordset"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFillsForSymbol
'' Description: Load the Fills for the given account/symbol/auto trade item
'' Inputs:      Account, Symbol, Auto Trade ID, Only Snapshot?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadFillsForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long, Optional ByVal bOnlySnapshot As Boolean = False)
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID for account passed in
    Dim lSymbolID As Long               ' Symbol ID for symbol passed in
    Dim strSymbol As String             ' Symbol for symbol passed in
    Dim strSymbolSearch As String       ' Symbol search for the recordset
    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As cPtFill                 ' Fill object
    
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    
    If InStr(strSymbol, "-0") Then
        strSymbolSearch = "[Symbol] LIKE '" & Parse(strSymbol, "-", 1) & "-*'"
    ElseIf lSymbolID = 0& Then
        strSymbolSearch = "[Symbol]='" & strSymbol & "'"
    Else
        strSymbolSearch = "[SymbolID]=" & Str(lSymbolID)
    End If
    
    If lAtID = -1& Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                    "WHERE [AccountID]=" & Str(lAccountID) & " AND " & strSymbolSearch & ";", dbOpenDynaset)
    Else
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                    "WHERE [AccountID]=" & Str(lAccountID) & " AND " & strSymbolSearch & " AND [AutoTradeID]=" & Str(lAtID) & ";", dbOpenDynaset)
    End If
    Do While Not rs.EOF
        Set Fill = New cPtFill
        If Fill.Load(rs!FillID, rs) Then
            If (bOnlySnapshot = False) Or (Fill.IsSnapshot = True) Then
                Add Fill
            End If
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFills.LoadFillsForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFillsForAutoTradeItem
'' Description: Load the Fills for the given auto trade item
'' Inputs:      Auto Trade ID, Only Snapshot?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadFillsForAutoTradeItem(ByVal lAtID As Long, Optional ByVal bOnlySnapshot As Boolean = False)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As cPtFill                 ' Fill object
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] WHERE [AutoTradeID]=" & Str(lAtID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        Set Fill = New cPtFill
        If Fill.Load(rs!FillID, rs) Then
            If (bOnlySnapshot = False) Or (Fill.IsSnapshot = True) Then
                Add Fill
            End If
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFills.LoadFillsForAutoTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save all of the fills to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To Count
        Item(lIndex).Save
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFills.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearSnapshotRefresh
'' Description: Clear the refreshed flag on the snapshot fills
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearSnapshotRefresh()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To Count
        Item(lIndex).Refreshed = False
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFills.ClearSnapshotRefresh"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshSnapshotFill
'' Description: Refresh the given shapshot fill in the collection
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshSnapshotFill(Fill As cPtFill)
On Error GoTo ErrSection:

    Fill.IsSnapshot = True
    Fill.Refreshed = True
    Fill.Save

    If Exists(Fill.BrokerID) Then
        Item(Fill.BrokerID) = Fill.MakeCopy
    Else
        Add Fill.MakeCopy
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFills.RefreshSnapshotFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveSnapshotFills
'' Description: Remove fills or move them out of snapshot as appropriate
'' Inputs:      Earliest Session Date, Fill Symbols
'' Returns:     True if fill moved to history, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemoveSnapshotFills(ByVal EarliestDate As cEarliestDate, ByVal astrFillSymbols As cGdArray) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim dCurrentTime As Double          ' Current time
    Dim dDropOffTime As Double          ' Drop off time
    Dim strBaseSymbol As String         ' Base symbol
    Dim Bars As cGdBars                 ' Bars object
    Dim strFillDisplay As String        ' Fill display
    Dim lEarliestSessionDate As Long    ' Earliest order session date

    bReturn = False
    
    For lIndex = Count To 1 Step -1
        If Item(lIndex).Refreshed = False Then
            lEarliestSessionDate = EarliestDate(Item(lIndex).AccountID)
            
            ' DAJ 04/30/2009: If the refresh was empty and this is a manual fill, then
            ' if the session date for the fill would not have made it a snapshot fill
            ' then move it to history...
            If ((lEarliestSessionDate = Abs(kNullData)) And (Item(lIndex).IsManual = True)) Then
                If g.Broker.DateIsSnapshot(Item(lIndex).SessionDate, Item(lIndex).Broker) = False Then
                    Item(lIndex).IsSnapshot = False
                    Item(lIndex).Save
                    bReturn = True
                    
                    Remove lIndex
                End If
            
            ElseIf (Item(lIndex).SessionDate < lEarliestSessionDate) Or ((lEarliestSessionDate = Abs(kNullData)) And (Item(lIndex).IsManual = False)) Then
                Item(lIndex).IsSnapshot = False
                Item(lIndex).Save
                bReturn = True
                
                Remove lIndex
            
            ElseIf Item(lIndex).IsManual = False Then
                ' DAJ 03/11/2009: If this is a TransAct account, and the fill refresh only had TF fills
                ' in it, and the current item is NOT a TF fill, then move it to history instead of
                ' deleting it.  When TransAct comes up after their daily maintenance at 3:40pm, TF and
                ' only TF fills for the day will remain until 4:00pm mountain time because the TF
                ' trades until 4:00pm...
                If (Item(lIndex).Broker = eTT_AccountType_TransAct) And (Not g.Transact Is Nothing) Then
                    dCurrentTime = CurrentTime("CHI")
                    dDropOffTime = g.Transact.DropOffTime(Item(lIndex).Symbol)
                    strBaseSymbol = Parse(Item(lIndex).Symbol, "-", 1)
                    
                    If ((dCurrentTime > dDropOffTime) And (astrFillSymbols.BinarySearch(strBaseSymbol) = False)) Then
                        Item(lIndex).IsSnapshot = False
                        Item(lIndex).Save True, "After " & DateFormat(dDropOffTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & " Chicago Time"
                        bReturn = True
                    Else
                        strFillDisplay = mTradeTracker.FillDisplay(Item(lIndex))
                        g.Broker.BrokerDebug Item(lIndex).Broker, "Removing Snapshot Fill: '" & strFillDisplay & "'; Session Date = " & DateFormat(Item(lIndex).SessionDate, MM_DD_YYYY) & "; ESD = " & DateFormat(lEarliestSessionDate, MM_DD_YY) & "; Current = " & DateFormat(dCurrentTime, MM_DD_YYYY, HH_MM_SS) & "; Drop Off = " & DateFormat(dDropOffTime, MM_DD_YYYY, HH_MM_SS)
                        Item(lIndex).Delete "Non-Manual Fill no longer exists in refresh"
                    End If
                        
                    Remove lIndex
                    
                ' DAJ 02/23/2012: Found out today that with Rithmic, different symbols can drop out of a
                ' refresh at different times of day.  Not only that, but a symbol can drop out of a
                ' refresh at a different time on different days (e.g. early closing).  So, we will move
                ' the fill into history instead of deleting it if:
                '     a) it is past the closing time for that symbol, or
                '     b) there are no other fills for that symbol in the refresh
                ' DAJ 04/14/2015: Added OEC brokers to this check as well...
                ElseIf g.Broker.IsRithmicBroker(Item(lIndex).Broker) Or g.Broker.IsOecBroker(Item(lIndex).Broker) Then
                    Set Bars = New cGdBars
                    SetBarProperties Bars, Item(lIndex).SymbolOrSymbolID
                    dCurrentTime = CurrentTime(Bars.Prop(eBARS_ExchangeTimeZoneInf))
                    dDropOffTime = CDbl(Item(lIndex).SessionDate) + (Bars.Prop(eBARS_EndTime) / 1440#)
                    strBaseSymbol = Parse(Item(lIndex).Symbol, "-", 1)
                    
                    If dCurrentTime > dDropOffTime Then
                        Item(lIndex).IsSnapshot = False
                        Item(lIndex).Save True, "After " & DateFormat(dDropOffTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
                        bReturn = True
                    ElseIf astrFillSymbols.BinarySearch(strBaseSymbol) = False Then
                        Item(lIndex).IsSnapshot = False
                        Item(lIndex).Save True, "No fills for '" & Item(lIndex).Symbol & "' in refresh"
                        bReturn = True
                    Else
                        strFillDisplay = mTradeTracker.FillDisplay(Item(lIndex))
                        g.Broker.BrokerDebug Item(lIndex).Broker, "Removing Snapshot Fill: '" & strFillDisplay & "'; Session Date = " & DateFormat(Item(lIndex).SessionDate, MM_DD_YYYY) & "; ESD = " & DateFormat(lEarliestSessionDate, MM_DD_YY) & "; Current = " & DateFormat(dCurrentTime, MM_DD_YYYY, HH_MM_SS) & "; Drop Off = " & DateFormat(dDropOffTime, MM_DD_YYYY, HH_MM_SS)
                        Item(lIndex).Delete "Non-Manual Fill no longer exists in refresh"
                    End If
                    
                ' DAJ 11/01/2013: Found out that if you close the Trader Workstation and bring it
                ' back up and reconnect, you are only guaranteed to get the fills that were done this
                ' calendar day ( not session date ).  Because of that, we are just going to move fills
                ' into history that no longer exist in the refresh...
                ElseIf g.Broker.IsIbBroker(Item(lIndex).Broker) Then
                    dCurrentTime = CurrentTime
                    
                    If Int(Item(lIndex).FillDate) < Int(dCurrentTime) Then
                        Item(lIndex).IsSnapshot = False
                        Item(lIndex).Save True, "No longer in refresh"
                        bReturn = True
                    Else
                        Item(lIndex).Delete "Non-Manual Fill no longer exists in refresh"
                        Remove lIndex
                    End If
                
                Else
                    Item(lIndex).Delete "Non-Manual Fill no longer exists in refresh"
                    Remove lIndex
                End If
            End If
        End If
    Next lIndex
    
    RemoveSnapshotFills = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtFills.RemoveSnapshotFills"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveForAccount
'' Description: Remove any fills for the given account
'' Inputs:      Account ID
'' Returns:     Number of items deleted
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemoveForAccount(ByVal lAccountID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object
    
    lReturn = 0&
    For lIndex = m.Fills.Count To 1 Step -1&
        Set Fill = m.Fills(lIndex)
        If Fill.AccountID = lAccountID Then
            m.Fills.Remove lIndex
            lReturn = lReturn + 1&
        End If
    Next lIndex
    
    RemoveForAccount = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtFills.RemoveForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the fills
'' Inputs:      None
'' Returns:     Copy of the fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cPtFills
On Error GoTo ErrSection:

    Dim aCopy As New cPtFills           ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtFills.MakeCopy"

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
    
    Set m.Fills = mCopyFrom.Fills.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFills.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Fills = New cGdTree
    Set m.astrFillKeys = New cGdArray
    m.astrFillKeys.Create eGDARRAY_Strings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFills.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:
    
    Set m.Fills = Nothing
    Set m.astrFillKeys = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFills.Class_Terminate"
    
End Sub
