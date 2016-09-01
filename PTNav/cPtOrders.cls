VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPtOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPtOrders.cls
'' Description: Collection of cPtOrder objects
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 09/01/2009   DAJ         Use new Parked order status
'' 10/07/2009   DAJ         Added support for Linked Orders at broker
'' 07/21/2011   DAJ         Enhanced logging for removing snapshot
'' 07/29/2011   DAJ         Only use previous broker ID to find index if specified
'' 10/21/2011   DAJ         Clear collection before load
'' 10/16/2013   DAJ         Removed PFG
'' 02/05/2014   DAJ         Fixes for deleting an account
'' 04/08/2014   DAJ         Call OrderCallback for conditional orders in RemoveSnapshotOrders
'' 10/09/2014   DAJ         Clear GenesisOrderID for IB orders when move to history
'' 10/29/2014   DAJ         Remove old synthetic order/MIT code
'' 12/10/2014   DAJ         Determine snapshot/history line per account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Orders As cGdTree                   ' Collection of orders
    astrIdMap As cGdArray               ' Map of broker ID's
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cPtOrder
Attribute Item.VB_UserMemId = 0
    If m.Orders.Exists(vIndexOrKey) Then
        Set Item = m.Orders(vIndexOrKey)
    Else
        Set Item = Nothing
    End If
End Property
Public Property Let Item(ByVal vIndexOrKey As Variant, Order As cPtOrder)
    Set m.Orders(vIndexOrKey) = Order
    AddOrderToIdMap Order
End Property

Public Property Get Key(ByVal vIndexOrKey As Variant) As String
    Key = m.Orders.Key(vIndexOrKey)
End Property
Public Property Let Key(ByVal vIndexOrKey As Variant, ByVal strNewKey As String)
    m.Orders.Key(vIndexOrKey) = strNewKey
End Property

Public Property Get Count() As Long
    Count = m.Orders.Count
End Property

Public Function Add(Order As cPtOrder) As Long
    AddOrderToIdMap Order
    Add = m.Orders.Add(Order, Order.BrokerID)
End Function

Public Function Remove(ByVal vIndexOrKey As Variant) As Boolean
    Remove = m.Orders.Remove(vIndexOrKey)
End Function

Public Function Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.Orders.Exists(vIndexOrKey)
End Function

Public Sub Clear()
    m.Orders.Clear
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IndexForOrder
'' Description: Find the index for the order
'' Inputs:      Order, Use the Previous Broker ID?
'' Returns:     Index in the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IndexForOrder(ByVal Order As cPtOrder, Optional ByVal bUsePrevious As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = m.Orders.Index(Order.BrokerID)
    If lReturn = 0& Then
        If (Len(Order.GenesisOrderID) > 0) And ((Len(Order.PreviousBrokerID) = 0) Or (bUsePrevious = False)) Then
            For lIndex = 1 To m.Orders.Count
                If Item(lIndex).GenesisOrderID = Order.GenesisOrderID Then
                    lReturn = lIndex
                    Exit For
                End If
            Next lIndex
        ElseIf (Len(Order.PreviousBrokerID) > 0) And (bUsePrevious = True) Then
            lReturn = m.Orders.Index(Order.PreviousBrokerID)
        End If
    End If
    If lReturn = 0& Then lReturn = -1&
    
    IndexForOrder = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrders.IndexForOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IndexForOrderIDs
'' Description: Find the index for the given order ID(s)
'' Inputs:      Broker ID, Genesis ID, Previous ID
'' Returns:     Index in the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IndexForOrderIDs(ByVal strBrokerID As String, Optional ByVal strGenesisID As String = "", Optional ByVal strPreviousID As String = "") As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = m.Orders.Index(strBrokerID)
    If lReturn = 0& Then
        If Len(strGenesisID) > 0 And Len(strPreviousID) = 0 Then
            For lIndex = 1 To m.Orders.Count
                If Item(lIndex).GenesisOrderID = strGenesisID Then
                    lReturn = lIndex
                    Exit For
                End If
            Next lIndex
        ElseIf Len(strPreviousID) > 0 Then
            lReturn = m.Orders.Index(strPreviousID)
        End If
    End If
    If lReturn = 0& Then lReturn = -1&
    
    IndexForOrderIDs = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrders.IndexForOrderIDs"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSnapshotOrdersForBroker
'' Description: Load the snapshot orders for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadSnapshotOrdersForBroker(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Order As cPtOrder               ' Order object
    
    Clear
    
    ' DAJ 08/27/2009: Also need to include parked orders in the query because if they
    ' were immediately parked and never submitted, the snapshot flag isn't set (Issue #5274)...
    Set rs = g.dbPaper.OpenRecordset("SELECT tblOrders.* " & _
                "FROM tblOrders INNER JOIN tblAccounts ON tblOrders.AccountID=tblAccounts.AccountID " & _
                "WHERE tblAccounts.AccountType=" & Str(nBroker) & " AND (tblOrders.IsSnapshot=-1 OR tblOrders.Status=" & Str(eTT_OrderStatus_Parked) & ");", dbOpenDynaset)
    Do While Not rs.EOF
        Set Order = New cPtOrder
        If Order.Load(rs!OrderID, rs) Then
            Add Order
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrders.LoadSnapshotOrdersForBroker"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrdersFromRecordset
'' Description: Load the orders from the given recordset
'' Inputs:      Orders Recordset, Fills Recordset, Snapshot?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadOrdersFromRecordset(rsOrders As Recordset, rsFills As Recordset, ByVal bSnapshot As Boolean)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object

    If Not (rsOrders.BOF And rsOrders.EOF) Then
        rsOrders.MoveFirst
        Do While Not rsOrders.EOF
            If rsOrders!IsSnapshot = bSnapshot Then
                Set Order = New cPtOrder
                If Order.Load(rsOrders!OrderID, rsOrders, rsFills) Then
                    Add Order
                End If
            End If
            
            rsOrders.MoveNext
        Loop
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrders.LoadOrdersFromRecordset"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save all of the orders to the database
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
    RaiseError "cPtOrders.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearSnapshotRefresh
'' Description: Clear the refreshed flag on the snapshot orders
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
    RaiseError "cPtOrders.ClearSnapshotRefresh"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshSnapshotOrder
'' Description: Refresh the given shapshot order in the collection
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshSnapshotOrder(Order As cPtOrder)
On Error GoTo ErrSection:

    Order.IsSnapshot = True
    Order.Refreshed = True
    Order.Save

    If Exists(Order.BrokerID) Then
        Item(Order.BrokerID) = Order.MakeCopy
    Else
        Add Order.MakeCopy
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrders.RefreshSnapshotOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveSnapshotOrders
'' Description: Remove orders or move them out of snapshot as appropriate
'' Inputs:      Earliest Session Date, Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveSnapshotOrders(ByVal EarliestDate As cEarliestDate, ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bHadPendingOrders As Boolean    ' Were there pending orders that are being removed?
    Dim strBrokerName As String         ' Name for the given broker
    Dim bOrdersMoved As Boolean         ' Have orders been moved?
    Dim lEarliestSessionDate As Long    ' Earliest order session date

    bHadPendingOrders = False
    bOrdersMoved = False
    For lIndex = Count To 1 Step -1
        If (Item(lIndex).Refreshed = False) Then
            If HasBeenSent(Item(lIndex).Status) Then
                lEarliestSessionDate = EarliestDate(Item(lIndex).AccountID)
                
                If bOrdersMoved = False Then
                    bOrdersMoved = True
                    If lEarliestSessionDate = Abs(kNullData) Then
                        g.Broker.BrokerDebug Item(lIndex).Broker, "Remove Snapshot Orders: No Earliest Session Date"
                    Else
                        g.Broker.BrokerDebug Item(lIndex).Broker, "Remove Snapshot Orders: Earliest Session Date = " & DateFormat(lEarliestSessionDate, MM_DD_YYYY)
                    End If
                End If
                
                If OrderIsPending(Item(lIndex)) Then bHadPendingOrders = True
                
                Item(lIndex).IsSnapshot = False
                If (Item(lIndex).SessionDate < lEarliestSessionDate) Then
                    If IsOpenOrder(Item(lIndex).Status, True) Then
                        Item(lIndex).Status = eTT_OrderStatus_Expired
                    End If
                Else
                    If IsOpenOrder(Item(lIndex).Status, True) Then
                        Item(lIndex).Status = eTT_OrderStatus_Cancelled
                    End If
                End If
                
                ' 10/02/2014 DAJ: Customer ran into an issue where his GenesisOrderID's for Interactive
                ' Brokers ( the 'NextId' from the TWS ) reset.  This caused us to start finding old orders
                ' and think he was amending them.  To fix this, I think we can clear the GenesisOrderID
                ' when it moves to history for Interactive Brokers orders...
                If g.Broker.IsIbBroker(Item(lIndex).Broker) = True Then
                    Item(lIndex).GenesisOrderID = ""
                End If
                
                g.Broker.BrokerDebug Item(lIndex).Broker, "Remove Snapshot Order: '" & Item(lIndex).OrderText(True, True, True) & "', Status = '" & OrderStatus(Item(lIndex).Status) & "', SessionDate = " & DateFormat(Item(lIndex).SessionDate, MM_DD_YYYY)
                Item(lIndex).Save
                
                ' Need to do an order callback here to let any automated trading items or auto exits
                ' know that the order status has been changed...
                OrderCallback Item(lIndex)
                
                RemoveOrderFromIdMap Item(lIndex)
                Remove lIndex
            
            ' 04/08/2014 DAJ: Mark Bruels ran into an issue where we stopped checking his conditional
            ' order on a live broker after he restarted Trade Navigator.  I verified that if you shut down
            ' Trade Navigator, we are not adding the order back into the conditional orders collection.  By
            ' doing this here, it should make its way back in...
            ElseIf Item(lIndex).Status = eTT_OrderStatus_TriggerPending Then
                OrderCallback Item(lIndex)
            End If
        End If
    Next lIndex
    
    If bHadPendingOrders Then
        strBrokerName = g.Broker.BrokerName(nBroker)
        InfBox "Trade Navigator found unconfirmed orders that do not exist on the " & strBrokerName & " server.  They have been marked as Cancelled in Trade Navigator.|Please verify orders with your broker.|", "!", "+-OK", strBrokerName & " Order Synchronization", True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrders.RemoveSnapshotOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAltBrokerID
'' Description: Get the alternate broker ID given the broker ID
'' Inputs:      Broker ID
'' Returns:     Alternate Broker ID (Blank if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAltBrokerID(ByVal strBrokerID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lPos As Long                    ' Position in the array
    
    strReturn = ""
    If (m.astrIdMap.Size > 0) And (Len(strBrokerID) > 0) Then
        If m.astrIdMap.BinarySearch(strBrokerID & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
            strReturn = Parse(m.astrIdMap(lPos), vbTab, 2)
        End If
    End If
    
    GetAltBrokerID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrders.GetAltBrokerID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDatabaseIdForBrokerId
'' Description: Get the database ID for the order with the given the broker ID
'' Inputs:      Broker ID
'' Returns:     Database ID (Zero if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDatabaseIdForBrokerId(ByVal strBrokerID As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lPos As Long                    ' Position in the array
    
    lReturn = 0&
    If (m.astrIdMap.Size > 0) And (Len(strBrokerID) > 0) Then
        If m.astrIdMap.BinarySearch(strBrokerID & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
            lReturn = CLng(Val(Parse(m.astrIdMap(lPos), vbTab, 4)))
        End If
    End If
    
    GetDatabaseIdForBrokerId = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrders.GetDatabaseIdForBrokerId"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveForAccount
'' Description: Remove any orders for the given account
'' Inputs:      Account ID
'' Returns:     Number of items deleted
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemoveForAccount(ByVal lAccountID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cPtOrder               ' Order object
    
    lReturn = 0&
    For lIndex = m.Orders.Count To 1 Step -1&
        Set Order = m.Orders(lIndex)
        If Order.AccountID = lAccountID Then
            m.Orders.Remove lIndex
            lReturn = lReturn + 1&
        End If
    Next lIndex
    
    RemoveForAccount = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrders.RemoveForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the orders
'' Inputs:      None
'' Returns:     Copy of the orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cPtOrders
On Error GoTo ErrSection:

    Dim aCopy As New cPtOrders          ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrders.MakeCopy"

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
    
    Set m.Orders = mCopyFrom.Orders.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrders.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOrderToIdMap
'' Description: Add the broker ID of the given order to the ID map if applicable
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddOrderToIdMap(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

'    Dim lPos As Long                    ' Position in the array
'    Dim strBrokerID As String           ' True broker ID
'    Dim strAltBrokerID As String        ' Alternate broker ID
'    Dim strToAdd As String              ' String to add to the array
'
'    If Order.Broker = eTT_AccountType_PFG Then
'        strAltBrokerID = Order.BrokerID
'        If InStr(strAltBrokerID, "|") <> 0 Then
'            strBrokerID = Parse(strAltBrokerID, "|", 2)
'        Else
'            strBrokerID = strAltBrokerID
'        End If
'
'        strToAdd = strBrokerID & vbTab & strAltBrokerID & vbTab & Order.GenesisOrderID & vbTab & Str(Order.OrderID)
'
'        If m.astrIdMap.BinarySearch(strBrokerID & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
'            m.astrIdMap(lPos) = strToAdd
'        Else
'            m.astrIdMap.Add strToAdd, lPos
'        End If
'    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrders.AddOrderToIdMap"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveOrderFromIdMap
'' Description: Remove the broker ID of the given order from the ID map
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveOrderFromIdMap(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

'    Dim lPos As Long                    ' Position in the array
'    Dim strBrokerID As String           ' True broker ID
'    Dim strAltBrokerID As String        ' Alternate broker ID
'
'    If Order.Broker = eTT_AccountType_PFG Then
'        strAltBrokerID = Order.BrokerID
'        If InStr(strAltBrokerID, "|") <> 0 Then
'            strBrokerID = Parse(strAltBrokerID, "|", 2)
'        Else
'            strBrokerID = strAltBrokerID
'        End If
'
'        If m.astrIdMap.BinarySearch(strBrokerID & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
'            m.astrIdMap.Remove lPos
'        End If
'    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrders.RemoveOrderFromIdMap"
    
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

    Set m.Orders = New cGdTree
    
    Set m.astrIdMap = New cGdArray
    m.astrIdMap.Create eGDARRAY_Strings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrders.Class_Initialize"
    
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
    
    Set m.Orders = Nothing
    Set m.astrIdMap = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrders.Class_Terminate"
    
End Sub
