VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOrderLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cOrderLinks.cls
'' Description: Class to manange a collection of order links
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 12/01/2009   DAJ         Enhancements for Auto Exits held at broker
'' 12/09/2009   DAJ         Fixes for broker held OCOs
'' 12/16/2009   DAJ         Return value for unlink and park functions
'' 12/21/2009   DAJ         Return value for unlink and submit functions
'' 07/15/2011   DAJ         Handle order link failure from PFG servers
'' 08/02/2011   DAJ         Wait for orders to both go working before linking
'' 09/19/2011   DAJ         Turned logging on
'' 09/29/2011   DAJ         When checking links, make sure orders are not nothing
'' 07/13/2012   DAJ         In DumpDebug, ensure that PFG is not nothing
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    OrderLinks As cGdTree               ' Collection of order links
    WaitList As cOrderLinkWaitList      ' List of things waiting for unlink
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cOrderLink
On Error GoTo ErrSection:

    Dim OrderLink As cOrderLink         ' Order link to return from the function

    Set OrderLink = Nothing
    If Exists(vIndexOrKey) Then
        Set OrderLink = m.OrderLinks(vIndexOrKey)
    End If
    
    Set Item = OrderLink
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cOrderLinks.Item.Get"
    
End Property

Public Function Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.OrderLinks.Exists(vIndexOrKey)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add an order link object to the collection
'' Inputs:      Order Link
'' Returns:     Position in the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(OrderLink As cOrderLink) As Long
On Error GoTo ErrSection:

    Add = m.OrderLinks.Add(OrderLink, Str(OrderLink.OrderLinkID))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinks.Add"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load working order links from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim OrderLink As cOrderLink         ' Order link object
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderLinks];", dbOpenDynaset)
    Do While Not rs.EOF
        Set OrderLink = New cOrderLink
        If OrderLink.Load(rs!OrderLinkID, rs) Then
            DumpDebug "OrderLinks: OrderLink Loaded: " & OrderLink.ToString
            Add OrderLink
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinks.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshOrder
'' Description: Check to see if new link exists, link changed, or link removed
'' Inputs:      Order, Was the order parked?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshOrder(Order As cPtOrder, Optional ByVal bWasParked As Boolean = False)
On Error GoTo ErrSection:

    Dim OrderLink As cOrderLink         ' Order Link object
    Dim OtherOrder As cPtOrder          ' Other order in the link
    Dim bDoCheck As Boolean             ' Do we want to do the check?
    
    Set OrderLink = OrderLinkForOrder(Order)
    
    bDoCheck = False
    If OrderLink Is Nothing Then
        ' A new broker order link exists...
        If Order.BrokerCancelOrderID <> 0& Then
            DumpDebug "OrderLinks: Refresh Order '" & Order.OrderText(True, True, True) & "': Status = '" & OrderStatus(Order.Status) & "', Broker OCO: " & Str(Order.BrokerCancelOrderID)
            
            Set OtherOrder = New cPtOrder
            OtherOrder.Load Abs(Order.BrokerCancelOrderID)
            
            Set OrderLink = OrderLinkForOrder(OtherOrder)
            
            ' Neither order is currently in an order link...
            If OrderLink Is Nothing Then
                If IsOpenOrder(Order.Status) And IsOpenOrder(OtherOrder.Status) Then
                    Set OrderLink = New cOrderLink
                    
                    OrderLink.DatabaseId1 = Order.OrderID
                    OrderLink.BrokerId1 = Order.BrokerID
                    OrderLink.Order1 = Order.MakeCopy
                    OrderLink.DatabaseId2 = Abs(Order.BrokerCancelOrderID)
                    OrderLink.BrokerId2 = OtherOrder.BrokerID
                    OrderLink.Order2 = OtherOrder.MakeCopy
                    
                    If Order.BrokerCancelOrderID < 0 Then
                        OrderLink.Status = eGDOrderLinkStatus_New
                    Else
                        OrderLink.Status = eGDOrderLinkStatus_Confirmed
                    End If
                    
                    OrderLink.Save
                    
                    DumpDebug "     Order Link Created: " & OrderLink.ToString
                    Add OrderLink
                    
                    bDoCheck = True
                End If
                
            ' The "other order" is currently in an order link, so we need to change the link...
            Else
                DumpDebug "     " & Str(OtherOrder.OrderID) & " has changed its BrokerCancelOrderID to " & Str(OtherOrder.BrokerCancelOrderID)
                If OtherOrder.BrokerCancelOrderID > 0 Then
                    OrderLink.DatabaseId1 = OtherOrder.OrderID
                    OrderLink.BrokerId1 = OtherOrder.BrokerID
                    OrderLink.Order1 = OtherOrder.MakeCopy
                    OrderLink.DatabaseId2 = OtherOrder.BrokerCancelOrderID
                    OrderLink.BrokerId2 = Order.BrokerID
                    OrderLink.Order2 = Order.MakeCopy
                    OrderLink.Status = eGDOrderLinkStatus_Confirmed
                    OrderLink.Save
                
                    DumpDebug "     Order Link Changed: " & OrderLink.ToString
                    Set m.OrderLinks(Str(OrderLink.OrderLinkID)) = OrderLink
                Else
                    If AddToWaitList(OrderLink, OtherOrder, eGDWaitListCommand_Link) Then
                        DumpDebug "     Order Link Unlinking (Different ID): " & OrderLink.ToString
                        SendUnlinkRequest OrderLink
                    End If
                End If
            End If
        End If
    Else
        DumpDebug "OrderLinks: Refresh Order '" & Order.OrderText(True, True, True) & "': Status = '" & OrderStatus(Order.Status) & "', Broker OCO: " & Str(Order.BrokerCancelOrderID)
        
        ' An existing link no longer exists...
        If (Order.BrokerCancelOrderID = 0&) And (OrderLink.Status = eGDOrderLinkStatus_Confirmed) Then
            DumpDebug "     Order Link Unlinking: " & OrderLink.ToString
            SendUnlinkRequest OrderLink
            
        ' The order has been closed (but not Amended and the Order Link is not parked or the order didn't
        ' just go from Parked to Cancelled)...
        ElseIf (IsOpenOrder(Order.Status) = False) And (Order.Status <> eTT_OrderStatus_Amended) And ((OrderLink.Status <> eGDOrderLinkStatus_Parked) Or (bWasParked = False)) Then
            DumpDebug "     Removing link because closed: " & OrderLink.ToString
            RemoveLink OrderLink
        
        ' The given order is now linked to a different order...
        ElseIf (Order.BrokerCancelOrderID <> 0) And (Abs(Order.BrokerCancelOrderID) <> OrderLink.OtherDatabaseID(Order.OrderID)) Then
            DumpDebug "     " & Str(Order.OrderID) & " has changed its BrokerCancelOrderID to " & Str(Order.BrokerCancelOrderID)
            If Order.BrokerCancelOrderID > 0 Then
                Set OtherOrder = New cPtOrder
                OtherOrder.Load Abs(Order.BrokerCancelOrderID)
                
                OrderLink.DatabaseId1 = Order.OrderID
                OrderLink.BrokerId1 = Order.BrokerID
                OrderLink.Order1 = Order.MakeCopy
                OrderLink.DatabaseId2 = Order.BrokerCancelOrderID
                OrderLink.BrokerId2 = OtherOrder.BrokerID
                OrderLink.Order2 = OtherOrder.MakeCopy
                OrderLink.Status = eGDOrderLinkStatus_Confirmed
                OrderLink.Save
            
                DumpDebug "     Order Link Changed: " & OrderLink.ToString
                Set m.OrderLinks(Str(OrderLink.OrderLinkID)) = OrderLink
            Else
                If OrderLink.Status = eGDOrderLinkStatus_Parked Then
                    OrderIDChanged OrderLink.OtherDatabaseID(Order.OrderID), Order.BrokerCancelOrderID
                Else
                    If AddToWaitList(OrderLink, Order, eGDWaitListCommand_Link) Then
                        DumpDebug "     Order Link Unlinking (Different ID): " & OrderLink.ToString
                        SendUnlinkRequest OrderLink
                    End If
                End If
            End If
            
        ' If both orders are parked, we can call the link parked...
        ElseIf (Order.Status = eTT_OrderStatus_Parked) And (OrderLink.Status <> eGDOrderLinkStatus_Parked) Then
            Set OtherOrder = New cPtOrder
            If OtherOrder.Load(Abs(OtherDatabaseID(OrderLink, Order.OrderID))) Then
                If OtherOrder.Status = eTT_OrderStatus_Parked Then
                    If OrderLink.DatabaseId1 = Order.OrderID Then
                        OrderLink.Order1 = Order.MakeCopy
                        OrderLink.Order2 = OtherOrder.MakeCopy
                    Else
                        OrderLink.Order1 = OtherOrder.MakeCopy
                        OrderLink.Order2 = Order.MakeCopy
                    End If
                    
                    OrderLink.BrokerId1 = ""
                    OrderLink.BrokerId2 = ""
                    OrderLink.Status = eGDOrderLinkStatus_Parked
                    OrderLink.Save
                    Set m.OrderLinks(Str(OrderLink.OrderLinkID)) = OrderLink
                    DumpDebug "     Both orders are Parked, so marking the link as parked: " & OrderLink.ToString
                    
                    Order.BrokerCancelOrderID = OtherOrder.OrderID * -1
                    Order.Save
                    g.Broker.AddOrder Order
                    OrderCallback Order
                End If
            End If
        
        ' Update the Broker ID if necessary...
        Else
            If Order.OrderID = OrderLink.DatabaseId1 Then
                OrderLink.Order1 = Order.MakeCopy
                
                If Order.BrokerID <> OrderLink.BrokerId1 Then
                    OrderLink.BrokerId1 = Order.BrokerID
                    OrderLink.Save
                    Set m.OrderLinks(Str(OrderLink.OrderLinkID)) = OrderLink
                    
                    DumpDebug "     OrderLink.BrokerID1 Set to " & Order.BrokerID
                End If
            ElseIf Order.OrderID = OrderLink.DatabaseId2 Then
                OrderLink.Order2 = Order.MakeCopy
                
                If Order.BrokerID <> OrderLink.BrokerId2 Then
                    OrderLink.BrokerId2 = Order.BrokerID
                    OrderLink.Save
                    Set m.OrderLinks(Str(OrderLink.OrderLinkID)) = OrderLink
                    
                    DumpDebug "     OrderLink.BrokerID2 Set to " & Order.BrokerID
                End If
            End If
        End If
        
        bDoCheck = True
    End If
    
    If bDoCheck = True Then
        CheckOrderLinks
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cOrderLinks.RefreshOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIdChanged
'' Description: Handle an Order ID change
'' Inputs:      Original Order ID, New Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderIDChanged(ByVal lOriginalOrderID As Long, ByVal lNewOrderID As Long)
On Error GoTo ErrSection:

    Dim OrderLink As cOrderLink         ' Order Link object
    Dim NewOrder As New cPtOrder        ' Order object
    Dim strOldOrderLink As String       ' Old order link information

    Set OrderLink = OrderLinkForDatabaseID(lOriginalOrderID)
    If Not OrderLink Is Nothing Then
        DumpDebug "OrderLinks: Order ID Changed from " & Str(lOriginalOrderID) & " to " & Str(lNewOrderID)
            
#If 0 Then
        If NewOrder.Load(lNewOrderID) Then
            strOldOrderLink = OrderLink.ToString
            
            If OrderLink.DatabaseId1 = lOriginalOrderID Then
                OrderLink.DatabaseId1 = lNewOrderID
                OrderLink.BrokerId1 = NewOrder.BrokerID
            ElseIf OrderLink.DatabaseId2 = lOriginalOrderID Then
                OrderLink.DatabaseId2 = lNewOrderID
                OrderLink.BrokerId2 = NewOrder.BrokerID
            End If
            OrderLink.Status = eGDOrderLinkStatus_New
            
            ' Remove the wait list item...
            m.WaitList.Remove Str(OrderLink.OrderLinkID)
            
            DumpDebug "     Order Link Changed from '" & strOldOrderLink & "' to '" & OrderLink.ToString & "'"
            OrderLink.Save
            
            CheckOrderLinks
        End If
#End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinks.OrderIdChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckOrderLinks
'' Description: Check all of the order links to see if an action is necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckOrderLinks()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim OrderLink As cOrderLink         ' Order Link object
    
    For lIndex = m.OrderLinks.Count To 1 Step -1
        Set OrderLink = Item(lIndex)
        If Not OrderLink Is Nothing Then
            If OrderLink.Status = eGDOrderLinkStatus_New Then
                If (Not (OrderLink.Order1 Is Nothing)) And (Not (OrderLink.Order2 Is Nothing)) Then
                    If (OrderLink.Order1.Status = eTT_OrderStatus_Working) And (OrderLink.Order2.Status = eTT_OrderStatus_Working) Then
                        DumpDebug "OrderLinks: Sending Link Request: " & OrderLink.ToString
                        SendLinkRequest OrderLink
                    End If
                End If
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinks.CheckOrderLinks"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LinkConfirm
'' Description: Callback to state that a link has been confirmed
'' Inputs:      Broker ID1, Broker ID2, Failed?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LinkConfirm(ByVal strBrokerId1 As String, ByVal strBrokerId2, Optional ByVal bFailed As Boolean = False)
On Error GoTo ErrSection:

    Dim OrderLink As cOrderLink         ' Order link object
    Dim WaitListItem As cOrderLinkWaitListItem
    
    Set OrderLink = OrderLinkForBrokerID(strBrokerId1)
    If Not OrderLink Is Nothing Then
        If OrderLink.OtherBrokerID(strBrokerId1) = strBrokerId2 Then
            If bFailed Then
                RemoveLink OrderLink
            Else
                OrderLink.Status = eGDOrderLinkStatus_Confirmed
                OrderLink.Save
                
                Set m.OrderLinks(Str(OrderLink.OrderLinkID)) = OrderLink
                DumpDebug "OrderLinks: Link Confirmed: " & OrderLink.ToString
                
                Set WaitListItem = m.WaitList.Item(Str(OrderLink.OrderLinkID))
                If Not WaitListItem Is Nothing Then
                    If WaitListItem.Command = eGDWaitListCommand_Link Then
                        If m.WaitList.Remove(Str(OrderLink.OrderLinkID)) Then
                            DumpDebug "OrderLinks: Wait List Item Removed: '" & WaitListItem.ToString & "'"
                        End If
                    End If
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cOrderLinks.LinkConfirm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnlinkConfirm
'' Description: Callback to state that an unlink has been confirmed
'' Inputs:      Broker ID1, Broker ID2
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UnlinkConfirm(ByVal strBrokerId1 As String, ByVal strBrokerId2 As String)
On Error GoTo ErrSection:

    Dim OrderLink As cOrderLink         ' Order link object
    Dim bRemoveLink As Boolean          ' Remove the link object?
    Dim WaitListItem As cOrderLinkWaitListItem
    
    DumpDebug "OrderLinks: Unlink Confirmed(" & strBrokerId1 & ", " & strBrokerId2 & ")"
    Set OrderLink = OrderLinkForBrokerID(strBrokerId1)
    If Not OrderLink Is Nothing Then
        If OrderLink.OtherBrokerID(strBrokerId1) = strBrokerId2 Then
            DumpDebug "     Unlink Confirmed: " & OrderLink.ToString
            
            bRemoveLink = True
            Set WaitListItem = m.WaitList.Item(Str(OrderLink.OrderLinkID))
            If Not WaitListItem Is Nothing Then
                Select Case WaitListItem.Command
                    Case eGDWaitListCommand_AmendOne, eGDWaitListCommand_AmendBoth, eGDWaitListCommand_ParkBoth
                        bRemoveLink = False
                End Select
            End If
            
            If bRemoveLink = True Then
                RemoveLink OrderLink
            End If
            
            PerformCommandForWaitList OrderLink.OrderLinkID
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cOrderLinks.UnlinkConfirm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnlinkAndCancelOrder
'' Description: Unlink two orders and then cancel the one passed in
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UnlinkAndCancelOrder(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim OrderLink As cOrderLink         ' Order Link object
    
    DumpDebug "OrderLinks: UnlinkAndCancel: '" & Order.OrderText & "'"
    Set OrderLink = OrderLinkForOrder(Order)
    DumpDebug "     OrderLink: '" & OrderLink.ToString & "'"
    
    If Not OrderLink Is Nothing Then
        If AddToWaitList(OrderLink, Order, eGDWaitListCommand_Cancel) Then
            DumpDebug "     Added to WaitList: '" & OrderLink.ToString & "', '" & Order.OrderText & "'"
            DumpDebug "     Unlinking Orders (Cancel): '" & OrderLink.ToString & "'"
            SendUnlinkRequest OrderLink
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinks.UnlinkAndCancelOrder"
    
End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnlinkAmendOneAndRelink
'' Description: Unlink two orders, amend the one, and relink
'' Inputs:      Original Order, New Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UnlinkAmendOneAndRelink(ByVal OriginalOrder As cPtOrder, ByVal NewOrder As cPtOrder)
On Error GoTo ErrSection:

    Dim OrderLink As cOrderLink         ' Order Link object
    
    DumpDebug "OrderLinks: UnlinkAmendOneAndRelink: '" & OriginalOrder.OrderText & "', '" & NewOrder.OrderText & "'"
    Set OrderLink = OrderLinkForOrder(OriginalOrder)
    DumpDebug "     OrderLink: '" & OrderLink.ToString & "'"
    
    If Not OrderLink Is Nothing Then
        If AddToWaitList(OrderLink, OriginalOrder, eGDWaitListCommand_AmendOne, NewOrder) Then
            DumpDebug "     Added to WaitList: '" & OrderLink.ToString & "', '" & OriginalOrder.OrderText & "'"
            DumpDebug "     Unlinking Orders (Amend): '" & OrderLink.ToString & "'"
            SendUnlinkRequest OrderLink
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinks.UnlinkAmendOneAndRelink"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnlinkAmendBothAndRelink
'' Description: Unlink two orders, amend both, and relink
'' Inputs:      Original Order1, New Order1, Original Order2, New Order2
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UnlinkAmendBothAndRelink(ByVal OriginalOrder1 As cPtOrder, ByVal NewOrder1 As cPtOrder, ByVal OriginalOrder2 As cPtOrder, ByVal NewOrder2 As cPtOrder)
On Error GoTo ErrSection:

    Dim OrderLink As cOrderLink         ' Order Link object
    
    DumpDebug "OrderLinks: UnlinkAmendBothAndRelink: '" & OriginalOrder1.OrderText & "', '" & NewOrder1.OrderText & "', '" & OriginalOrder2.OrderText & "', '" & NewOrder2.OrderText & "'"
    Set OrderLink = OrderLinkForOrder(OriginalOrder1)
    DumpDebug "     OrderLink: '" & OrderLink.ToString & "'"
    
    If Not OrderLink Is Nothing Then
        If AddToWaitList(OrderLink, OriginalOrder1, eGDWaitListCommand_AmendBoth, NewOrder1, OriginalOrder2, NewOrder2) Then
            DumpDebug "     Added to WaitList: '" & OrderLink.ToString & "', '" & OriginalOrder1.OrderText & "', '" & OriginalOrder2.OrderText & "'"
            DumpDebug "     Unlinking Orders (Amend): '" & OrderLink.ToString & "'"
            SendUnlinkRequest OrderLink
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinks.UnlinkAmendBothAndRelink"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnlinkAndParkOne
'' Description: Unlink two orders and then park the one passed in
'' Inputs:      Order
'' Returns:     True if Order Link found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UnlinkAndParkOne(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim OrderLink As cOrderLink         ' Order Link object
    Dim bReturn As Boolean              ' Return value for the function
    
    DumpDebug "OrderLinks: UnlinkAndParkOne: '" & Order.OrderText & "'"
    bReturn = False
    Set OrderLink = OrderLinkForOrder(Order)
    
    If Not OrderLink Is Nothing Then
        DumpDebug "     OrderLink: '" & OrderLink.ToString & "'"
        
        If AddToWaitList(OrderLink, Order, eGDWaitListCommand_ParkOne) Then
            DumpDebug "     Added to WaitList: '" & OrderLink.ToString & "', '" & Order.OrderText & "'"
            DumpDebug "     Unlinking Orders (ParkOne): '" & OrderLink.ToString & "'"
            SendUnlinkRequest OrderLink
        End If
        
        bReturn = True
    End If
    
    UnlinkAndParkOne = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinks.UnlinkAndParkOne"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnlinkAndParkBoth
'' Description: Park both orders of a link
'' Inputs:      Order, Other Order
'' Returns:     True if Order Link found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UnlinkAndParkBoth(ByVal Order As cPtOrder, ByVal OtherOrder As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim OrderLink As cOrderLink         ' Order Link object
    Dim bReturn As Boolean              ' Return value for the function
    
    DumpDebug "OrderLinks: UnlinkAndParkBoth: '" & Order.OrderText & "', '" & OtherOrder.OrderText & "'"
    bReturn = False
    Set OrderLink = OrderLinkForOrder(Order)
    
    If Not OrderLink Is Nothing Then
        DumpDebug "     OrderLink: '" & OrderLink.ToString & "'"
        
        If AddToWaitList(OrderLink, Order, eGDWaitListCommand_ParkBoth, OtherOrder) Then
            DumpDebug "     Added to WaitList: '" & OrderLink.ToString & "', '" & Order.OrderText & "', '" & OtherOrder.OrderText & "'"
            DumpDebug "     Unlinking Orders (ParkBoth): '" & OrderLink.ToString & "'"
            SendUnlinkRequest OrderLink
        End If
        
        bReturn = True
    End If
    
    UnlinkAndParkBoth = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinks.UnlinkAndParkBoth"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnlinkAndSubmitOne
'' Description: Unlink two orders and then submit the one passed in
'' Inputs:      Order
'' Returns:     True if Order Link found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UnlinkAndSubmitOne(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim OrderLink As cOrderLink         ' Order Link object
    Dim bReturn As Boolean              ' Return value for the function
    
    DumpDebug "OrderLinks: UnlinkAndSubmitOne: '" & Order.OrderText & "'"
    bReturn = False
    Set OrderLink = OrderLinkForOrder(Order)
    DumpDebug "     OrderLink: '" & OrderLink.ToString & "'"
    
    If Not OrderLink Is Nothing Then
        If AddToWaitList(OrderLink, Order, eGDWaitListCommand_SubmitOne) Then
            DumpDebug "     Added to WaitList: '" & OrderLink.ToString & "', '" & Order.OrderText & "'"
            DumpDebug "     Unlinking Orders (SubmitOne): '" & OrderLink.ToString & "'"
            SendUnlinkRequest OrderLink
        End If
        
        bReturn = True
    End If
    
    UnlinkAndSubmitOne = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinks.UnlinkAndSubmitOne"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnlinkAndSubmitBoth
'' Description: Submit both orders of a link
'' Inputs:      Order, Other Order
'' Returns:     True if Order Link found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UnlinkAndSubmitBoth(ByVal Order As cPtOrder, ByVal OtherOrder As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim OrderLink As cOrderLink         ' Order Link object
    Dim bReturn As Boolean              ' Return value for the function
    
    DumpDebug "OrderLinks: UnlinkAndSubmitBoth: '" & Order.OrderText & "', '" & OtherOrder.OrderText & "'"
    bReturn = False
    Set OrderLink = OrderLinkForOrder(Order)
    DumpDebug "     OrderLink: '" & OrderLink.ToString & "'"
    
    If Not OrderLink Is Nothing Then
        If AddToWaitList(OrderLink, Order, eGDWaitListCommand_SubmitBoth, OtherOrder) Then
            DumpDebug "     Added to WaitList: '" & OrderLink.ToString & "', '" & Order.OrderText & "', '" & OtherOrder.OrderText & "'"
            DumpDebug "     Unlinking Orders (SubmitBoth): '" & OrderLink.ToString & "'"
            SendUnlinkRequest OrderLink
        End If
        
        bReturn = True
    End If
    
    UnlinkAndSubmitBoth = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinks.UnlinkAndSubmitBoth"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderLinkForOrder
'' Description: Determine if order exists in an order link
'' Inputs:      Order
'' Returns:     Order Link or Nothing if not found
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderLinkForOrder(ByVal Order As cPtOrder) As cOrderLink
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim OrderLink As cOrderLink         ' Order link object
    
    Set OrderLink = Nothing
    For lIndex = 1 To m.OrderLinks.Count
        If (Item(lIndex).DatabaseId1 = Order.OrderID) Or (Item(lIndex).DatabaseId2 = Order.OrderID) Then
            Set OrderLink = Item(lIndex)
            Exit For
        End If
    Next lIndex
    
    Set OrderLinkForOrder = OrderLink

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinks.OrderLinkForOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderLinkForBrokerID
'' Description: Get the order link that contains the given Broker ID
'' Inputs:      Broker ID
'' Returns:     Order Link or Nothing if not found
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderLinkForBrokerID(ByVal strBrokerID As String) As cOrderLink
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim OrderLink As cOrderLink         ' Order link object
    
    Set OrderLink = Nothing
    For lIndex = 1 To m.OrderLinks.Count
        If (Item(lIndex).BrokerId1 = strBrokerID) Or (Item(lIndex).BrokerId2 = strBrokerID) Then
            Set OrderLink = Item(lIndex)
            Exit For
        End If
    Next lIndex
    
    Set OrderLinkForBrokerID = OrderLink

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinks.OrderLinkForBrokerID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderLinkForDatabaseID
'' Description: Get the order link that contains the given database ID
'' Inputs:      Database ID
'' Returns:     Order Link or Nothing if not found
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderLinkForDatabaseID(ByVal lDatabaseID As Long) As cOrderLink
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim OrderLink As cOrderLink         ' Order link object
    
    Set OrderLink = Nothing
    For lIndex = 1 To m.OrderLinks.Count
        If (Item(lIndex).DatabaseId1 = lDatabaseID) Or (Item(lIndex).DatabaseId2 = lDatabaseID) Then
            Set OrderLink = Item(lIndex)
            Exit For
        End If
    Next lIndex
    
    Set OrderLinkForDatabaseID = OrderLink

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinks.OrderLinkForDatabaseID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OtherDatabaseId
'' Description: Retrieve the other database ID from an order link
'' Inputs:      Order Link, Database ID
'' Returns:     Other Database ID (Zero if not applicable)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OtherDatabaseID(OrderLink As cOrderLink, ByVal lDatabaseID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function

    lReturn = 0&
    If lDatabaseID = OrderLink.DatabaseId1 Then
        lReturn = OrderLink.DatabaseId2
    ElseIf lDatabaseID = OrderLink.DatabaseId2 Then
        lReturn = OrderLink.DatabaseId1
    End If
    
    OtherDatabaseID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinks.OtherDatabaseId"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendLinkRequest
'' Description: Send a link request
'' Inputs:      Order Link Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SendLinkRequest(OrderLink As cOrderLink)
On Error GoTo ErrSection:

    If OrderLink.Status = eGDOrderLinkStatus_New Then
        If (Len(OrderLink.BrokerId1) > 0) And (Len(OrderLink.BrokerId2) > 0) Then
            g.PFG.LinkOrders OrderLink.BrokerId1, OrderLink.BrokerId2
            
            OrderLink.Status = eGDOrderLinkStatus_LinkSent
            OrderLink.Save
            
            Set m.OrderLinks(Str(OrderLink.OrderLinkID)) = OrderLink
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinks.SendLinkRequest"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendUnlinkRequest
'' Description: Send an unlink request
'' Inputs:      Order Link Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SendUnlinkRequest(OrderLink As cOrderLink)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object

    Select Case OrderLink.Status
        Case eGDOrderLinkStatus_Confirmed
            If (Len(OrderLink.BrokerId1) > 0) And (Len(OrderLink.BrokerId2) > 0) Then
                g.PFG.UnlinkOrders OrderLink.BrokerId1, OrderLink.BrokerId2
                
                OrderLink.Status = eGDOrderLinkStatus_UnlinkSent
                OrderLink.Save
                
                Set m.OrderLinks(Str(OrderLink.OrderLinkID)) = OrderLink
            End If
            
        ' 12/09/2009 DAJ: If the link was parked, break the link and confirm the unlink...
        Case eGDOrderLinkStatus_Parked
            Set Order = New cPtOrder
            If Order.Load(OrderLink.DatabaseId1) Then
                Order.BrokerCancelOrderID = 0&
                Order.Save
                OrderCallback Order
            End If
            UnlinkConfirm OrderLink.BrokerId1, OrderLink.BrokerId2
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinks.SendUnlinkRequest"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveLink
'' Description: Remove the order link
'' Inputs:      Order Link
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveLink(OrderLink As cOrderLink)
On Error GoTo ErrSection:

    m.OrderLinks.Remove Str(OrderLink.OrderLinkID)
    OrderLink.Delete

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinks.RemoveLink"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddToWaitList
'' Description: Add the specified information to the wait for unlink list
'' Inputs:      Order Link, Order, Command
'' Returns:     True if added, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddToWaitList(ByVal OrderLink As cOrderLink, ByVal Order As cPtOrder, ByVal nCommand As eGDWaitListCommands, Optional ByVal NewOrder As cPtOrder = Nothing, Optional ByVal OriginalOrder2 As cPtOrder = Nothing, Optional ByVal NewOrder2 As cPtOrder = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim WaitListItem As cOrderLinkWaitListItem
    
    bReturn = False
    If m.WaitList.Exists(Str(OrderLink.OrderLinkID)) = False Then
        Set WaitListItem = New cOrderLinkWaitListItem
        WaitListItem.OrderLinkID = OrderLink.OrderLinkID
        WaitListItem.OriginalOrder = Order
        WaitListItem.NewOrder = NewOrder
        WaitListItem.Command = nCommand
        WaitListItem.OriginalOrder2 = OriginalOrder2
        WaitListItem.NewOrder2 = NewOrder2
    
        m.WaitList.Add WaitListItem
        DumpDebug "     Item added to Wait For Unlink List: " & WaitListItem.ToString
        bReturn = True
    Else
        If (nCommand = eGDWaitListCommand_AmendOne) Or (nCommand = eGDWaitListCommand_AmendBoth) Then
            Set WaitListItem = m.WaitList.Item(Str(OrderLink.OrderLinkID))
            If (Order.OrderID <> WaitListItem.OriginalOrder.OrderID) And (WaitListItem.OriginalOrder.OrderID = OrderLink.OtherDatabaseID(Order.OrderID)) Then
                WaitListItem.OriginalOrder2 = Order
                WaitListItem.NewOrder2 = NewOrder
                WaitListItem.Command = eGDWaitListCommand_AmendBoth
                
                m.WaitList.Item(Str(OrderLink.OrderLinkID)) = WaitListItem
                DumpDebug "     Other order added to Wait For Unlink item: " & WaitListItem.ToString
                bReturn = True
            End If
        ElseIf nCommand = eGDWaitListCommand_Cancel Then
            Set WaitListItem = m.WaitList.Item(Str(OrderLink.OrderLinkID))
            If (Order.OrderID <> WaitListItem.OriginalOrder.OrderID) And (WaitListItem.OriginalOrder.OrderID = OrderLink.OtherDatabaseID(Order.OrderID)) Then
                WaitListItem.OriginalOrder2 = Order
                WaitListItem.NewOrder2 = NewOrder
                
                m.WaitList.Item(Str(OrderLink.OrderLinkID)) = WaitListItem
                DumpDebug "     Other order added to Wait For Unlink item: " & WaitListItem.ToString
                bReturn = True
            End If
        End If
    End If
    
    AddToWaitList = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinks.AddToWaitList"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PerformCommandForWaitList
'' Description: Perform the appropriate command for a waitlist item
'' Inputs:      Old Order Link ID
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PerformCommandForWaitList(ByVal lOrderLinkID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim WaitListItem As cOrderLinkWaitListItem
    
    DumpDebug "OrderLinks: PerformCommandForWaitList(" & Str(lOrderLinkID) & ")"
    bReturn = False
    Set WaitListItem = m.WaitList.Item(Str(lOrderLinkID))
    If Not WaitListItem Is Nothing Then
        Select Case WaitListItem.Command
            Case eGDWaitListCommand_Link
                DumpDebug "     Create New Link"
                If CreateLinkFromWaitList(WaitListItem.OriginalOrder) Then
                    bReturn = m.WaitList.Remove(Str(lOrderLinkID))
                End If
            
            Case eGDWaitListCommand_Cancel
                WaitListItem.OriginalOrder.Reload
                DumpDebug "     Cancel Order: '" & WaitListItem.OriginalOrder.OrderText & "'"
                CancelOrder WaitListItem.OriginalOrder, False, , True
                If Not WaitListItem.OriginalOrder2 Is Nothing Then
                    CancelOrder WaitListItem.OriginalOrder2, False, , True
                End If
                bReturn = m.WaitList.Remove(Str(lOrderLinkID))
                
            Case eGDWaitListCommand_ParkOne
                WaitListItem.OriginalOrder.Reload
                DumpDebug "     Park Order: '" & WaitListItem.OriginalOrder.OrderText & "'"
                ParkOrder WaitListItem.OriginalOrder
                bReturn = m.WaitList.Remove(Str(lOrderLinkID))
                
            Case eGDWaitListCommand_ParkBoth
                WaitListItem.OriginalOrder.Reload
                WaitListItem.NewOrder.Reload
                DumpDebug "     Park Order: '" & WaitListItem.OriginalOrder.OrderText & "'"
                ParkOrder WaitListItem.OriginalOrder
                DumpDebug "     Park Order: '" & WaitListItem.NewOrder.OrderText & "'"
                ParkOrder WaitListItem.NewOrder
                bReturn = m.WaitList.Remove(Str(lOrderLinkID))
            
            Case eGDWaitListCommand_AmendOne
                WaitListItem.OriginalOrder.Reload
                WaitListItem.NewOrder.BrokerCancelOrderID = WaitListItem.OriginalOrder.BrokerCancelOrderID
                DumpDebug "     AmendOrder: '" & WaitListItem.OriginalOrder.OrderText & "' --> '" & WaitListItem.NewOrder.OrderText & "' (" & WaitListItem.NewOrder.BrokerCancelOrderID & ")"
                SubmitAmend WaitListItem.OriginalOrder, WaitListItem.NewOrder
                bReturn = True
                
            Case eGDWaitListCommand_AmendBoth
                WaitListItem.OriginalOrder.Reload
                WaitListItem.NewOrder.BrokerCancelOrderID = WaitListItem.OriginalOrder.BrokerCancelOrderID
                DumpDebug "     AmendOrder: '" & WaitListItem.OriginalOrder.OrderText & "' --> '" & WaitListItem.NewOrder.OrderText & "' (" & WaitListItem.NewOrder.BrokerCancelOrderID & ")"
                SubmitAmend WaitListItem.OriginalOrder, WaitListItem.NewOrder
                
                WaitListItem.OriginalOrder2.Reload
                WaitListItem.NewOrder2.BrokerCancelOrderID = WaitListItem.OriginalOrder2.BrokerCancelOrderID
                DumpDebug "     AmendOrder: '" & WaitListItem.OriginalOrder2.OrderText & "' --> '" & WaitListItem.NewOrder2.OrderText & "' (" & WaitListItem.NewOrder2.BrokerCancelOrderID & ")"
                SubmitAmend WaitListItem.OriginalOrder2, WaitListItem.NewOrder2
                bReturn = True
                
        
            Case eGDWaitListCommand_SubmitOne
                WaitListItem.OriginalOrder.Reload
                DumpDebug "     Submit Order: '" & WaitListItem.OriginalOrder.OrderText & "'"
                SubmitOrder WaitListItem.OriginalOrder
                bReturn = m.WaitList.Remove(Str(lOrderLinkID))
                
            Case eGDWaitListCommand_SubmitBoth
                WaitListItem.OriginalOrder.Reload
                WaitListItem.NewOrder.Reload
                DumpDebug "     Submit Order: '" & WaitListItem.OriginalOrder.OrderText & "'"
                SubmitOrder WaitListItem.OriginalOrder
                DumpDebug "     Submit Order: '" & WaitListItem.NewOrder.OrderText & "'"
                SubmitOrder WaitListItem.NewOrder
                bReturn = m.WaitList.Remove(Str(lOrderLinkID))
            
        End Select
    End If
    
    PerformCommandForWaitList = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinks.PerformCommandForWaitList"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateLinkFromWaitList
'' Description: Create a new link from a wait list entry
'' Inputs:      Order
'' Returns:     True if created, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateLinkFromWaitList(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim OtherOrder As New cPtOrder      ' Second order
    Dim OrderLink As cOrderLink         ' Order link object
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If OtherOrder.Load(Order.BrokerCancelOrderID) Then
        If IsOpenOrder(Order.Status) And IsOpenOrder(OtherOrder.Status) Then
            Set OrderLink = New cOrderLink
            
            OrderLink.DatabaseId1 = Order.OrderID
            OrderLink.BrokerId1 = Order.BrokerID
            OrderLink.DatabaseId2 = OtherOrder.OrderID
            OrderLink.BrokerId2 = OtherOrder.BrokerID
            OrderLink.Status = eGDOrderLinkStatus_New
            
            OrderLink.Save
            
            DumpDebug "OrderLinks: Order Link Created: " & OrderLink.ToString
            Add OrderLink
            
            CheckOrderLinks
            
            bReturn = True
        End If
    End If
        
    CreateLinkFromWaitList = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinks.CreateLinkFromWaitList"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Dump the message to the appropriate logging place(s)
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpDebug(ByVal strMessage As String)
On Error GoTo ErrSection:

    'If FileExist(AddSlash(App.Path) & "BrokerOCO.FLG") Then
    '    DebugLog strMessage
    '    frmTest2.AddList strMessage
    'End If
    If Not g.PFG Is Nothing Then
        g.PFG.DumpDebug strMessage
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinks.DumpDebug"
    
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

    Set m.OrderLinks = New cGdTree
    Set m.WaitList = New cOrderLinkWaitList
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinks.Class_Initialize"
    
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

    Set m.OrderLinks = Nothing
    Set m.WaitList = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinks.Class_Terminate"
    
End Sub
