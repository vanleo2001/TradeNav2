VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConditionalOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cConditionalOrders.cls
'' Description: Collection of conditional orders
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 05/25/2010   DAJ         Clear out conditional info after True
'' 10/12/2010   DAJ         Clear out trigger info as well after condition True
'' 05/30/2012   DAJ         Don't check order until data is available
'' 08/10/2012   DAJ         Don't call amend on trailing stop order not submitted yet
'' 07/30/2013   DAJ         Data Pending Order status for Conditional Orders
'' 08/01/2013   DAJ         Change to whether or not order is considered conditional
'' 08/01/2013   DAJ         Fix for whether conditional orders are data pending/trigger pending/working
'' 08/28/2013   DAJ         If not connected to the broker upon submit, park the order
'' 11/04/2013   DAJ         Add bars to stream if bars handle changed
'' 10/30/2015   DAJ         Don't have a triggered-by order trail until it is triggered
'' 11/06/2015   DAJ         Moved trailing stop code from order into conditional orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    ConditionalOrders As cGdTree        ' Collection of conditional orders
    EachTick As cGdTree                 ' Collection of each-tick bars for trailing stops
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cPtOrder
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Return value for the function
    
    If m.ConditionalOrders.Exists(vIndexOrKey) Then
        Set Order = m.ConditionalOrders(vIndexOrKey)
    Else
        Set Order = Nothing
    End If
    
    Set Item = Order

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cConditionalOrders.Item.Get"
    
End Property

Public Property Let Item(ByVal vIndexOrKey As Variant, Order As cPtOrder)
On Error GoTo ErrSection:

    Dim PrevOrder As cPtOrder           ' Previous order

    If m.ConditionalOrders.Exists(vIndexOrKey) Then
        Set PrevOrder = m.ConditionalOrders(vIndexOrKey)
        
        If Order.HasExpression = True Then
            If PrevOrder.BarsHandle <> Order.BarsHandle Then
                g.Broker.BrokerDebug Order.Broker, "Bars added to stream for " & Order.OrderText(True, True, True) & " because Bars handle changed from " & Str(PrevOrder.BarsHandle) & " to " & Order.BarsHandle
                Order.AddToStream
            End If
        End If
        
        m.ConditionalOrders(vIndexOrKey) = Order
    Else
        m.ConditionalOrders.Add Order, Str(Order.OrderID)
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cConditionalOrders.Item.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Remove
'' Description: Remove items from the collections with the given index or key
'' Inputs:      Index or Key
'' Returns:     True if removed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Remove(ByVal vIndexOrKey As Variant) As Boolean
On Error GoTo ErrSection:

    If m.ConditionalOrders.Exists(vIndexOrKey) Then
        m.EachTick.Remove m.ConditionalOrders.Key(vIndexOrKey)
    End If
    
    Remove = m.ConditionalOrders.Remove(vIndexOrKey)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cConditionalOrders.Remove"
    
End Function

Public Function Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.ConditionalOrders.Exists(vIndexOrKey)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshOrder
'' Description: Refresh the given order in the conditional collection (if there)
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshOrder(Order As cPtOrder)
On Error GoTo ErrSection:

    If m.ConditionalOrders.Exists(Str(Order.OrderID)) Then
        m.ConditionalOrders(Str(Order.OrderID)) = Order
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cConditionalOrders.RefreshConditionalOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckOrders
'' Description: Check the conditional orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckOrders()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bDoCheck As Boolean             ' Do a check on the conditional orders?
    Dim bDoTrail As Boolean             ' Do a check on the trail orders?
    Static sdLastCheck As Double        ' Last check for conditional orders
    Static sdLastTrail As Double        ' Last check for trail amount
    Dim dTrail As Double                ' Trailing stop value
    Dim Order As cPtOrder               ' Order object
    Dim strMessage As String            ' Message to show to the user
    
    bDoCheck = (Now > sdLastCheck + (1# / 86400#))
    bDoTrail = (Now > sdLastTrail + (1# / 86400#))
    
    For lIndex = m.ConditionalOrders.Count To 1 Step -1
        Set Order = Item(lIndex)
        If Not Order Is Nothing Then
            If Order.AllRtDataAvailable Then
                If Order.Status = eTT_OrderStatus_DataPending Then
                    Order.ChangeOrderStatus eTT_OrderStatus_TriggerPending
                End If
                
                Order.UpdateBars
                
                If bDoCheck Then
                    If ((Order.CheckTime = True) And (Order.IsConditional(True) = True)) And ((Len(Order.Condition) = 0) Or (Order.RunExpression = True)) And (Order.Status = eTT_OrderStatus_TriggerPending) Then
                        If Order.HasOpenTriggerOrder = False Then
                            If g.Broker.ConnectionStatusForAccount(Order.AccountID) = eGDConnectionStatus_Connected Then
                                g.Broker.BrokerDebug Order.Broker, "Submitting Conditional Order from Trade Console: " & Order.OrderText, True
                                
                                ' 05/24/2010 DAJ: Clear the conditional order stuff here so that it is never
                                ' considered a conditional order again...
                                g.Broker.BrokerDebug Order.Broker, vbTab & "Clearing out Conditional Order information (" & Order.ConditionOptions & ") because condition is True"
                                'Order.ConditionOptions = ""
                                Order.ClearConditionOptions
                                
                                ' 10/12/2010 DAJ: Also need to clear out the trigger information here in case this
                                ' was a conditional and a triggered order...
                                Order.TriggerOrderID = 0&
                                Order.TriggerOptions = ""
                                
                                Order.Save
                                
                                SubmitOrder Order
                            Else
                                Order.ChangeOrderStatus eTT_OrderStatus_Parked
                                strMessage = "Order has been parked:||" & Order.OrderText(True, True, False) & "||because you are not connected to the broker"
                                ShowAlertPopup strMessage, "Conditional Order", vbCenter
                                g.Broker.BrokerDebug Order.Broker, "Message shown to user: '" & Replace(strMessage, "||", " ") & "'"
                            End If
                        End If
                    ElseIf (Order.ExpireTime <> 0) And (Now > Order.ExpireTime) Then
                        g.Broker.BrokerDebug Order.Broker, "Cancelling Conditional Order because expire time has elapsed: " & Order.OrderText, True
                        CancelOrder Order, False
                    ElseIf bDoTrail And Order.TrailAmount <> 0 Then
                        If Order.HasOpenTriggerOrder = False Then
                            'dTrail = Order.TrailingStopValue(False)
                            dTrail = TrailingStopValue(Order)
                            If (dTrail <> Order.StopPrice) And (dTrail <> kNullData) Then
                                '*** PRICE OR QUANTITY ON ORDER CHANGED...***
                                Order.StopPrice = dTrail
                                
                                If Order.HasOpenTriggerOrder Then
                                    Order.Save
                                    OrderCallback Order
                                ElseIf IsOpenOrder(Order.Status, False) Then
                                    If g.Broker.ConnectionStatusForAccount(Order.AccountID) = eGDConnectionStatus_Connected Then
                                        If Order.Status = eTT_OrderStatus_Open Then
                                            g.Broker.BrokerDebug Order.Broker, "Creating Trailing Order from Trade Console: " & Order.OrderText, True
                                            SubmitOrder Order
                                        Else
                                            g.Broker.BrokerDebug Order.Broker, "Modifying Trailing Order from Trade Console: " & Order.OrderText, True
                                            EditOrder Order, , False, eGDEditOrderReturn_Submit
                                        End If
                                    Else
                                        Order.ChangeOrderStatus eTT_OrderStatus_Parked
                                        strMessage = "Order has been parked:||" & Order.OrderText(True, True, False) & "||because you are not connected to the broker"
                                        ShowAlertPopup strMessage, "Conditional Order", vbCenter
                                        g.Broker.BrokerDebug Order.Broker, "Message shown to user: '" & Replace(strMessage, "||", " ") & "'"
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next lIndex
    
    If bDoCheck Then
        sdLastCheck = Now
    End If
    If bDoTrail Then
        sdLastTrail = Now
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cConditionalOrders.CheckOrders"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIDChanged
'' Description: Notification that an order ID changed
'' Inputs:      Old Order ID, New Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderIDChanged(ByVal lOldOrderID As Long, ByVal lNewOrderID As Long)
On Error GoTo ErrSection:

    If m.ConditionalOrders.Exists(Str(lOldOrderID)) = True Then
        m.ConditionalOrders.Key(Str(lOldOrderID)) = Str(lNewOrderID)
    End If
    
    If m.EachTick.Exists(Str(lOldOrderID)) = True Then
        m.EachTick.Key(Str(lOldOrderID)) = Str(lNewOrderID)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cConditionalOrders.OrderIDChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddEachTick
'' Description: Add an each-tick bars object to the collection for the given
''              order if it is not already there
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddEachTick(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim Bars As cGdBars                 ' Bars to add to the collection
    Dim dSubmitTime As Double           ' Date and time the order was submitted
    Dim lTickNum As Long                ' Tick number of the given date time

    If m.EachTick.Exists(Str(Order.OrderID)) = False Then
        dSubmitTime = Val(Parse(Order.TrailOptions, ",", 2))
        
        ' Only add to the collection if the order has been submitted...
        If dSubmitTime <> 0 Then
            If g.RealTime.RtDataAvailable(Order.SymbolOrSymbolID, ePRD_EachTick) = True Then
                Set Bars = New cGdBars
                
                ' Load the each-tick bars starting with the day the order was submitted...
                DM_GetBars Bars, Order.SymbolOrSymbolID, ePRD_EachTick, Int(dSubmitTime)
            
                ' Add the bars object to the stream and splice them appropriately...
                g.RealTime.AddTickBuffer Bars
                g.RealTime.SpliceBars Bars
                
                ' Trim off any ticks that are in the bars before the order was sufbmitted...
                lTickNum = Bars.FindDateTime(dSubmitTime)
                If lTickNum >= Bars.Size - 1 Then
                    lTickNum = Bars.Size - 2
                End If
                
                g.Broker.BrokerDebug Order.Broker, vbTab & "EachTick bars size = " & Str(Bars.Size) & "; TickNum = " & Str(lTickNum) & "; SubmitTime = " & DateFormat(dSubmitTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
                
                Bars.DeleteFirstBars lTickNum
                
                ' Add the bars to the collection keyed off of the order Access ID...
                m.EachTick.Add Bars, Str(Order.OrderID)
                
                g.Broker.BrokerDebug Order.Broker, vbTab & "EachTick bars added for Order " & Str(Order.OrderID) & "; BarsHandle = " & Str(Bars.BarsHandle) & "; BarsSize = " & Str(Bars.Size)
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cConditionalOrders.AddEachTick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TrailingStopValue
'' Description: Determine the current value of the given trailing stop order
'' Inputs:      Order
'' Returns:     Trailing Stop Value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TrailingStopValue(ByVal Order As cPtOrder) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim Bars As cGdBars                 ' Bars to add to the collection
    Dim dTrailingStopValue As Double    ' Trailing stop value
        
    ' Add the order to the each tick collection if it is not already there...
    AddEachTick Order
    
    ' Default the return value to the current stop price of the order...
    dReturn = Order.StopPrice
    
    ' If the corresponding item in the each bar collection exists ( which it should by now )...
    If m.EachTick.Exists(Str(Order.OrderID)) Then
        If g.RealTime.RtDataAvailableForBars(m.EachTick(Str(Order.OrderID))) Then
            ' ... and the there are new ticks on the bars ...
            If g.RealTime.UpdateBars(m.EachTick(Str(Order.OrderID))) Then
                Set Bars = m.EachTick(Str(Order.OrderID))
                
                If Bars.Size = 0 Then
                    dReturn = dReturn
                End If
                
                dTrailingStopValue = Order.TrailingStopValue(gdMinValue(Bars.ArrayHandle(eBARS_Close), 0, Bars.Size), gdMaxValue(Bars.ArrayHandle(eBARS_Close), 0, Bars.Size))
                If dTrailingStopValue <> kNullData Then
                    dReturn = dTrailingStopValue
                    
                    If dReturn <> Order.StopPrice Then
                        g.Broker.BrokerDebug Order.Broker, vbTab & vbTab & "Trailing Stop Value changed from " & Str(Order.StopPrice) & " to " & Str(dReturn)
                    End If
                End If
            End If
        End If
    End If
    
    TrailingStopValue = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cConditionalOrders.TrailingStopValue"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.ConditionalOrders = New cGdTree
    Set m.EachTick = New cGdTree
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cConditionalOrders.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.ConditionalOrders = Nothing
    Set m.EachTick = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cConditionalOrders.Class_Terminate"
    
End Sub

