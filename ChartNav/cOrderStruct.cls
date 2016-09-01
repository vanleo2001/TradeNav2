VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOrderStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    Chart As cChart
    DailyBar As New cGdBars
    OrderStruct As chart_order
    
    nSymID As Long
    bAddedToDLL As Boolean
End Type
Dim m As mPrivate

Private Sub Class_Initialize()
    m.OrderStruct.nOrderId = -1
    m.OrderStruct.nLong = -1
    m.OrderStruct.nEntry = -1
    m.OrderStruct.nQty = -1
    m.OrderStruct.nStatus = -1
    m.OrderStruct.nOrderType = -1
    m.OrderStruct.nHorzLine = 0
    m.OrderStruct.dPrice = 0#
    
    m.OrderStruct.glhTriggerPId = gdCreateArray(eGDARRAY_Longs, , kNullData)
    m.OrderStruct.glhTriggerCId = gdCreateArray(eGDARRAY_Longs, , kNullData)
    m.OrderStruct.glhCancelId = gdCreateArray(eGDARRAY_Longs, , kNullData)
    m.OrderStruct.gshExtra = gdCreateArray(eGDARRAY_Strings)
End Sub

Private Sub Class_Terminate()
On Error Resume Next

    If Not m.Chart Is Nothing Then
        geRemoveItem m.Chart.geChartObj, 5, m.OrderStruct
    End If
    
    gdDestroyArray m.OrderStruct.glhTriggerPId
    gdDestroyArray m.OrderStruct.glhTriggerCId
    gdDestroyArray m.OrderStruct.glhCancelId
    gdDestroyArray m.OrderStruct.gshExtra
    
    Set m.Chart = Nothing
    Set m.DailyBar = Nothing
    
End Sub

Public Function OrderToStruct(Chart As cChart, Order As cPtOrder, _
    Optional ByVal bAddToDLL As Boolean = True, _
    Optional ByVal nLongOveride As Long = 0) As Long
'returns: 0=success, >0 = error code
On Error GoTo ErrSection:

    Dim TriggerChildren As cGdTree
    Dim ChildOrder As cPtOrder
    Dim ParentOrder As cPtOrder
    Dim strChartSym$
    Dim i&, j&, rc&, iDrawExitBracket&
    Dim dDate#

    If Chart Is Nothing Then
        OrderToStruct = 87          'system error code for invalid parameter
        Exit Function
    End If
    
    If m.Chart Is Nothing Then
        If Chart.SymbolID = 0 Then
            OrderToStruct = 1639    'system error code for unsupported type
            Exit Function           'for now not handling charts with symbolID=0 (e.g. spread charts)
        Else
            Set m.Chart = Chart
            m.nSymID = Chart.SymbolID
        End If
    End If
        
    strChartSym = m.Chart.Symbol
    If Not bAddToDLL Then
        Exit Function
    End If
    
    If m.bAddedToDLL And m.OrderStruct.nOrderId <> Order.OrderID Then
        OrderToStruct = 87          'system error code for invalid parameter
        Exit Function
    End If
    
    'set entry/exit flag
    m.OrderStruct.nEntry = Abs(OrderIsEntry(Order))
    'set long/short flag
    If Abs(Order.Buy) Then
        m.OrderStruct.nLong = m.OrderStruct.nEntry
    ElseIf m.OrderStruct.nEntry = 1 Then
        m.OrderStruct.nLong = 0
    Else
        m.OrderStruct.nLong = 1
    End If
    
    'flag for drawing bracket to right of a non Auto-Exit order if auto-exit check box is checked
    If Order.IsAutoExit Then
        iDrawExitBracket = 0
    Else
        iDrawExitBracket = Chart.Form.chkAutoExit.Value
    End If
    'override for drawing non-working order, eg parked order, as other color instead of buy or sell color
    If nLongOveride = -1 Then m.OrderStruct.nLong = -1
    
    m.OrderStruct.nOrderId = Order.OrderID
    m.OrderStruct.nQty = Order.RemainingQuantity
    m.OrderStruct.nOrderType = Order.OrderType
    
    If OrderIsPending(Order) Then
        m.OrderStruct.nStatus = -1
    Else
        m.OrderStruct.nStatus = Order.Status    'grapheng needs to know 'working' status
    End If
        
    
    'set symbol & brokerID info
    gdSetStr m.OrderStruct.gshExtra, 0, RollSymbolForDate(strChartSym)      '5199
    gdSetStr m.OrderStruct.gshExtra, 1, Order.Symbol
    gdSetStr m.OrderStruct.gshExtra, 2, Order.BrokerID
    gdSetStr m.OrderStruct.gshExtra, 3, iDrawExitBracket    'flag for drawing bracket to right of non auto-exit orders
    gdSetStr m.OrderStruct.gshExtra, 4, 0                   '1=connect auto exits with vertical line
    gdSetStr m.OrderStruct.gshExtra, 5, ""                  'if len > 0 then draw specified character on OCO line to indicate broker-held OCO
    
    'as implemented by the trade console
    '  trigger orders have one-to-many relationships (Parent <--->> child)
    '  OCO orders have one-to-one relationships
    gdSetSize m.OrderStruct.glhTriggerPId, 0, 0
    gdSetSize m.OrderStruct.glhTriggerCId, 0, 0
    gdSetSize m.OrderStruct.glhCancelId, 0, 0
    
    If Order.TriggerOrderID < 0 Then
        gdSetNum m.OrderStruct.glhTriggerPId, 0, 0
    Else
        gdSetNum m.OrderStruct.glhTriggerPId, 0, Order.TriggerOrderID                  '5756
    End If
    If Order.CancelOrderID <> 0 Then
        gdSetNum m.OrderStruct.glhCancelId, 0, Order.CancelOrderID
        gdSetStr m.OrderStruct.gshExtra, 5, ""
    ElseIf Order.BrokerCancelOrderID <> 0 Then
        gdSetNum m.OrderStruct.glhCancelId, 0, Abs(Order.BrokerCancelOrderID)
        If Order.BrokerCancelOrderID < 0 Then
            gdSetStr m.OrderStruct.gshExtra, 5, "*-"        'pending
        Else
            gdSetStr m.OrderStruct.gshExtra, 5, "*"
        End If
    End If
    
    Set TriggerChildren = g.Broker.TriggeredOrdersForOrderID(m.OrderStruct.nOrderId)
    For i = 1 To TriggerChildren.Count
        Set ChildOrder = TriggerChildren(i)
        j = gdGetSize(m.OrderStruct.glhTriggerCId)
        gdSetNum m.OrderStruct.glhTriggerCId, j, ChildOrder.OrderID
    Next
    
    'note: market order has a price of zero
    m.OrderStruct.dPrice = Order.OrderPrice(True)
    
    'set flag for whether to draw extended horizontal line across chart
    m.OrderStruct.nHorzLine = Chart.ShowOrderHorzLine
    
    If Not m.bAddedToDLL Then
        rc = geAddItem(m.Chart.geChartObj, 5, m.OrderStruct)
        If rc = 0 Then m.bAddedToDLL = True
    End If
    
    OrderToStruct = rc
        
ErrExit:
    Exit Function
        
ErrSection:
    RaiseError "cOrderStruct.OrderToStruct"

End Function

Public Property Get ItemID() As Long
    ItemID = m.OrderStruct.nOrderId
End Property

Public Function geDrawOrder(Chart As cChart) As Long
On Error Resume Next

    Dim rc&
       
    If Not Chart Is Nothing Then
        Chart.Form.pbChart.AutoRedraw = False
        rc = Chart.geDrawChart()
        Chart.Form.pbChart.AutoRedraw = True
    End If
        
    geDrawOrder = rc
    
End Function

Public Sub AlertTip(oToolTip As cToolTip, ByRef vCtrl As Control, ByVal nLoc&)
On Error Resume Next

    Dim l&, t&, b&, r&, nLen&
    Dim oAlert As cAlert
    Dim strText$
    Dim Order As cPtOrder

    If nLoc <> 1 And nLoc <> 2 Then Exit Sub

    Set oToolTip = Nothing
    If g.ChartGlobals.bChartTips Then
        t = m.OrderStruct.nPixTop
        b = m.OrderStruct.nPixBottom
        l = m.OrderStruct.nPixLeft
        r = m.OrderStruct.nPixRight
        If nLoc = 2 Then
            nLen = m.OrderStruct.nPixRight - m.OrderStruct.nPixLeft
            l = r + nLen * 2
            r = l + nLen * 2
        End If
        If l >= 0 And t >= 0 And b >= 0 And r >= 0 Then
            Set oToolTip = New cToolTip
            oToolTip.Create m.Chart.Form
            
            oToolTip.DelayTime(ttDelayInitial) = 0
            oToolTip.DelayTime(ttDelayReshow) = 0
            oToolTip.DelayTime(ttDelayShow) = 3000
            oToolTip.MaxTipWidth = 1
            
            If nLoc = 1 Then
                Set Order = New cPtOrder
                If Order.Load(m.OrderStruct.nOrderId) Then
                    strText = Order.OrderText
                Else
                    strText = "Error: load order " & Str(m.OrderStruct.nOrderId) & " failed."
                End If
            Else
                strText = "Click to cancel order"
            End If
            
            oToolTip.AddRect vCtrl, l, t, r, b, strText, False
        End If
    End If

End Sub

Public Property Get ItemPrice() As Double
    ItemPrice = m.OrderStruct.dPrice
End Property

Public Property Get ItemStatus() As eTT_OrderStatus
    ItemStatus = m.OrderStruct.nStatus
End Property

Public Property Get TradePrice(ePriceType As eBarsArray) As String
On Error GoTo ErrSection:

    Dim dPrice#, i&
    
    If Not m.DailyBar Is Nothing Then
        dPrice = kNullData
        m.DailyBar.ArrayMask = m.DailyBar.ArrayMask Or eBARS_Bid Or eBARS_Ask
        g.RealTime.UpdateBidAsk m.DailyBar, True
        
        Select Case ePriceType
            Case eBARS_Bid
                dPrice = m.DailyBar(eBARS_Bid, 0)
            Case eBARS_Ask
                dPrice = m.DailyBar(eBARS_Ask, 0)
            Case eBARS_Close
                i = m.Chart.LastGoodDataBar(False)
                dPrice = m.Chart.Bars(eBARS_Close, i)
        End Select
            
        If dPrice = kNullData Then
            TradePrice = "0.00"
        Else
            TradePrice = m.DailyBar.PriceDisplay(dPrice)
        End If
    End If
    
ErrExit:
    Exit Property
        
ErrSection:
    RaiseError "cOrderStruct.TradePrice"

End Property

Public Property Let ConnectAutoExit(ByVal iConnect&)
    gdSetStr m.OrderStruct.gshExtra, 4, Str(iConnect)
End Property
