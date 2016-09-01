VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cActiveExitOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cActiveExitOrder.cls
'' Description: Handles an active exit order
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/04/2009   DAJ         Use the OrderID let property throughout
'' 05/17/2010   DAJ         Added support for TradeSense orders
'' 12/11/2012   DAJ         Use the flatten queue for position reversals
'' 12/17/2012   DAJ         Fix for setting order quantities when one auto exit order fills
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Order As cPtOrder                   ' The active exit order
    
    lOrderID As Long                    ' Order ID for the order
    lPrevOrderID As Long                ' Previous Order ID for when OrderID has been changed to -2
    lQuantity As Long                   ' Quantity for the order
    bReset As Boolean                   ' Reset the order?
    bCancelRequested As Boolean         ' Has a cancel been requested for this symbol?
    bModify As Boolean                  ' Need to modify the order when it goes working
    bWaiting As Boolean                 ' Are we waiting for a modify?
    bBreakEvenMoved As Boolean          ' Has the break even stop been moved?
    bTradeSenseMoved As Boolean         ' Has the Trade Sense order been moved?
    
    lAccountID As Long                  ' Account ID
    nAcctType As Long                   ' Account type
    strSymbol As String                 ' Symbol
    strName As String                   ' Name of the auto exit order
    strAbbreviation As String           ' Abbreviation of the auto exit order
End Type
Private m As mPrivate

Public Property Get Order() As cPtOrder
    Set Order = m.Order
End Property
Public Property Let Order(AutoExitOrder As cPtOrder)
    Set m.Order = AutoExitOrder
End Property

Public Property Get OrderID() As Long
    OrderID = m.lOrderID
End Property
Public Property Let OrderID(ByVal lOrderID As Long)
    If lOrderID <> m.lOrderID Then
        If lOrderID = -2 Then
            m.lPrevOrderID = m.lOrderID
        Else
            m.lPrevOrderID = kNullData
        End If
        
        m.lOrderID = lOrderID
    End If
End Property

Public Property Get PrevOrderID() As Long
    PrevOrderID = m.lPrevOrderID
End Property
Public Property Let PrevOrderID(ByVal lPrevOrderID As Long)
    m.lPrevOrderID = lPrevOrderID
End Property

Public Property Get Quantity() As Long
    Quantity = m.lQuantity
End Property
Public Property Let Quantity(ByVal lQuantity As Long)
    m.lQuantity = lQuantity
End Property

Public Property Get Reset() As Boolean
    Reset = m.bReset
End Property
Public Property Let Reset(ByVal bReset As Boolean)
    m.bReset = bReset
End Property

Public Property Get CancelRequested() As Boolean
    CancelRequested = m.bCancelRequested
End Property
Public Property Let CancelRequested(ByVal bCancelRequested As Boolean)
    m.bCancelRequested = bCancelRequested
End Property

Public Property Get Modify() As Boolean
    Modify = m.bModify
End Property
Public Property Let Modify(ByVal bModify As Boolean)
    m.bModify = bModify
End Property

Public Property Get Waiting() As Boolean
    Waiting = m.bWaiting
End Property
Public Property Let Waiting(ByVal bWaiting As Boolean)
    m.bWaiting = bWaiting
End Property

Public Property Get BreakEvenMoved() As Boolean
    BreakEvenMoved = m.bBreakEvenMoved
End Property
Public Property Let BreakEvenMoved(ByVal bBreakEvenMoved As Boolean)
    m.bBreakEvenMoved = bBreakEvenMoved
End Property

Public Property Get TradeSenseMoved() As Boolean
    TradeSenseMoved = m.bTradeSenseMoved
End Property
Public Property Let TradeSenseMoved(ByVal bTradeSenseMoved As Boolean)
    m.bTradeSenseMoved = bTradeSenseMoved
End Property

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal lAccountID As Long)
    m.lAccountID = lAccountID
    m.nAcctType = g.Broker.AccountTypeForID(lAccountID)
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Let Symbol(ByVal strSymbol As String)
    m.strSymbol = strSymbol
End Property

Public Property Get Name() As String
    Name = m.strName
End Property
Public Property Let Name(ByVal strName As String)
    m.strName = strName
End Property

Public Property Get Abbreviation() As String
    Abbreviation = m.strAbbreviation
End Property
Public Property Let Abbreviation(ByVal strAbbreviation As String)
    m.strAbbreviation = strAbbreviation
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the active exit order
'' Inputs:      Name, Abbreviation
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal strName As String, ByVal strAbbreviation As String)
On Error GoTo ErrSection:

    Name = strName
    Abbreviation = strAbbreviation

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExitOrder.Init"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load up the order given the Order ID
'' Inputs:      Order ID
'' Returns:     True if Loaded, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal lOrderID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    OrderID = lOrderID
    
    If lOrderID <> 0 Then
        Set m.Order = New cPtOrder
        bReturn = m.Order.Load(lOrderID)
    End If
    
    Load = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExitOrder.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderCallback
'' Description: Handle an order callback
'' Inputs:      Order, Name, Current Position
'' Returns:     0=No Action, 1=Cancel, 2=Modify
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderCallback(Order As cPtOrder, ByVal strName As String, ByVal lCurrentPosition As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function

    lReturn = 0&
    Set m.Order = Order
    
    DumpDebug "OrderCallback " & m.strAbbreviation & " (" & OrderStatus(Order.Status) & "): " & Order.OrderText
    
    ' If the order has gone working, we don't need to wait for it anymore...
    If IsOpenOrder(Order.Status, False) Then
        m.bWaiting = False
    End If
    
    If ((Order.Status = eTT_OrderStatus_Cancelled) And (m.bCancelRequested = False) And (Flattening = False)) Or (Order.Status = eTT_OrderStatus_Expired) Then
        DumpDebug vbTab & "Unsolicited Cancel or Expired"
        OrderID = -1&
    ElseIf (IsOpenOrder(Order.Status) = False) And (Order.Status <> eTT_OrderStatus_Amended) Then
        If m.bReset Or Flattening Then
            DumpDebug vbTab & "Closed Order, but not Amended (Reset)"
            OrderID = 0&
        Else
            DumpDebug vbTab & "Closed Order, but not Amended (No Reset)"
            OrderID = -2&
        End If
    ElseIf (IsOpenOrder(Order.Status, False) = True) And (lCurrentPosition = 0) Then
        DumpDebug vbTab & "Open Order and Flat"
        
        CancelOrder
        If g.Broker.AutoJournalAutomated Then
            g.JournalBridge.CreateOrderJournal m.Order.OrderID, CurrentTime, "Auto exit is cancelling " & m.strName & " order because position is now flat"
        End If
        
        lReturn = 1&
    ElseIf m.bModify Then
        DumpDebug vbTab & "Modify"
        lReturn = 2&
    End If
    
    OrderCallback = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExitOrder.OrderCallback"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIdChanged
'' Description: The order ID has changed for the order
'' Inputs:      New Order ID
'' Returns:     True if Changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderIDChanged(ByVal lNewOrderID As Long) As Boolean
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order object
    
    If Order.Load(lNewOrderID) = True Then
        DumpDebug "OrderID Changed " & m.strAbbreviation & " - Old: " & Str(m.lOrderID) & "; New: " & Str(lNewOrderID)
        Set m.Order = Order
        OrderID = lNewOrderID
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cActiveExitOrder.OrderIdChanged"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetInfo
'' Description: Reset the order information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ResetInfo()
On Error GoTo ErrSection:

    OrderID = 0&
    m.bReset = False
    m.bCancelRequested = False
    DumpDebug "Cancel Requested on " & m.strName & " = False"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExitOrder.ResetInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelWorkingOrder
'' Description: Attempt to cancel the order
'' Inputs:      Deactivating, Skip Order ID, Reset ID's, Cancelling All?, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelWorkingOrder(ByVal bDeactivating As Boolean, Optional ByVal lSkipOrderID As Long = -1&, Optional ByVal bResetIDs As Boolean = True, Optional ByVal bCancellingAll As Boolean = False, Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    If (m.lOrderID <> 0) And (m.lOrderID <> lSkipOrderID) And (Not m.Order Is Nothing) Then
        If Len(strReason) = 0 Then
            DumpDebug "Cancelling " & m.strAbbreviation & " (" & OrderStatus(m.Order.Status) & "): " & m.Order.OrderText
        Else
            DumpDebug "Cancelling " & m.strAbbreviation & " (" & OrderStatus(m.Order.Status) & "): " & m.Order.OrderText & "(" & strReason & ")"
        End If
        
        If m.bCancelRequested = False Then
            m.bCancelRequested = True
            DumpDebug "Cancel Requested on " & m.strName & " = True"
        End If
        
        ' We cannot let a "sent" order stop us from deactivating the auto exit, but we have to wait
        ' for a "sent" order to go "working" if we are not deactivating so that if the position goes
        ' to zero before it goes "working", we can cancel it for real right after it goes "working"...
        If IsOpenOrder(m.Order.Status, False) Or ((m.Order.Status = eTT_OrderStatus_Sent) And (bDeactivating = True)) Then
            CancelOrder bCancellingAll
            If bResetIDs Then OrderID = -2&
            m.bReset = False
            
            If g.Broker.AutoJournalAutomated Then
                g.JournalBridge.CreateOrderJournal m.Order.OrderID, CurrentTime, "Auto Exit cancelling order"
            End If
            
        ' If an order has just gone to Amended, we don't want to consider it "closed" yet because
        ' it will be probably changing ID's and then going Working.  Once it goes Working, then we
        ' will cancel it properly (DAJ: 05/07/2009)...
        ElseIf (OrderIsPending(m.Order) = False) And (m.Order.Status <> eTT_OrderStatus_Amended) Then
            If m.bReset Then
                OrderID = 0&
            Else
                OrderID = -2&
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExitOrder.CancelWorkingOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelFromUpdateSession
'' Description: Attempt to cancel the order because session date changed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelFromUpdateSession()
On Error GoTo ErrSection:

    If (m.lOrderID > -2&) And (Not m.Order Is Nothing) Then
        DumpDebug "Cancelling " & m.strAbbreviation & " (" & OrderStatus(m.Order.Status) & "): " & m.Order.OrderText
        If IsOpenOrder(m.Order.Status, False) And (m.lOrderID > 0&) Then
            CancelOrder
        End If
        
        OrderID = 0&
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExitOrder.CancelFromUpdateSession"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckModify
'' Description: Check to see if we can and should modify the order
'' Inputs:      None
'' Returns:     True if can modify, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckModify() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If m.bModify = True Then
        If m.lOrderID > 0 Then
            If OrderIsPending(m.Order) = False Then
                If CanModify(m.Order, , False) Then
                    bReturn = True
                End If
            End If
        End If
    End If
    
    CheckModify = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExitOrder.CheckModify"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ModifyTheOrder
'' Description: Attempt to modify the given order
'' Inputs:      None
'' Returns:     True if need to create new order, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ModifyTheOrder(Optional ByVal lOtherID As Long = 0&) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    m.bModify = False
    If (m.lQuantity = 0&) And (m.lOrderID > -2&) Then
        If (lOtherID = -2) And (m.Order.BrokerCancelOrderID > 0) Then
            DumpDebug "ModifyTheOrder " & m.strAbbreviation & " --> Not cancelling because broker OCO should do that"
        ElseIf CanModify(m.Order) And (m.lOrderID > 0&) Then
            DumpDebug "Modify Orders --> Cancel Stop Order"
            m.bReset = True
            CancelOrder
            If g.Broker.AutoJournalAutomated Then
                g.JournalBridge.CreateOrderJournal m.Order.OrderID, CurrentTime, "Auto exit is cancelling " & m.strName & " because it is no longer needed"
            End If
        ElseIf (m.lOrderID = -1&) Then
            OrderID = 0&
        ElseIf (m.lOrderID > 0) Then
            DumpDebug "ModifyTheOrder " & m.strAbbreviation & " --> Need to Cancel " & m.strName & " when not Pending"
            m.bModify = True
        End If
    ElseIf (m.lOrderID > 0&) Then
        If (m.Order.RemainingQuantity <> m.lQuantity) Then
            If CanModify(m.Order) Then
                DumpDebug "ModifyTheOrder " & m.strAbbreviation & " --> Modify " & m.strName & " OldQty: " & Str(m.Order.Quantity) & ", NewQty: " & Str(m.lQuantity)
                m.Order.Quantity = m.lQuantity
                m.bWaiting = True
                If ModifyOrder = False Then
                    DumpDebug "ModifyTheOrder " & m.strAbbreviation & " --> Need to Modify " & m.strName & " when possible"
                    m.bModify = True
                Else
                    m.bWaiting = False
                    If g.Broker.AutoJournalAutomated Then
                        g.JournalBridge.CreateOrderJournal m.Order.OrderID, CurrentTime, m.strName & " quantity is being updated to " & m.Order.Quantity & " by auto exit because position changed"
                    End If
                End If
            Else
                DumpDebug "ModifyTheOrder " & m.strAbbreviation & " --> Need to Modify " & m.strName & " when possible"
                m.bModify = True
            End If
        End If
    ElseIf (m.lQuantity > 0&) Then
        If m.lOrderID = 0& Then
            DumpDebug "ModifyTheOrder " & m.strAbbreviation & " --> Creating " & m.strName
            m.bWaiting = True
            bReturn = True
        Else
            DumpDebug "ModifyTheOrder " & m.strAbbreviation & " --> Could not create " & m.strName & " because one already exists"
        End If
    End If
    
    ModifyTheOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExitOrder.ModifyTheOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TrailOrder
'' Description: Attempt to modify the order to trail if necessary
'' Inputs:      New Stop Price, New With Limit Price
'' Returns:     True if Sent, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TrailOrder(ByVal dStopPrice As Double, ByVal dWithLimitPrice As Double) As Boolean
On Error GoTo ErrSection:

    Dim bBetterPrice As Boolean         ' Is this a better price than the stop?
    Dim bReturn As Boolean              ' Return value for the function
    Dim bPending As Boolean             ' Did CanModify fail because order is pending?
    
    bReturn = False
    If m.lOrderID > 0 Then
        If m.Order.Buy Then
            bBetterPrice = ((RoundToSigDigits(dStopPrice) < RoundToSigDigits(m.Order.StopPrice)) Or (m.Order.StopPrice = 0))
        Else
            bBetterPrice = (RoundToSigDigits(dStopPrice) > RoundToSigDigits(m.Order.StopPrice))
        End If
        
        If (bBetterPrice = True) Then
            If m.bWaiting = False Then
                DumpDebug "Modifying Order " & m.strAbbreviation & " (TrailOrder) - Old Stop Price: " & m.Order.StopPrice & ", New Stop Price: " & dStopPrice
                If CanModify(m.Order, dStopPrice, False, bPending) = True Then
                    m.Order.StopPrice = dStopPrice
                    m.Order.LimitPrice = dWithLimitPrice
                    m.Order.Quantity = m.lQuantity
                    
                    m.bWaiting = True
                    If ModifyOrder And g.Broker.AutoJournalAutomated Then
                        g.JournalBridge.CreateOrderJournal m.Order.OrderID, CurrentTime, "Auto exit has moved Stop Loss order to " & m.Order.StopPriceString & " to keep order trailing the market"
                        bReturn = True
                    Else
                        m.bWaiting = False
                    End If
                End If
            Else
                DumpDebug "Modifying Order " & m.strAbbreviation & " (TrailOrder) - Could not move order because we are waiting for it to go working"
                m.bWaiting = bPending
            End If
        End If
    Else
        'DumpDebug "Modifying Order " & m.strAbbreviation & " (TrailOrder) - Cannot move order because order ID <=0"
    End If
    
    TrailOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExitOrder.TrailOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrder
'' Description: Submit the order
'' Inputs:      Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SubmitOrder(Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    Waiting = True
    m.Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(AccountID))
    mTradeTracker.SubmitOrder m.Order
    
    If g.Broker.AutoJournalAutomated Then
        If Len(strReason) > 0 Then
            g.JournalBridge.CreateOrderJournal m.Order.OrderID, CurrentTime, m.strName & " order generated by auto exit at " & m.Order.StopPriceString & " because " & strReason
        Else
            g.JournalBridge.CreateOrderJournal m.Order.OrderID, CurrentTime, m.strName & " order generated by auto exit at " & m.Order.StopPriceString
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExitOrder.SubmitOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MoveBreakEven
'' Description: Attempt to modify the break even stop if possible
'' Inputs:      New Stop Price, New With Limit Price, Equation
'' Returns:     True if moved, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MoveBreakEven(ByVal dStopPrice As Double, ByVal dWithLimitPrice As Double, ByVal strEquation As String) As Boolean
On Error GoTo ErrSection:

    Dim bPending As Boolean             ' Did CanModify fail because order is pending?

    If m.lOrderID > 0 Then
        If m.bWaiting = False Then
            DumpDebug "Modifying Order " & m.strAbbreviation & " (BreakEven) - MoveOrder: " & m.Order.OrderText & " (" & strEquation & ")"
            
            If CanModify(m.Order, dStopPrice, False, bPending) Then
                m.Order.StopPrice = dStopPrice
                m.Order.LimitPrice = dWithLimitPrice
            
                m.Order.Quantity = m.lQuantity
                
                m.bWaiting = True
                If ModifyOrder And g.Broker.AutoJournalAutomated Then
                    g.JournalBridge.CreateOrderJournal m.Order.OrderID, CurrentTime, "Auto exit is moving Stop Loss order to " & m.Order.StopPriceString & " because market has reached the break even price"
                Else
                    m.bWaiting = False
                End If
                
                m.bBreakEvenMoved = True
            Else
                DumpDebug "Modifying Order " & m.strAbbreviation & " (BreakEven) - Cannot move order because we need to wait for it to go working (" & strEquation & ")"
                m.bWaiting = bPending
            End If
        Else
            DumpDebug "Modifying Order " & m.strAbbreviation & " (BreakEven) - Cannot move order because we are waiting for it to go working (" & strEquation & ")"
        End If
    Else
        DumpDebug "Modifying Order " & m.strAbbreviation & " (BreakEven) - Cannot move order because order ID <=0 (" & strEquation & ")"
        
        ' DAJ 11/02/2009: Going to set this to True here so that it won't keep trying to come in here...
        m.bBreakEvenMoved = True
    End If
    
    MoveBreakEven = m.bBreakEvenMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExitOrder.MoveBreakEven"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MoveTradeSense
'' Description: Attempt to modify the Trade Sense order if possible
'' Inputs:      New Price, New With Limit Price, Equation
'' Returns:     True if moved, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MoveTradeSense(ByVal dPrice As Double, ByVal dWithLimitPrice As Double, ByVal strEquation As String) As Boolean
On Error GoTo ErrSection:

    Dim bPending As Boolean             ' Did CanModify fail because order is pending?

    If m.lOrderID > 0 Then
        If OrderIsDifferent(dPrice, dWithLimitPrice, m.lQuantity) Then
            If m.bWaiting = False Then
                DumpDebug "Modifying Order " & m.strAbbreviation & " (TradeSense) - MoveOrder: " & m.Order.OrderText & " (" & strEquation & ")"
            
                If CanModify(m.Order, dPrice, False, bPending) Then
                    Select Case m.Order.OrderType
                        Case eTT_OrderType_Market
                            m.Order.StopPrice = 0#
                            m.Order.LimitPrice = 0#
                            
                        Case eTT_OrderType_Stop
                            m.Order.StopPrice = dPrice
                            m.Order.LimitPrice = 0#
                            
                        Case eTT_OrderType_Limit
                            m.Order.StopPrice = 0#
                            m.Order.LimitPrice = dPrice
                            
                        Case eTT_OrderType_StopWithLimit
                            m.Order.StopPrice = dPrice
                            m.Order.LimitPrice = dWithLimitPrice
                    End Select
                
                    m.Order.Quantity = m.lQuantity
                    
                    m.bWaiting = True
                    If ModifyOrder And g.Broker.AutoJournalAutomated Then
                        g.JournalBridge.CreateOrderJournal m.Order.OrderID, CurrentTime, "Auto exit is moving Stop Loss order to " & m.Order.StopPriceString & " because market has reached the break even price"
                    Else
                        m.bWaiting = False
                    End If
                    
                    m.bTradeSenseMoved = True
                Else
                    DumpDebug "Modifying Order " & m.strAbbreviation & " (TradeSense) - Cannot move order because we need to wait for it to go working (" & strEquation & ")"
                    m.bWaiting = bPending
                End If
            Else
                DumpDebug "Modifying Order " & m.strAbbreviation & " (TradeSense) - Cannot move order because we are waiting for it to go working (" & strEquation & ")"
            End If
        End If
    Else
        DumpDebug "Modifying Order " & m.strAbbreviation & " (TradeSense) - Cannot move order because order ID <=0 (" & strEquation & ")"
        m.bTradeSenseMoved = True
    End If
    
    MoveTradeSense = m.bTradeSenseMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExitOrder.MoveTradeSense"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Attempt to cancel the given order
'' Inputs:      Cancelling All?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelOrder(Optional ByVal bCancellingAll As Boolean = False)
On Error GoTo ErrSection:

    If Not m.Order Is Nothing Then
        DumpDebug "Cancelling Order: " & m.Order.OrderText & " (" & m.Order.GenesisOrderID & ", " & m.Order.BrokerID & ")"
        
        ' To err on the side of safety, make sure that we never confirm a cancel
        ' of an auto exit order with the user even if the confirm orders flag is
        ' turned on.  (7/9/2008 DAJ)...
        mTradeTracker.CancelOrder m.Order, False, , , bCancellingAll
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExitOrder.CancelOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ModifyOrder
'' Description: Attempt to amend the given order
'' Inputs:      Show Can Modify Message?
'' Returns:     True if submitted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ModifyOrder(Optional ByVal bShowCanModifyMsg As Boolean = True) As Boolean
On Error GoTo ErrSection:

    DumpDebug "Modifying Order: " & m.Order.OrderText & " (" & m.Order.GenesisOrderID & ", " & m.Order.BrokerID & ")"
    ModifyOrder = mTradeTracker.EditOrder(m.Order, , False, eGDEditOrderReturn_Submit, bShowCanModifyMsg)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.ModifyOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIsDifferent
'' Description: Determine if the changes will make the order different
'' Inputs:      New Price, New With Limit Price, New Quantity
'' Returns:     True if Different, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderIsDifferent(ByVal dNewPrice As Double, ByVal dNewWithLimitPrice As Double, ByVal lNewQuantity As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If m.Order.Quantity <> lNewQuantity Then
        bReturn = True
    Else
        Select Case m.Order.OrderType
            Case eTT_OrderType_Stop
                bReturn = dNewPrice <> m.Order.StopPrice
            Case eTT_OrderType_Limit
                bReturn = dNewPrice <> m.Order.LimitPrice
            Case eTT_OrderType_StopWithLimit
                bReturn = (dNewPrice <> m.Order.StopPrice) Or (dNewWithLimitPrice <> m.Order.LimitPrice)
        End Select
    End If
    
    OrderIsDifferent = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExitOrder.OrderIsDifferent"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Flattening
'' Description: Are we currently being flattened?
'' Inputs:      None
'' Returns:     True if flattening, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Flattening() As Boolean
On Error GoTo ErrSection:

    Flattening = g.FlattenQueue.IsGettingFlattened(g.Broker.AccountNumberForID(m.lAccountID), m.strSymbol, 0&, eGDFlattenQueueOperation_Flatten)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExitOrder.Flattening"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Send the given debug text to the appropriate place(s)
'' Inputs:      Debug Text
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpDebug(ByVal strText As String)
On Error GoTo ErrSection:

    Dim strToDump As String             ' String to dump to the log
    
    strToDump = "XOS (" & g.Broker.AccountNameForID(AccountID) & ", " & Symbol & "): " & strText

    g.Broker.BrokerDebug m.nAcctType, strToDump
    frmTest2.AddList strToDump

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExitOrder.DumpDebug"
    
End Sub

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

    m.lPrevOrderID = kNullData
    Set m.Order = New cPtOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExitOrder.Class_Initialize"
    
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

    Set m.Order = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExitOrder.Class_Terminate"
    
End Sub
