VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExitAllOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cExitAllOrders.cls
'' Description: Collection of "exit all" orders
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private WithEvents m_tmrExitAll As Timer
Attribute m_tmrExitAll.VB_VarHelpID = -1

Private Type mPrivate
    ExitAllOrders As cGdTree            ' Collection of "exit all" orders
    bExitAllInProgress As Boolean       ' Is the exit all timer in progress?
End Type
Private m As mPrivate

Public Property Get InProgress() As Boolean
    InProgress = m.bExitAllInProgress
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize some class members
'' Inputs:      Timer
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(tmrExitAll As Timer)
On Error GoTo ErrSection:

    Set m_tmrExitAll = tmrExitAll

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExitAllOrders.Init"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateOrder
'' Description: Add or replace exit all order in the collection
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateOrder(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    If m.ExitAllOrders.Exists(Str(Order.OrderID)) Then
        Set m.ExitAllOrders(Str(Order.OrderID)) = Order
    Else
        m.ExitAllOrders.Add Order, Str(Order.OrderID)
    End If
    m_tmrExitAll.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExitAllOrders.UpdateOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckOrders
'' Description: Check the exit all orders to see if they need to change
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckOrders()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Order As cPtOrder               ' Temporary order object
    Dim lPosition As Long               ' Current position
    Dim lNewQty As Long                 ' New quantity for the order
    Dim AcctPos As cAccountPosition     ' Account position object
    
    m.bExitAllInProgress = True
    
    For lIndex = m.ExitAllOrders.Count To 1 Step -1
        lPosition = 0&
        Set Order = m.ExitAllOrders(lIndex)
        
        If IsOpenOrder(Order.Status) Then
            If OrderIsPending(Order) = False Then
                lPosition = g.Broker.FillSummary(Order.AccountID, Order.SymbolOrSymbolID, Order.AutoTradeItemID).CurrentPositionSnapshot
                
                If Order.Buy Then
                    lNewQty = Order.FillQuantity - lPosition
                Else
                    lNewQty = ((Order.FillQuantity * -1&) - lPosition) * -1&
                End If
                
                If lNewQty <> Order.Quantity Then
                    If NotSent(Order.Status) Then
                        Order.Quantity = lNewQty
                        Order.Save
                        OrderCallback Order
                    ElseIf lNewQty > 0 Then
                        '*** PRICE OR QUANTITY ON ORDER CHANGED...***
                        Order.Quantity = lNewQty
                        g.Broker.BrokerDebug Order.Broker, "Modifying Exit All Order from Trade Console: " & Order.OrderText, True
                        EditOrder Order, , False, eGDEditOrderReturn_Submit
                        'Order.Save
                        'SubmitOrder Order, Order.GenesisOrderID
                    Else
                        g.Broker.BrokerDebug Order.Broker, "Cancelling Exit All Order from Trade Console: " & Order.OrderText, True
                        CancelOrder Order, False
                    End If
                End If
            End If
        Else
            m.ExitAllOrders.Remove lIndex
        End If
    Next lIndex
    
    If m.ExitAllOrders.Count = 0 Then
        m_tmrExitAll.Enabled = False
    End If
    
    m.bExitAllInProgress = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExitAllOrders.CheckOrders"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_tmrExitAll_Timer
'' Description: Modify "exit all" orders as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_tmrExitAll_Timer()
On Error GoTo ErrSection:

    CheckOrders

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExitAllOrders.m_tmrExitAll_Timer"
    
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

    Set m.ExitAllOrders = New cGdTree
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExitAllOrders.Class_Initialize"
    
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

    Set m.ExitAllOrders = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExitAllOrders.Class_Terminate"
    
End Sub

