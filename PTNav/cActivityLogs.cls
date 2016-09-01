VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cActivityLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cActivityLogs.cls
'' Description: Class to handle a collection of Activity Log objects
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/13/2010   DAJ         Added code for Rithmic
'' 11/01/2010   DAJ         Added Optimus, OpVest, and Vision (Rithmic Brokers)
'' 12/10/2010   DAJ         Moved the ActivityLogSource stuff over to cBrokerDispatch
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    ActivityLogs As cGdTree             ' Collection of activity log objects
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add an activity log object to the collection
'' Inputs:      Form name, Activity Log Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add(ByVal strForm As String, ActivityLog As cActivityLog)
On Error GoTo ErrSection:

    If m.ActivityLogs.Exists(strForm) Then
        Set m.ActivityLogs(strForm) = ActivityLog
    Else
        m.ActivityLogs.Add ActivityLog, strForm
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLogs.Add"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Remove
'' Description: Remove an activity log object from the collection
'' Inputs:      Form name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Remove(ByVal strForm As String)
On Error GoTo ErrSection:

    If m.ActivityLogs.Exists(strForm) Then
        m.ActivityLogs.Remove strForm
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLogs.Remove"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Item
'' Description: Get an activity log object for the given index or key
'' Inputs:      Index or Key
'' Returns:     Activity Log Object (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Item(ByVal vIndexOrKey As Variant) As cActivityLog
On Error GoTo ErrSection:

    If m.ActivityLogs.Exists(vIndexOrKey) Then
        Set Item = m.ActivityLogs(vIndexOrKey)
    Else
        Set Item = Nothing
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cActivityLogs.Item.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddActivityLogEntry
'' Description: Add an activity log entry to each of the grids
'' Inputs:      Broker, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddActivityLogEntry(ByVal nBroker As eTT_AccountType, ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bSaved As Boolean               ' Has the activity log been saved?
    Dim strSource As String             ' Source for the given broker
    
    bSaved = False
    strSource = g.Broker.ActivityLogSource(nBroker)
    
    For lIndex = 1 To m.ActivityLogs.Count
        Item(lIndex).Add strSource, strMessage
        
        ' Have the first activity log object that is not filtered on an account
        ' save the activity log...
        If (bSaved = False) And (Item(lIndex).AccountID = -1&) Then
            Item(lIndex).SaveGrid
            bSaved = True
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLogs.AddActivityLogEntry"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOrderToActivityLog
'' Description: Add an order to the activity log
'' Inputs:      Broker, Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddOrderToActivityLog(ByVal nBroker As eTT_AccountType, ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    If Order.Status = eTT_OrderStatus_Error Then
        If Len(Order.BrokerID) > 0 Then
            AddActivityLogEntry nBroker, "Status for Order #" & Order.BrokerID & ": " & UCase(OrderStatus(Order.Status)) & "  (" & Order.Message & ")"
        Else
            AddActivityLogEntry nBroker, "Status for Order " & Order.OrderText & ": " & UCase(OrderStatus(Order.Status)) & "  (" & Order.Message & ")"
        End If
    Else
        If Len(Order.BrokerID) > 0 Then
            AddActivityLogEntry nBroker, "Status for Order #" & Order.BrokerID & ": " & UCase(OrderStatus(Order.Status)) & "  (" & Order.OrderText & ")"
        Else
            AddActivityLogEntry nBroker, "Status for Order " & Order.OrderText & ": " & UCase(OrderStatus(Order.Status))
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLogs.AddOrderToActivityLog"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFillToActivityLog
'' Description: Add a fill to the activity log
'' Inputs:      Broker, Order, Busted?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddFillToActivityLog(ByVal nBroker As eTT_AccountType, ByVal Fill As cPtFill, ByVal bBusted As Boolean)
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary Bars structure

    SetBarProperties Bars, Fill.SymbolOrSymbolID
    If bBusted = True Then
        If Fill.Buy Then
            AddActivityLogEntry nBroker, "Status for Fill #" & Str(Fill.BrokerID) & ": BUSTED  (BUY " & Str(Fill.Quantity) & " " & Fill.Symbol & " at " & Bars.PriceDisplay(Fill.Price) & ")"
        Else
            AddActivityLogEntry nBroker, "Status for Fill #" & Str(Fill.BrokerID) & ": BUSTED  (SELL " & Str(Fill.Quantity) & " " & Fill.Symbol & " at " & Bars.PriceDisplay(Fill.Price) & ")"
        End If
    Else
        If Fill.Buy Then
            AddActivityLogEntry nBroker, "Status for Fill #" & Str(Fill.BrokerID) & ": FILLED  (BUY " & Str(Fill.Quantity) & " " & Fill.Symbol & " at " & Bars.PriceDisplay(Fill.Price) & ")"
        Else
            AddActivityLogEntry nBroker, "Status for Fill #" & Str(Fill.BrokerID) & ": FILLED  (SELL " & Str(Fill.Quantity) & " " & Fill.Symbol & " at " & Bars.PriceDisplay(Fill.Price) & ")"
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLogs.AddFillToActivityLog"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddStatusToActivityLog
'' Description: Add a connection status message to the activity log
'' Inputs:      Broker, Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddStatusToActivityLog(ByVal nBroker As eTT_AccountType, ByVal nStatus As eGDConnectionStatus)
On Error GoTo ErrSection:

    Dim strSource As String             ' Source name for the given broker
    
    strSource = g.Broker.BrokerName(nBroker)

    Select Case nStatus
        Case eGDConnectionStatus_Disconnected
            AddActivityLogEntry nBroker, "Disconnected from " & strSource & " servers"
        Case eGDConnectionStatus_Disconnecting
            AddActivityLogEntry nBroker, "Disconnecting from " & strSource & " servers"
        Case eGDConnectionStatus_Connecting
            AddActivityLogEntry nBroker, "Connecting to " & strSource & " servers"
        Case eGDConnectionStatus_Connected
            AddActivityLogEntry nBroker, "Connected to " & strSource & " servers"
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLogs.AddStatusToActivityLog"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.ActivityLogs = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLogs.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.ActivityLogs = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLogs.Class_Terminate"
    
End Sub
