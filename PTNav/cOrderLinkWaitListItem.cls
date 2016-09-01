VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOrderLinkWaitListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cOrderLinkWaitListItem.cls
'' Description: Class to manange an order link wait list item
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lOrderLinkID As Long                ' Unique identifier of the order link
    OriginalOrder As cPtOrder           ' Original Order
    NewOrder As cPtOrder                ' New order in the case of an amend
    OriginalOrder2 As cPtOrder          ' Original Order for the second order
    NewOrder2 As cPtOrder               ' New Order for the second order
    nCommand As eGDWaitListCommands     ' Command to perform
End Type
Private m As mPrivate

Public Property Get OrderLinkID() As Long
    OrderLinkID = m.lOrderLinkID
End Property
Public Property Let OrderLinkID(ByVal lOrderLinkID As Long)
    m.lOrderLinkID = lOrderLinkID
End Property

Public Property Get OriginalOrder() As cPtOrder
    Set OriginalOrder = m.OriginalOrder
End Property
Public Property Let OriginalOrder(ByVal Order As cPtOrder)
    Set m.OriginalOrder = Order
End Property

Public Property Get NewOrder() As cPtOrder
    Set NewOrder = m.NewOrder
End Property
Public Property Let NewOrder(ByVal Order As cPtOrder)
    Set m.NewOrder = Order
End Property

Public Property Get OriginalOrder2() As cPtOrder
    Set OriginalOrder2 = m.OriginalOrder2
End Property
Public Property Let OriginalOrder2(ByVal Order As cPtOrder)
    Set m.OriginalOrder2 = Order
End Property

Public Property Get NewOrder2() As cPtOrder
    Set NewOrder2 = m.NewOrder2
End Property
Public Property Let NewOrder2(ByVal Order As cPtOrder)
    Set m.NewOrder2 = Order
End Property

Public Property Get Command() As eGDWaitListCommands
    Command = m.nCommand
End Property
Public Property Let Command(ByVal nCommand As eGDWaitListCommands)
    m.nCommand = nCommand
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Create a string for the wait list item
'' Inputs:      None
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    strReturn = Str(OrderLinkID)
    
    If OriginalOrder Is Nothing Then
        strReturn = strReturn & ";"
    Else
        strReturn = strReturn & ";" & Str(OriginalOrder.OrderID)
    End If
    
    If NewOrder Is Nothing Then
        strReturn = strReturn & ";"
    Else
        strReturn = strReturn & ";" & Str(NewOrder.OrderID)
    End If
    
    strReturn = strReturn & ";" & WaitListCommandString(Command)
    
    If OriginalOrder2 Is Nothing Then
        strReturn = strReturn & ";"
    Else
        strReturn = strReturn & ";" & Str(OriginalOrder2.OrderID)
    End If
    
    If NewOrder2 Is Nothing Then
        strReturn = strReturn & ";"
    Else
        strReturn = strReturn & ";" & Str(NewOrder2.OrderID)
    End If
    
    ToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinkWaitListItem.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.OriginalOrder = Nothing
    Set m.NewOrder = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinkWaitListItem.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.OriginalOrder = Nothing
    Set m.NewOrder = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinkWaitListItem.Class_Terminate"
    
End Sub
