VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOrderLinkWaitList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cOrderLinkWaitList.cls
'' Description: Class to manange an order link wait list
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    WaitListItems As cGdTree            ' Collection of wait list items
End Type
Private m As mPrivate

Public Function Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.WaitListItems.Exists(vIndexOrKey)
End Function

Public Function Remove(ByVal vIndexOrKey As Variant) As Boolean
    Remove = m.WaitListItems.Remove(vIndexOrKey)
End Function

Public Property Get Item(ByVal vIndexOrKey As Variant) As cOrderLinkWaitListItem
On Error GoTo ErrSection:

    Dim ReturnItem As cOrderLinkWaitListItem

    If Exists(vIndexOrKey) Then
        Set ReturnItem = m.WaitListItems.Item(vIndexOrKey)
    Else
        Set ReturnItem = Nothing
    End If
    
    Set Item = ReturnItem

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cOrderLinkWaitList.Item.Get"
    
End Property
Public Property Let Item(ByVal vIndexOrKey As Variant, WaitListItem As cOrderLinkWaitListItem)
On Error GoTo ErrSection:

    If Exists(vIndexOrKey) Then
        Set m.WaitListItems.Item(vIndexOrKey) = WaitListItem
    Else
        Add WaitListItem
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cOrderLinkWaitList.Item.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add a wait list item to the wait list
'' Inputs:      Wait List Item
'' Returns:     Position in the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(WaitListItem As cOrderLinkWaitListItem) As Long
On Error GoTo ErrSection:

    Add = m.WaitListItems.Add(WaitListItem, Str(WaitListItem.OrderLinkID))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLinkWaitList.Add"
    
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

    Set m.WaitListItems = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinkWaitList.Class_Initialize"
    
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

    Set m.WaitListItems = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLinkWaitList.Class_Terminate"
    
End Sub
