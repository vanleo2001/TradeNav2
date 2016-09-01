VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type mPrivate
    lListID As Long
    strListName As String
    astrListItems As cGdArray
End Type
Private m As mPrivate

Property Get ListID() As Long
    ListID = m.lListID
End Property
Property Get ListName() As String
    ListName = m.strListName
End Property
Property Get ListItems() As Object
    Set ListItems = m.astrListItems
End Property
Property Get ListItemsHandle() As Long
    ListItemsHandle = m.astrListItems.ArrayHandle
End Property

Property Let ListID(ByVal pData As Long)
    m.lListID = pData
End Property
Property Let ListName(ByVal pData As String)
    m.strListName = pData
End Property
Property Let ListItems(pData As Object)
    Set m.astrListItems = pData
End Property
Property Let ListItemsHandle(ByVal hArray As Long)
    m.astrListItems.CopyFromHandle hArray
End Property

Private Sub Class_Initialize()
    Set m.astrListItems = New cGdArray
    m.astrListItems.Create eGDARRAY_Strings
End Sub

Private Sub Class_Terminate()
    Set m.astrListItems = Nothing
End Sub
