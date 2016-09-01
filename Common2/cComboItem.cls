VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComboItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type mPrivate
    lComboItemID As Long
    strComboItem As String
End Type
Private m As mPrivate

Property Get ComboItemID() As Long
    ComboItem = m.lComboItemID
End Property
Property Get ComboItem() As String
    ComboItem = m.strComboItem
End Property

Property Let ComboItemID(pData As Long)
    m.lComboItemID = pData
End Property
Property Let ComboItem(pData As String)
    m.strComboItem = pData
End Property

