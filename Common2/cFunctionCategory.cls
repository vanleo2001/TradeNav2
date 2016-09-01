VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type mPrivate
    lFunctionCategoryID As Long
    strFunctionCategory As String
    bCannotDelete As Boolean
End Type
Private m As mPrivate

Property Get FunctionCategoryID() As Long
    FunctionCategoryID = m.lFunctionCategoryID
End Property
Property Get FunctionCategory() As String
    FunctionCategory = m.strFunctionCategory
End Property
Property Get CannotDelete() As Boolean
    CannotDelete = m.bCannotDelete
End Property

Property Let FunctionCategoryID(ByVal pData As Long)
    m.lFunctionCategoryID = pData
End Property
Property Let FunctionCategory(pData As String)
    m.strFunctionCategory = pData
End Property
Property Let CannotDelete(pData As Boolean)
    m.bCannotDelete = pData
End Property


