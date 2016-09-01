VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    bHide As Boolean
    strFieldName As String
    strFieldDesc As String
    lFieldNumber As Long
    strFieldFormat As String
    strFieldOrder As String
    strOperator As String
    dFilterValue As Double
    bOptimized As Boolean
End Type
Private m As mPrivate

Property Get Hide() As Boolean
    Hide = m.bHide
End Property
Property Get FieldName() As String
    FieldName = m.strFieldName
End Property
Property Get FieldDesc() As String
    FieldDesc = m.strFieldDesc
End Property
Property Get FieldNumber() As Long
    FieldNumber = m.lFieldNumber
End Property
Property Get FieldFormat() As String
    FieldFormat = m.strFieldFormat
End Property
Property Get FieldOrder() As String
    FieldOrder = m.strFieldOrder
End Property
Property Get Operator() As String
    Operator = m.strOperator
End Property
Property Get FilterValue() As Double
    FilterValue = m.dFilterValue
End Property
Property Get Optimized() As Boolean
    Optimized = m.bOptimized
End Property

Property Let Hide(pData As Boolean)
    m.bHide = pData
End Property
Property Let FieldName(pData As String)
    m.strFieldName = pData
End Property
Property Let FieldDesc(pData As String)
    m.strFieldDesc = pData
End Property
Property Let FieldNumber(pData As Long)
    m.lFieldNumber = pData
End Property
Property Let FieldFormat(pData As String)
    m.strFieldFormat = pData
End Property
Property Let FieldOrder(pData As String)
    m.strFieldOrder = pData
End Property
Property Let Operator(pData As String)
    m.strOperator = pData
End Property
Property Let FilterValue(pData As Double)
    m.dFilterValue = pData
End Property
Property Let Optimized(pData As Boolean)
    m.bOptimized = pData
End Property

