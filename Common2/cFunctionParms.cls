VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionParms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private Type mPrivate
    FunctionParms As Collection
End Type
Private m As mPrivate

Public Sub Add(pParm As cFunctionParm, ByVal lParmID As Long)
    m.FunctionParms.Add pParm, Str(lParmID)
End Sub

Public Sub Remove(ByVal lParmID As Long)
    m.FunctionParms.Remove Str(lParmID)
End Sub

Property Get Item(ByVal vKey As Variant) As cFunctionParm
    Set Item = m.FunctionParms(vKey)
End Property

Property Get Count() As Long
    Count = m.FunctionParms.Count
End Property

Private Sub Class_Initialize()
    Set m.FunctionParms = New Collection
End Sub

Private Sub Class_Terminate()
    Set m.FunctionParms = Nothing
End Sub
