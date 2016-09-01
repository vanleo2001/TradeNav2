VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSystemsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    Systems As cGdTable
End Type
Private m As mPrivate

'Properties for manipulating Trade detail rows and columns
Property Get FieldNum(ByVal pData As ensy_Systems) As Long
    FieldNum = pData
End Property
Property Get FieldName(ByVal pData As ensy_Systems) As String
    FieldName = m.Systems.FieldName(pData)
End Property
Property Get FieldType(ByVal pData As ensy_Systems) As eGdArray_Type
    FieldType = m.Systems.FieldType(pData)
End Property
Property Get Num(ByVal pRecordNbr, ByVal pField As ensy_Systems) As Double
    Num = m.Systems.Num(pField, pRecordNbr)
End Property
Property Get Item(ByVal pRecordNbr, ByVal pField As ensy_Systems) As Variant
    Item = m.Systems.Item(pField, pRecordNbr)
End Property
Property Get FieldHandle(ByVal pField As ensy_Systems) As Long
    FieldHandle = m.Systems.FieldArrayHandle(pField)
End Property
Property Get NumRecords() As Long
    NumRecords = m.Systems.NumRecords
End Property
Property Get TableHandle() As Long
    If m.Systems Is Nothing Then
        TableHandle = 0
    Else
        TableHandle = m.Systems.TableHandle
    End If
End Property

Property Let Num(ByVal pRecordNbr, ByVal pField As ensy_Systems, ByVal pValue As Double)
    m.Systems.Num(pField, pRecordNbr) = pValue
End Property
Property Let Item(ByVal pRecordNbr, ByVal pField As ensy_Systems, ByVal pValue As Variant)
    m.Systems.Item(pField, pRecordNbr) = pValue
End Property
Property Let NumRecords(ByVal lNumRecords As Long)
    m.Systems.NumRecords = lNumRecords
End Property

Public Function IsNullAt(ByVal pRecordNbr, ByVal pField As ensy_Systems) As Boolean
    IsNullAt = m.Systems.IsNullAt(pField, pRecordNbr)
End Function

Public Sub SetSize(ByVal lSize As Long)
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    
    For lIndex = 1 To ensy_Cols
        gdReserve FieldHandle(lIndex), lSize, 0
        gdSetSize FieldHandle(lIndex), lSize, 0
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystemsList.SetSize", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Function CreateIndex(Optional ByVal nFilterByField As ensy_Systems = -1) As cGdArray
    Set CreateIndex = m.Systems.CreateIndex(nFilterByField)
End Function

Public Function SortIndex(ByVal aIndex As cGdArray, ByVal nSortByField As ensy_Systems, Optional ByVal eSortFlags As eGdSort_Flags = eGdSort_Default) As Long
    SortIndex = m.Systems.SortIndex(aIndex, nSortByField, eSortFlags)
End Function

Public Sub CopyTableFromHandle(ByVal lNewTableHandle As Long)
    m.Systems.CopyFromHandle lNewTableHandle
End Sub

Public Function Serialize(ByVal strFile As String, ByVal bPut As Boolean) As Boolean
On Error GoTo ErrSection:

    Serialize = True
    If Not m.Systems.Serialize(strFile, bPut) Then
        Serialize = False
    End If
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cSystemsList.Serialize", eGDRaiseError_Raise, g.strAppPath

End Function

Private Sub Class_Initialize()
On Error GoTo ErrSection:
    
    Dim Sys         As ensy_Systems
    
    Set m.Systems = New cGdTable
    Sys = ensy_SystemNumber
    m.Systems.CreateField eGDARRAY_Longs, Sys, "SystemNumber"
    Sys = ensy_SystemName
    m.Systems.CreateField eGDARRAY_Strings, Sys, "SystemName"
    Sys = ensy_Symbol
    m.Systems.CreateField eGDARRAY_Strings, Sys, "Symbol"
    Sys = ensy_DefaultUnits
    m.Systems.CreateField eGDARRAY_Floats, Sys, "DefaultUnits"
    Sys = ensy_TickMove
    m.Systems.CreateField eGDARRAY_Doubles, Sys, "TickMove"
    Sys = ensy_TickValue
    m.Systems.CreateField eGDARRAY_Doubles, Sys, "TickValue"
    Sys = ensy_MinMoveInTicks
    m.Systems.CreateField eGDARRAY_Doubles, Sys, "MinMoveInTicks"
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSystemsList.Initialize", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub Class_Terminate()
    
    Set m.Systems = Nothing

End Sub
