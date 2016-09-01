VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cChartDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IVSFlexDataSource

Private Type mPrivate
    nPrevX As Long
    aFields() As String
End Type
Private m As mPrivate

Private Sub Class_Initialize()
    ReDim aFields(0) As String
End Sub

Private Function IVSFlexDataSource_GetData(ByVal Field As Long, ByVal Record As Long) As String
On Error GoTo ErrSection:

    Dim strText$
    
    If Record <> m.nPrevX Then
        Call GetChartData(Record)
    End If
    If Field < UBound(m.aFields) Then
        strText = Parse(m.aFields(Field + 1), vbTab, 2)
    End If
    
    IVSFlexDataSource_GetData = strText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChartDataGrid.IVSFlexDataSource.GetData", eGDRaiseError_Raise
    
End Function

Private Function IVSFlexDataSource_GetFieldCount() As Long
On Error GoTo ErrSection:

    Dim nCount&
    
    If GetChartData(-1) Then
        nCount = UBound(m.aFields)
    End If
    
    IVSFlexDataSource_GetFieldCount = nCount

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChartDataGrid.IVSFlexDataSource.GetFieldCount", eGDRaiseError_Raise
    
End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal Field As Long) As String
On Error GoTo ErrSection:

    Dim strText$
    
    If Field < UBound(m.aFields) Then
        strText = Parse(m.aFields(Field + 1), vbTab, 1)
    End If

    IVSFlexDataSource_GetFieldName = strText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChartDataGrid.IVSFlexDataSource.GetFieldName", eGDRaiseError_Raise
    
End Function

Private Function IVSFlexDataSource_GetRecordCount() As Long
On Error GoTo ErrSection:

    Dim nCount&, frm As Form
    
    Set frm = ActiveChart
    If Not frm Is Nothing Then
        nCount = frm.Chart.aXBar.Size
    End If
    
    IVSFlexDataSource_GetRecordCount = nCount

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChartDataGrid.IVSFlexDataSource.GetRecordCount", eGDRaiseError_Raise
    
End Function

Private Sub IVSFlexDataSource_SetData(ByVal Field As Long, ByVal Record As Long, ByVal newData As String)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChartDataGrid.IVSFlexDataSource.SetData", eGDRaiseError_Raise
    
End Sub

Private Function GetChartData(ByVal nX&) As Boolean
On Error GoTo ErrSection:

    Dim strData$, frm As Form
    
    Set frm = ActiveChart
    If Not frm Is Nothing Then
        strData = frm.Chart.GetDataWindowLabel(nX)
    End If
    
    If Len(strData) > 0 Then
        m.aFields = Split(strData, "|")
        m.nPrevX = nX
        GetChartData = True
    Else
        ReDim m.aFields(0) As String
        m.nPrevX = -2
        GetChartData = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChartDataGrid.GetChartData", eGDRaiseError_Raise
    
End Function

