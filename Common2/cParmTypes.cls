VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cParmTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit

Private Type mPrivate
    ParmTypes As Collection
End Type
Private m As mPrivate

Property Get Found(ByVal lID As Long) As Boolean
On Error GoTo ErrSection:

    Dim ParmType As cParmType
    Found = True
    Set ParmType = Item(Str(lID))

ErrExit:
    Exit Property
    
ErrSection:
    If Err.Number = 91 Or Err.Number = 5 Then
        Found = False
        Resume Next
    Else
        RaiseError "cParmTypes.Found", eGDRaiseError_Raise, g.strAppPath
    End If

End Property

'Load Collection
Public Sub Load()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    Dim ParmType As cParmType
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblParmTypes] ORDER BY [ParmType];", dbOpenSnapshot)
    Do Until rs.EOF
        Set ParmType = New cParmType
        With ParmType
            .ParmTypeID = rs!ParmTypeID
            .ParmType = rs!ParmType
            .ParmTypeDesc = rs!ParmTypeDesc
            .ParmTypeUsage = rs!ParmUsage
        End With
        m.ParmTypes.Add ParmType, Str(rs!ParmTypeID)
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Set ParmType = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cParmTypes.Load", eGDRaiseError_Raise, g.strAppPath

End Sub

Property Get Item(ByVal vKey As Variant) As cParmType
    Set Item = m.ParmTypes(vKey)
End Property

Property Get Count() As Long
    Count = m.ParmTypes.Count
End Property

Private Sub Class_Initialize()
    Set m.ParmTypes = New Collection
End Sub

Private Sub Class_Terminate()
    Set m.ParmTypes = Nothing
End Sub
