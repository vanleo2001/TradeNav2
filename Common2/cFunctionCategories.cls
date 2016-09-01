VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit

Private Type mPrivate
    FunctionCategories As Collection
End Type
Private m As mPrivate

Property Get Found(ByVal lID As Long) As Boolean
On Error GoTo ErrSection:

    Dim FunctionCategory As cFunctionCategory
    Found = True
    Set FunctionCategory = Item(Str(lID))

ErrExit:
    Exit Property
    
ErrSection:
    If Err.Number = 91 Or Err.Number = 5 Then
        Found = False
        Resume Next
    Else
        RaiseError "cFunctionCategories.Found", eGDRaiseError_Raise, g.strAppPath
    End If

End Property

Public Sub Load()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    Dim FunctionCategory As cFunctionCategory
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionCategories] " & _
                "ORDER BY [FunctionCategory];", dbOpenSnapshot)
    Do Until rs.EOF
        Set FunctionCategory = New cFunctionCategory
        With FunctionCategory
            .FunctionCategoryID = rs!FunctionCategoryID
            .FunctionCategory = rs!FunctionCategory
            .CannotDelete = rs!CannotDelete
        End With
        m.FunctionCategories.Add FunctionCategory, Str(rs!FunctionCategoryID)
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Set FunctionCategory = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    Set FunctionCategory = Nothing
    RaiseError "cFunctionCategories.Load", eGDRaiseError_Raise, g.strAppPath

End Sub

Property Get Item(ByVal vKey As Variant) As cFunctionCategory
    Set Item = m.FunctionCategories(vKey)
End Property

Property Get Count() As Long
    Count = m.FunctionCategories.Count
End Property

Private Sub Class_Initialize()
    Set m.FunctionCategories = New Collection
End Sub

Private Sub Class_Terminate()
    Set m.FunctionCategories = Nothing
End Sub

