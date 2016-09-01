VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionsByName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Functions collection indexed by function name...
Option Explicit
Option Compare Text

Private Type mPrivate
    FunctionsByName As Collection
End Type
Private m As mPrivate

'Special Property used to determine if a given function is late calculating
'or not.  We have to remove the hardcoding eventually.
Property Get Late(ByVal lFunctionID As Long) As Boolean
    
    Select Case lFunctionID
        Case 122, 123, 124
            Late = True
            
        Case Else
            Late = False
    
    End Select

End Property

'Load functions from tblFunctions into cFunctionsByName collection...
Public Sub Load()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    Dim rsParms As Recordset
    Dim TmpFunc As cFunction
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
                    "ORDER BY [FunctionName];", dbOpenDynaset)
    ValidateCheckSums rs, "tblFunctions"
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do Until rs.EOF
        If rs!CheckSum <> 0.5 Then
            Set TmpFunc = New cFunction
            With TmpFunc
                'Set properties of function...
                .FunctionID = rs!FunctionID
                .FunctionName = rs!FunctionName
                .Description = rs!Description
                .FunctionCategoryID = rs!FunctionCategoryID
                .ImplementationTypeID = rs!ImplementationTypeID
                .CodedName = rs!CodedName
                .CodedText = DecryptField(rs!CodedText)
                .Usage = rs!Usage
                .DataTypeID = rs!ReturnTypeID
                .LibraryID = rs!LibraryID
                
                m.FunctionsByName.Add TmpFunc, Str(.FunctionName)
            End With
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    
ErrExit:
    Set TmpFunc = Nothing
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cFunctionsByName.Load", eGDRaiseError_Raise, g.strAppPath

End Sub

'Adds a system to the collection
Public Sub Add(ByVal strFunctionName As String, pFunction As cFunction)
    m.FunctionsByName.Add pFunction, Str(strFunctionName)
End Sub

Property Get Found(ByVal vKey As Variant) As Boolean
On Error GoTo ErrSection:
    
    Dim TmpFunc As cFunction
    Found = True
    Set TmpFunc = m.FunctionsByName(vKey)

ErrExit:
    Exit Property
    
ErrSection:
    If Err.Number = 91 Or Err.Number = 5 Then
        Found = False
        Resume Next
    Else
        RaiseError "cFunctionsByName.Found", eGDRaiseError_Raise, g.strAppPath
    End If
    
End Property

Property Get Item(ByVal vKey As Variant) As cFunction
    Set Item = m.FunctionsByName(vKey)
End Property

'Remove item from collection and pFunctionID
Public Sub Remove(ByVal lFunctionID As Long)
On Error GoTo ErrSection:
    
    Dim rs As Recordset

    m.FunctionsByName.Remove Str(lFunctionID)
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
            "WHERE [FunctionID]=" & Str(lFunctionID) & ";", dbOpenDynaset)
    If rs.EOF Then
        Err.Raise vbObjectError + 1000, "Delete cancelled because Function does not exist"
    Else
        rs.Delete
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cFunctionsByName.Remove", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Property Get Count() As Long
    Count = m.FunctionsByName.Count
End Property

Private Sub Class_Initialize()
    Set m.FunctionsByName = New Collection
End Sub

Private Sub Class_Terminate()
    Set m.FunctionsByName = Nothing
End Sub
