VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit

Private Type mPrivate
    Lists As Collection
    alListItemsID As cGdArray
    astrListItemsName As cGdArray
    astrListItems As cGdArray
End Type
Private m As mPrivate

Public Sub Add(ByVal vListID As Variant, List As cList)
    m.Lists.Add List, Str(vListID)
End Sub

Property Get Found(ByVal lID As Long) As Boolean
On Error GoTo ErrSection:

    Dim List As cList
    Found = True
    Set List = Item(Str(lID))

ErrExit:
    Exit Property
    
ErrSection:
    If Err.Number = 91 Or Err.Number = 5 Then
        Found = False
        Resume Next
    Else
        RaiseError "cLists.Found", eGDRaiseError_Raise, g.strAppPath
    End If
    
End Property

'Load all lists into collection
Public Sub Load()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    Dim lIndex As Long
    Dim List As cList
    Dim lListID As Long
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLists] ORDER BY [ListName];", dbOpenSnapshot)
    Do Until rs.EOF
        Set List = New cList
        With List
            lListID = rs!ListID
            .ListID = lListID
            .ListName = rs!ListName
            
            Set m.astrListItems = New cGdArray
            m.astrListItems.Create eGDARRAY_Strings
            For lIndex = 0 To m.alListItemsID.Size - 1
                If m.alListItemsID(lIndex) = lListID Then
                    m.astrListItems.Add m.astrListItemsName(lIndex)
                Else
                    If m.alListItemsID(lIndex) > lListID Then Exit For
                End If
            Next lIndex
            .ListItems = m.astrListItems
        End With
        m.Lists.Add List, Str(rs!ListID)
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Set List = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    Set List = Nothing
    RaiseError "cLists.Load", eGDRaiseError_Raise, g.strAppPath

End Sub

Property Get Item(ByVal vKey As Variant) As cList
    Set Item = m.Lists(vKey)
End Property

Property Get Count() As Long
    Count = m.Lists.Count
End Property

Private Sub Class_Initialize()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    
    Set m.Lists = New Collection
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblListItems] ORDER BY [ListID],[Seq];", dbOpenSnapshot)
    Set m.alListItemsID = New cGdArray
    m.alListItemsID.Create eGDARRAY_Longs
    Set m.astrListItemsName = New cGdArray
    m.astrListItemsName.Create eGDARRAY_Strings
    Do Until rs.EOF
        m.alListItemsID.Add rs!ListID
        m.astrListItemsName.Add rs!ListItem
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cLists.Initialize", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Sub Class_Terminate()
    Set m.Lists = Nothing
End Sub
