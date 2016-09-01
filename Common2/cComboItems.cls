VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComboItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
'This collection loads a two column lookup table.  The first field is always
'a counter (key), and the second a description field.  For example,  The
'table tblTradeDurations is loaded into this collection.  It looks as
'follows:
'
'   1  Long Term
'   2  Short Term
'   3  Medium Term
'
Private Type mPrivate
    ComboItems As Collection
    dbNav As Database
End Type
Private m As mPrivate

Property Let dbNav(pData As Database)
    Set m.dbNav = pData
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the collection of combo items from the database
'' Inputs:      Name of the Table, Name of the Sort Field
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load(ByVal strTableName As String, ByVal strSortFieldName As String)
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim TmpItem As cComboItem           ' Temporary combo item
    
    Set rs = m.dbNav.OpenRecordset("SELECT * FROM [" & strTableName & "] " & _
        "ORDER BY [" & strSortFieldName & "];", dbOpenSnapshot)
    Do Until rs.EOF
        Set TmpItem = New cComboItem
        With TmpItem
            .ComboItemID = rs.Fields(0).Value
            .ComboItem = rs.Fields(1).Value
            m.ComboItems.Add TmpItem, rs.Fields(1).Value
        End With
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Set TmpItem = Nothing
    Exit Sub

ErrSection:
    Set rs = Nothing
    Set TmpItem = Nothing
    RaiseError "cComboItems.Load", eGDRaiseError_Raise, g.strAppPath

End Sub

Property Get Item(ByVal vKey As Variant) As cComboItem
    Set Item = m.ComboItems(vKey)
End Property

Property Get Found(ByVal vKey As Variant) As Boolean
On Error GoTo ErrSection:

    Dim ComboItem As cComboItem
    Found = True
    Set ComboItem = Item(vKey)

ErrExit:
    Exit Property

ErrSection:
    If Err.Number = 91 Or Err.Number = 5 Then
        Found = False
        Resume Next
    Else
        RaiseError "cComboItems.Found", eGDRaiseError_Raise, g.strAppPath
    End If

End Property

Property Get Count() As Long
    Count = m.ComboItems.Count
End Property

Private Sub Class_Initialize()
    Set m.ComboItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set m.ComboItems = Nothing
End Sub
