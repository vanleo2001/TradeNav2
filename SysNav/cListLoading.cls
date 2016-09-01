VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'This classes loops through each list object and populates the list items
'vector for internal lists (system names, signal names, etc.)

Private Type mPrivate
    Lists As cLists
End Type
Private m As mPrivate

Property Get Lists() As cLists
    Set Lists = m.Lists
End Property

Public Sub Load()
On Error GoTo ErrSection:
    
    Dim X       As Long
    'Dim Y       As Long
    Dim rs As Recordset
    Dim strTemp As String
    
    Set m.Lists = New cLists
    m.Lists.Load
    
    For X = 1 To m.Lists.Count
        With m.Lists.Item(X)
            Select Case .ListName
                    
                Case "Symbol Names"
                    'For Y = 1 To g.Markets.Count
                    '    .ListItems.Add g.Markets.Item(Y).MarketSymbol
                    'Next Y
                    .ListItems.Add "SP"
                    .ListItems.Add "TQ"
                    .ListItems.Add "CR"
                    .ListItems.Add "GC"
            
                    .ListItems.Add "IBM"
                    .ListItems.Add "MSFT"
                    
                    strTemp = ",Market1,Daily,Weekly,Monthly,Unsplit,SP,TQ,CR,GC,IBM,MSFT,"
                    Set rs = g.dbNav.OpenRecordset("SELECT tblRuleParms.ParmName " & _
                            "FROM tblRuleParms INNER JOIN tblSystemSecurities ON tblRuleParms.ParmID = tblSystemSecurities.ParmID " & _
                            "GROUP BY tblRuleParms.ParmName;", dbOpenDynaset)
                    Do While Not rs.EOF
                        If InStr(UCase(strTemp), "," & UCase(rs!ParmName) & ",") = 0 Then
                            .ListItems.Add rs!ParmName
                            strTemp = strTemp & rs!ParmName & ","
                        End If
                        rs.MoveNext
                    Loop
                    .ListItems.Sort
                    
                    .ListItems.Add "Daily", 0
                    .ListItems.Add "Weekly", 1
                    .ListItems.Add "Monthly", 2
                    '.ListItems.Add "Unsplit", 3
            End Select
        End With
    Next X
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cListLoading.Load", eGDRaiseError_Raise

End Sub

Private Sub Class_Terminate()
    Set m.Lists = Nothing
End Sub
