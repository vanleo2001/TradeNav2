VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMarkets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit

Private Type mPrivate
    Markets As Collection
End Type
Private m As mPrivate

'Wrapper around Item property to see if it exists
Property Get Found(ByVal lID As Long) As Boolean
On Error GoTo ErrSection:

    Dim Market As cMarket
    Found = True
    Set Market = Item(Str(lID))

ErrExit:
    Exit Property
ErrSection:
    If Err.Number = 91 Or Err.Number = 5 Then
        Found = False
        Resume Next
    Else
        RaiseError "cMarkets.Found", eGDRaiseError_Raise, g.strAppPath
    End If
    
End Property

'Get ID based on desription
Public Sub GetID(ByVal strName As String, lID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Integer
    
    lID = 0
    For lIndex = 1 To Count
        If strName = Item(lIndex).MarketSymbol Then
            lID = Item(lIndex).MarketID
            Exit For
        End If
    Next lIndex
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cMarkets.GetID", eGDRaiseError_Raise, g.strAppPath

End Sub

'Load Collection
Public Sub Load()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    Dim Market As cMarket
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblMarkets] ORDER BY [MarketSymbol];", dbOpenSnapshot)
    Do Until rs.EOF
        Set Market = New cMarket
        With Market
            .MarketID = rs!MarketID
            .MarketSymbol = rs!MarketSymbol
            .TickMove = Val(rs!TickMove)
            .TickValue = rs!TickValue
            .MinMoveInTicks = rs!MinMoveInTicks
            .MarketDesc = rs!MarketDesc
            .Margin = rs!Margin
            .ContractSize = rs!ContractSize
        End With
        m.Markets.Add Market, Str(rs!MarketID)
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set rs = Nothing
    RaiseError "cMarkets.Load", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'Save the current item...
Public Sub Save(ByVal lMarketID As Long, ByVal dMargin As Currency)
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    
    Validate dMargin
    
    ' Update the table first...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblMarkets] " & _
            "WHERE [MarketID]=" & Str(lMarketID) & ";", dbOpenDynaset)
    If Not rs.EOF Then
        rs.Edit
        rs!Margin = dMargin
        rs.Update
    End If
    
    ' Update existing collection item...
    With Item(Str(lMarketID))
        .Margin = dMargin
    End With
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cMarkets.Save", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Sub Validate(ByVal dMargin As Currency)
On Error GoTo ErrSection:
    
    If dMargin < 0 Or dMargin > 10000000 Then
        Err.Raise vbObjectError + 1001, , "Margin must be between 0 and 10,000,000"
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cMarkets.Validate", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Property Get Item(ByVal vKey As Variant) As cMarket
    Set Item = m.Markets(vKey)
End Property

Property Get Count() As Long
    Count = m.Markets.Count
End Property

Private Sub Class_Initialize()
    Set m.Markets = New Collection
End Sub

Private Sub Class_Terminate()
    Set m.Markets = Nothing
End Sub
