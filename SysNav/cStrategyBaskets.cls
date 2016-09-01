VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStrategyBaskets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cStrategyBaskets.cls
'' Description: Collection of strategy baskets
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 04/03/2013   DAJ         Move Strategy Baskets into the database
'' 05/01/2013   DAJ         Shadow Trading
'' 05/14/2013   DAJ         Optionally allow loading a guru basket if not the owner
'' 08/05/2013   DAJ         Don't allow user to save strategy basket with an existing name
'' 08/05/2013   DAJ         Ignore case on existing strategy basket name check
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Baskets As cGdTree                  ' Collection of strategy baskets
End Type
Private m As mPrivate

Public Property Get Count() As Long
    Count = m.Baskets.Count
End Property

Public Property Get Item(ByVal vIndexOrKey As Variant) As cStrategyBasket
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrSection:

    Dim Basket As cStrategyBasket       ' Basket to return
    
    If m.Baskets.Exists(vIndexOrKey) Then
        Set Basket = m.Baskets(vIndexOrKey)
    Else
        Set Basket = Nothing
    End If
    
    Set Item = Basket

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cStrategyBasket.Item.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDb
'' Description: Load all the strategy baskets from the database
'' Inputs:      Include Hidden Strategy Basket if IDE?, Include even if not the owner?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadDb(Optional ByVal bIncludeHiddenIfIde As Boolean = True, Optional ByVal bIncludeGuruIfNotOwner As Boolean = False, _
        Optional ByVal bLoadBasketItems As Boolean = True)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Basket As cStrategyBasket       ' Strategy basket item

    Set rs = mSysNav.LoadStrategyBasketsRecordset
    Do While Not rs.EOF
        If mSysNav.IncludeStrategyBasketsFromRecordset(rs, bIncludeHiddenIfIde, bIncludeGuruIfNotOwner) Then
            Set Basket = New cStrategyBasket
            If Basket.LoadDbFromRs(rs, bIncludeHiddenIfIde, bIncludeGuruIfNotOwner, bLoadBasketItems) Then
                m.Baskets.Add Basket, "ID=" & Str(Basket.ID)
            End If
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBaskets.LoadDb"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NameExists
'' Description: Does the given name already exist on a strategy basket?
'' Inputs:      Name, Ignore ID
'' Returns:     True if Exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NameExists(ByVal strName As String, Optional ByVal lIgnoreID As Long = 0&, Optional ByVal bIgnoreCase As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Basket As cStrategyBasket       ' Strategy basket object
    
    bReturn = False
    For lIndex = 1 To m.Baskets.Count
        Set Basket = Item(lIndex)
        If Basket.ID <> lIgnoreID Then
            If bIgnoreCase Then
                If UCase(Basket.Name) = UCase(strName) Then
                    bReturn = True
                    Exit For
                End If
            Else
                If Basket.Name = strName Then
                    bReturn = True
                    Exit For
                End If
            End If
        End If
    Next lIndex
    
    NameExists = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBaskets.NameExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Baskets = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBaskets.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Baskets = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBaskets.Class_Terminate"
    
End Sub
