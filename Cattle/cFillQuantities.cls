VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFillQuantities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cFillQuantities.cls
'' Description: Collection of fill quantities per Account/Symbol/Date/Side/Price
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/13/2014   DAJ         Allow brokers to associate carried fill information
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    FillQuantities As cGdTree           ' Collection of fill quantities
    IdMap As cGdTree                    ' Fill ID map
    AccountInfo As cGdTree              ' Account information
End Type
Private m As mPrivate

Private Property Get Key(ByVal strAccount As String, ByVal strSymbol As String, ByVal strDate As String, ByVal strSide As String, ByVal strPrice As String) As String
    Key = strAccount & "|" & strSymbol & "|" & strDate & "|" & strSide & "|" & strPrice
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Clear
'' Description: Clear out the collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear()
On Error GoTo ErrSection:

    m.FillQuantities.Clear

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFillQuantities.Clear"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Count
'' Description: Number of items in the collection
'' Inputs:      None
'' Returns:     Number of items in the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Count() As Long
On Error GoTo ErrSection:

    Count = m.FillQuantities.Count

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFillQuantities.Count"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    KeyForIndex
'' Description: Return the key for the given index
'' Inputs:      Index
'' Returns:     Key
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function KeyForIndex(ByVal lIndex As Long) As String
On Error GoTo ErrSection:

    KeyForIndex = m.FillQuantities.Key(lIndex)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFillQuantities.KeyForIndex"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFromFill
'' Description: Add more quantity from the fill
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddFromFill(ByVal Fill As cBrokerMessage)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    Dim lQuantity As Long               ' Quantity to add
    Dim lOldQuantity As Long            ' Quantity already there
    Dim AccountInfo As cBrokerMessage   ' Account information object
    
    strKey = KeyForFill(Fill)
    
    If m.IdMap.Exists(Fill("ID")) Then
        lOldQuantity = m.IdMap(Fill("ID"))
        lQuantity = CLng(Val(Fill("Quantity")))
        If lOldQuantity = lQuantity Then
            lQuantity = 0
        Else
            m.IdMap(Fill("ID")) = lQuantity
            lQuantity = lQuantity - lOldQuantity
        End If
    Else
        lQuantity = CLng(Val(Fill("Quantity")))
        m.IdMap.Add lQuantity, Fill("ID")
    End If
    
    If lQuantity <> 0 Then
        If m.AccountInfo.Exists(Fill("BrokerAccountID")) = False Then
            Set AccountInfo = New cBrokerMessage
            
            AccountInfo.Add "BrokerAccountID", Fill("BrokerAccountID")
            AccountInfo.Add "Broker", Fill("Broker")
            AccountInfo.Add "BrokerAccountNumber", Fill("BrokerAccountNumber")
            AccountInfo.Add "FcmAccount", Fill("FcmAccount")
            
            m.AccountInfo.Add AccountInfo, AccountInfo("BrokerAccountID")
        End If
        
        If m.FillQuantities.Exists(strKey) Then
            lOldQuantity = m.FillQuantities(strKey)
            m.FillQuantities(strKey) = lOldQuantity + lQuantity
        Else
            m.FillQuantities.Add lQuantity, strKey
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFillQuantities.AddFromFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillForItem
'' Description: Build a fill for the given item
'' Inputs:      Index
'' Returns:     Fill
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillForItem(ByVal lIndex As Long) As cBrokerMessage
On Error GoTo ErrSection:

    Dim ReturnFill As cBrokerMessage    ' Fill to return from the function
    Dim astrFields As cGdArray          ' Fields in the key
    Dim strKey As String                ' Key for the given index
    Dim AccountInfo As cBrokerMessage   ' Account information
    Dim strDate As String               ' Date
    
    Set ReturnFill = New cBrokerMessage
    strKey = KeyForIndex(lIndex)
    If Len(strKey) > 0 Then
        Set astrFields = New cGdArray
        astrFields.SplitFields strKey, "|"
        
        If m.AccountInfo.Exists(astrFields(0)) Then
            Set AccountInfo = m.AccountInfo(astrFields(0))
            astrFields(0) = AccountInfo("BrokerAccountNumber")
            
            ReturnFill.Add "Broker", AccountInfo("Broker")
            ReturnFill.Add "BrokerAccountID", AccountInfo("BrokerAccountID")
            ReturnFill.Add "BrokerAccountNumber", AccountInfo("BrokerAccountNumber")
            ReturnFill.Add "FcmAccount", AccountInfo("FcmAccount")
        End If
        
        strDate = astrFields(2)
        astrFields(2) = Format(Val(astrFields(2)), "YYYYMMDD")
        
        ReturnFill.Add "BrokerOrderID", astrFields.JoinFields("-")
        ReturnFill.Add "BrokerFillID", ReturnFill("BrokerOrderID")
        ReturnFill.Add "FillTime", strDate
        ReturnFill.Add "Symbol", astrFields(1)
        ReturnFill.Add "IsBuy", astrFields(3)
        ReturnFill.Add "Quantity", m.FillQuantities(lIndex)
        ReturnFill.Add "Price", astrFields(4)
    End If
    
    Set FillForItem = ReturnFill

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFillQuantities.FillForItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Quantity
'' Description: Return the quantity for the given information
'' Inputs:      Account, Symbol, Date, Side, Price
'' Returns:     Quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Quantity(ByVal strAccount As String, ByVal strSymbol As String, ByVal strDate As String, ByVal strSide As String, ByVal strPrice As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim strKey As String                ' Key into the collection
    
    lReturn = 0&
    strKey = Key(strAccount, strSymbol, strDate, strSide, strPrice)
    If m.FillQuantities.Exists(strKey) Then
        lReturn = m.FillQuantities(strKey)
    End If
    
    Quantity = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFillQuantities.Quantity"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    QuantityForFill
'' Description: Return the quantity for information from the fill
'' Inputs:      Fill
'' Returns:     Quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function QuantityForFill(ByVal Fill As cBrokerMessage) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim strKey As String                ' Key into the collection
    
    lReturn = 0&
    strKey = KeyForFill(Fill)
    If m.FillQuantities.Exists(strKey) Then
        lReturn = m.FillQuantities(strKey)
    End If
    
    QuantityForFill = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFillQuantities.QuantityForFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    KeyForFill
'' Description: Build a key for the given fill
'' Inputs:      Fill
'' Returns:     Key
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function KeyForFill(ByVal Fill As cBrokerMessage) As String
On Error GoTo ErrSection:

    Dim strDate As String               ' Date of the fill
    
    strDate = Parse(Fill("FillTime"), ".", 1)
    KeyForFill = Key(Fill("BrokerAccountID"), Fill("Symbol"), strDate, Fill("IsBuy"), Fill("Price"))
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFillQuantities.KeyForFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.FillQuantities = New cGdTree
    Set m.IdMap = New cGdTree
    Set m.AccountInfo = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFillQuantities.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.FillQuantities = Nothing
    Set m.IdMap = Nothing
    Set m.AccountInfo = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFillQuantities.Class_Terminate"
    
End Sub
