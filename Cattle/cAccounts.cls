VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAccounts.cls
'' Description: Class that holds information about a collection of accounts
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Accounts As cGdTree                 ' Collection of accounts
End Type
Private m As mPrivate

Public Property Get Count() As Long
    Count = m.Accounts.Count
End Property

Public Property Get Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.Accounts.Exists(vIndexOrKey)
End Property

Public Property Get Item(ByVal vIndexOrKey As Variant) As cAccount
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrSection:

    Dim Account As cAccount             ' Account object to return
    
    If Exists(vIndexOrKey) Then
        Set Account = m.Accounts(vIndexOrKey)
    Else
        Set Account = Nothing
    End If
    
    Set Item = Account

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAccounts.Item.Get"
    
End Property
Public Property Let Item(ByVal vIndexOrKey As Variant, Account As cAccount)
On Error GoTo ErrSection:

    If Exists(vIndexOrKey) Then
        Set m.Accounts(vIndexOrKey) = Account
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAccounts.Item.Let"
    
End Property

Public Property Get Key(ByVal Account As cAccount) As String
    Key = Str(Account.ID)
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

    m.Accounts.Clear
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccounts.Clear"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add the given account to the collection
'' Inputs:      Account
'' Returns:     Index into the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(Account As cAccount) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim strKey As String                ' Key into the collection

    strKey = Key(Account)
    If Exists(strKey) Then
        Set m.Accounts(strKey) = Account
        lReturn = m.Accounts.Index(strKey)
    Else
        lReturn = m.Accounts.Add(Account, strKey)
    End If
    
    Add = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccounts.Add"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FcmNumberForId
'' Description: Get the FCM Number for the given account ID
'' Inputs:      Account ID, Value if Blank
'' Returns:     FCM Number ( Blank if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FcmNumberForId(ByVal strAccountID As String, Optional ByVal strValueIfBlank As String = "") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Exists(strAccountID) Then
        strReturn = Item(strAccountID).FcmNumber
    End If
    
    If (Len(strReturn) = 0) Then
        If (Len(strValueIfBlank) > 0) Then
            strReturn = strValueIfBlank
        End If
    End If
    
    FcmNumberForId = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccounts.FcmNumberForId"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountIdForBrokerNumber
'' Description: Get the account ID for the given broker account number
'' Inputs:      Broker Account Number, Broker
'' Returns:     Account ID (Blank if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountIdForBrokerNumber(ByVal strBrokerAccountNumber As String, ByVal nBroker As eTT_AccountType) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strReturn As String             ' Return value for the function
    Dim Account As cAccount             ' Account object
    
    strReturn = ""
    For lIndex = 1 To Count
        Set Account = Item(lIndex)
        If Account.AccountNumber = strBrokerAccountNumber Then
            If Account.Broker = nBroker Then
                strReturn = Account.ID
                Exit For
            End If
        End If
    Next lIndex
    
    AccountIdForBrokerNumber = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccounts.AccountIdForBrokerNumber"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayAccountNumber
'' Description: Determine the account number to display for the given ID
'' Inputs:      Account ID
'' Returns:     Account Number to Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DisplayAccountNumber(ByVal strAccountID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim Account As cAccount             ' Account
    
    strReturn = ""
    If Exists(strAccountID) Then
        Set Account = Item(strAccountID)
        If Len(Account.FcmNumber) = 0 Then
            strReturn = Account.AccountNumber
        Else
            strReturn = Account.FcmNumber
        End If
    End If
    
    DisplayAccountNumber = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccounts.DisplayAccountNumber"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAccountsCombo
'' Description: Load the accounts combo with the associated accounts
'' Inputs:      Combo Box, Selected Account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadAccountsCombo(cboAccounts As ComboBox, Optional ByVal strSelectedAccountID As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Account As cAccount             ' Account object
    Dim lListIndex As Long              ' List index to select
    Dim strBrokerName As String         ' Broker name
    
    lListIndex = 0&
    
    cboAccounts.Clear
    
    For lIndex = 1 To Count
        Set Account = m.Accounts(lIndex)
        
        strBrokerName = g.AppBridge.BrokerName(Account.Broker)
        cboAccounts.AddItem Account.AccountNumber & " (" & strBrokerName & ")"
        cboAccounts.ItemData(cboAccounts.NewIndex) = CLng(Val(Account.ID))
        
        If strSelectedAccountID = Account.ID Then
            lListIndex = lIndex
        End If
    Next lIndex
    
    cboAccounts.ListIndex = lListIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccounts.LoadAccountsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerForAccountID
'' Description: Determine the broker for the given Account ID
'' Inputs:      Account ID
'' Returns:     Broker ( Blank if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerForAccountID(ByVal strAccountID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Exists(strAccountID) Then
        strReturn = Str(Item(strAccountID).Broker)
    End If
    
    BrokerForAccountID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccounts.BrokerForAccountID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerNameForAccountID
'' Description: Determine the broker name for the given Account ID
'' Inputs:      Account ID
'' Returns:     Broker Name ( Blank if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerNameForAccountID(ByVal strAccountID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Exists(strAccountID) Then
        strReturn = g.AppBridge.BrokerName(Item(strAccountID).Broker)
    End If
    
    BrokerNameForAccountID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccounts.BrokerNameForAccountID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountNumberForAccountID
'' Description: Determine the account number for the given Account ID
'' Inputs:      Account ID
'' Returns:     Account Number ( Blank if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountNumberForAccountID(ByVal strAccountID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Exists(strAccountID) Then
        strReturn = Item(strAccountID).AccountNumber
    End If
    
    AccountNumberForAccountID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccounts.AccountNumberForAccountID"
    
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
    
    Set m.Accounts = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccounts.Class_Initialize"
    
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

    Set m.Accounts = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccounts.Class_Terminate"
    
End Sub
