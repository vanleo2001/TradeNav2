VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCattleKeyValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cCattleKeyValue.cls
'' Description: Code to go back and forth from Key Value collections to objects
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountFromKeyValue
'' Description: Build an account object from the given key value collection
'' Inputs:      Key Value Account
'' Returns:     Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountFromKeyValue(ByVal KeyValueAccount As cBrokerMessage) As cAccount
On Error GoTo ErrSection:

    Dim Account As cAccount             ' Account object to return from the function
    
    Set Account = New cAccount
    Account.ID = KeyValueAccount("ID")
    Account.AccountNumber = KeyValueAccount("Number")
    Account.Name = KeyValueAccount("Name")
    Account.FeedYardID = KeyValueAccount("FeedYardID")
    Account.Broker = CLng(Val(KeyValueAccount("Broker")))
    Account.FcmNumber = KeyValueAccount("FcmNumber")
    Account.Deleted = g.Cattle.StringToBool(KeyValueAccount("Deleted"))
    Account.InRefresh = g.Cattle.StringToBool(KeyValueAccount("InRefresh"))
    
    Set AccountFromKeyValue = Account

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattleKeyValue.AccountFromKeyValue"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountToKeyValue
'' Description: Build a key value collection form the given account object
'' Inputs:      Account
'' Returns:     Key Value Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountToKeyValue(ByVal Account As cAccount) As cBrokerMessage
On Error GoTo ErrSection:

    Dim KeyValueAccount As cBrokerMessage ' Key value account to return from the function
    
    Set KeyValueAccount = New cBrokerMessage
    KeyValueAccount.Add "ID", Account.ID
    KeyValueAccount.Add "Number", Account.AccountNumber
    KeyValueAccount.Add "Name", Account.Name
    KeyValueAccount.Add "FeedYardID", Account.FeedYardID
    KeyValueAccount.Add "Broker", Str(Account.Broker)
    KeyValueAccount.Add "FcmNumber", Account.FcmNumber
    If Account.Deleted = True Then
        KeyValueAccount.Add "Deleted", g.Cattle.BoolToString(Account.Deleted)
    End If
    
    Set AccountToKeyValue = KeyValueAccount

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattleKeyValue.AccountToKeyValue"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountFromServerMessage
'' Description: Build an account object from the given server message
'' Inputs:      Server Message
'' Returns:     Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountFromServerMessage(ByVal strServerMessage As String) As cAccount
On Error GoTo ErrSection:

    Dim KeyValueAccount As cBrokerMessage ' Key value account
    
    Set KeyValueAccount = New cBrokerMessage
    KeyValueAccount.FromString strServerMessage
    
    Set AccountFromServerMessage = AccountFromKeyValue(KeyValueAccount)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattleKeyValue.AccountFromServerMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountToServerMessage
'' Description: Build an server message from the given account
'' Inputs:      Account
'' Returns:     Server Message
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountToServerMessage(ByVal Account As cAccount) As String
On Error GoTo ErrSection:

    Dim KeyValueAccount As cBrokerMessage ' Key value account
    
    Set KeyValueAccount = AccountToKeyValue(Account)
    
    AccountToServerMessage = KeyValueAccount.ToString

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattleKeyValue.AccountToServerMessage"
    
End Function
