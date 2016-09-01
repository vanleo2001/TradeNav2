VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPtAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPtAccounts.cls
'' Description: Collection of cPtAccount objects
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 10/26/2010   DAJ         Added the LastModified property
'' 10/21/2011   DAJ         Clear collection before load
'' 09/16/2014   DAJ         Core Application functions for DLL's; Database Objects;
''                          Move account objects out of NavSuite into NavBroker.DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Accounts As cGdTree                 ' Collection of accounts
    dLastModified As Double             ' Date/Time the collection was last modified
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cPtAccount
Attribute Item.VB_UserMemId = 0
    If m.Accounts.Exists(vIndexOrKey) Then
        Set Item = m.Accounts(vIndexOrKey)
    Else
        Set Item = Nothing
    End If
End Property
Public Property Let Item(ByVal vIndexOrKey As Variant, Account As cPtAccount)
    Set m.Accounts.Item(vIndexOrKey) = Account
    UpdateLastModified
End Property

Public Property Get Count() As Long
    Count = m.Accounts.Count
End Property

Public Property Get LastModified() As Double
    LastModified = m.dLastModified
End Property

Public Function Add(Account As cPtAccount) As Long
    Add = m.Accounts.Add(Account, Account.AccountNumber)
    UpdateLastModified
End Function

Public Function Remove(ByVal vIndexOrKey As Variant) As Boolean
    Remove = m.Accounts.Remove(vIndexOrKey)
    UpdateLastModified
End Function

Public Function Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.Accounts.Exists(vIndexOrKey)
End Function

Public Sub Clear()
    m.Accounts.Clear
    UpdateLastModified
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadForBroker
'' Description: Load accounts for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadForBroker(ByVal nAccountType As eTT_AccountType)
On Error GoTo ErrSection:

    g.BrokerDB.LoadAccountsForAccountType Me, nAccountType

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cPtAccounts.LoadForBroker"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateLastModified
'' Description: Update the last modified property
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateLastModified()
On Error GoTo ErrSection:

    m.dLastModified = g.TnCore.CurrentTime

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cPtAccounts.UpdateLastModified"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the order
'' Inputs:      None
'' Returns:     Copy of the order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cPtAccounts
On Error GoTo ErrSection:

    Dim aCopy As New cPtAccounts        ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cPtAccounts.MakeCopy"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private data for copying purposes
'' Inputs:      Private members to copy from
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    m = mCopyFrom
    
    Set m.Accounts = mCopyFrom.Accounts.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cPtAccounts.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when object is initialized
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
    g.TnCore.RaiseError "cPtAccounts.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when object is terminated
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
    g.TnCore.RaiseError "cPtAccounts.Class_Terminate"

End Sub
