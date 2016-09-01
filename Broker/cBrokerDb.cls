VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerDb.cls
'' Description: Class for interaction with the Trade Tracker database
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/10/2014   DAJ         Created
'' 10/24/2014   DAJ         Core Application functions for DLL's; Database Objects
'' 12/01/2014   DAJ         Clear out accounts collection before loading it up
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAccount
'' Description: Load the account from the database
'' Inputs:      Account, Account ID, Recordset
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadAccount(Account As cPtAccount, ByVal lAccountID As Long, Optional ByVal rs As Recordset = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    If rs Is Nothing Then
        Set rs = g.TradeTrackerDB.AccountForID(lAccountID)
    End If
    
    bReturn = False
    If Not (rs.BOF And rs.EOF) Then
        Account.AccountID = rs!AccountID
        Account.AccountNumber = Trim(rs!AccountNumber)
        Account.Name = Trim(rs!Name)
        Account.StartingBalance = rs!StartingBalance
        Account.CurrentBalance = rs!CurrentBalance
        Account.StartingDate = rs!StartingDate
        Account.UserName = NullChk(rs!UserName)
        Account.Password = "" ' DecryptField(rs!Password)
        Account.AccountType = NullChk(rs!AccountType, 0)
        Account.Broker = NullChk(rs!Broker)
        Account.Comms = NullChk(rs!Comms, 0)
        Account.StockFees = NullChk(rs!StockFees, 0#)
        Account.SecTypeMask = NullChk(rs!SecTypeMask, 0)
        Account.FillRT = NullChk(rs!FillRT, True)
        Account.ClosedProfit = NullChk(rs!ClosedProfit, 0#)
        Account.TotalAdjustments = NullChk(rs!TotalAdjustments, 0#)
        Account.TotalFees = NullChk(rs!TotalFees, 0#)
        Account.FillMatchMode = NullChk(rs!FillMatchMode, 0&)
        Account.HoldOcoAtBroker = NullChk(rs!HoldOcoAtBroker, False)
        Account.FcmAccountNumber = NullChk(rs!FcmAccountNumber)
        
        Account.TypeOfAccount = mBroker.TypeOfAccount(Account.AccountType)
            
        bReturn = True
    End If

    LoadAccount = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerDb.LoadAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveAccount
'' Description: Save the account to the database
'' Inputs:      Account
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SaveAccount(Account As cPtAccount) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.TradeTrackerDB.AccountForID(Account.AccountID)
    If rs.BOF And rs.EOF Then
        rs.AddNew
    Else
        rs.Edit
    End If
    
    rs!AccountNumber = Trim(Account.AccountNumber)
    rs!Name = Trim(Account.Name)
    rs!StartingBalance = Account.StartingBalance
    rs!StartingDate = Account.StartingDate
    rs!CurrentBalance = Account.CurrentBalance
    rs!UserName = Account.UserName
    'EncryptField rs!Password, Account.Password
    rs!Password = ""
    rs!AccountType = Account.AccountType
    rs!Broker = Account.Broker
    rs!Comms = Account.Comms
    rs!StockFees = Account.StockFees
    rs!SecTypeMask = Account.SecTypeMask
    rs!FillRT = Account.FillRT
    rs!ClosedProfit = Account.ClosedProfit
    rs!TotalAdjustments = Account.TotalAdjustments
    rs!TotalFees = Account.TotalFees
    rs!FillMatchMode = Account.FillMatchMode
    rs!HoldOcoAtBroker = Account.HoldOcoAtBroker
    rs!FcmAccountNumber = Account.FcmAccountNumber
    
    Account.AccountID = rs!AccountID
    rs.Update
    
    SaveAccount = True

ErrExit:
    Exit Function

ErrSection:
    g.TnCore.RaiseError "cBrokerDb.SaveAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteAccount
'' Description: Delete the account from the database
'' Inputs:      Account ID
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DeleteAccount(ByVal lAccountID As Long) As Boolean
On Error GoTo ErrSection:

    DeleteAccount = g.TradeTrackerDB.DeleteRecord(g.TradeTrackerDB.AccountForID(lAccountID))

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerDb.DeleteAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FeesForAccount
'' Description: Calculate the total fees for all fills for the given account
'' Inputs:      Account ID
'' Returns:     Total Fees
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FeesForAccount(ByVal lAccountID As Long) As Double
On Error GoTo ErrSection:

    Dim rsFills As Recordset            ' Recordset into the database
    Dim dTotalFees As Double            ' Total sum of the fees

    ' Reload all of the fills for this account...
    Set rsFills = g.TradeTrackerDB.FillsForAccountID(lAccountID)
    
    ' Recalculate the total sum of fees for this account...
    dTotalFees = 0#
    If Not (rsFills.BOF And rsFills.EOF) Then
        rsFills.MoveFirst
        Do While Not rsFills.EOF
            If rsFills!Fees <> kNullData Then
                dTotalFees = dTotalFees + Abs(rsFills!Fees)
            End If
            rsFills.MoveNext
        Loop
    End If
    
    FeesForAccount = dTotalFees
    
ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerDb.FeesForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAccountsForAccountType
'' Description: Load all of the accounts for the given account type
'' Inputs:      Account Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadAccountsForAccountType(Accounts As cPtAccounts, ByVal nAccountType As eTT_AccountType)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Account As cPtAccount           ' Account object

    Accounts.Clear

    Set rs = g.TradeTrackerDB.AccountsForAccountType(nAccountType)
    Do While Not rs.EOF
        Set Account = New cPtAccount
        
        LoadAccount Account, rs!AccountID, rs
        Accounts.Add Account
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cBrokerDb.LoadAccountsForAccountType"
    
End Sub
