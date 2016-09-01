VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPtAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPtAccount.cls
'' Description: Properties and methods of a trade tracker account
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/05/2009   DAJ         Added TypeOfAccount and more fields to ON string
'' 06/08/2009   DAJ         Always send colors in the ON string
'' 06/08/2009   DAJ         Send VBButtonFace for non-broker-live accounts
'' 10/07/2009   DAJ         Added support for Linked Orders at broker
'' 12/01/2009   DAJ         Added support for automatic commissions on fills
'' 03/11/2010   DAJ         Call broker RemoveAccount on delete
'' 03/30/2011   DAJ         Format closed balance going to Option Navigator
'' 07/19/2011   DAJ         Added RecalculateFees function
'' 08/23/2011   DAJ         Send FIFO/LIFO flag to Option Navigator
'' 10/31/2011   DAJ         Send Option Nav a message when account deleted
'' 08/21/2012   DAJ         Added FcmAccountNumber for accounts
'' 09/04/2014   DAJ         Pulled Option Navigator conversions out of trade objects
'' 09/10/2014   DAJ         Consolidate the delete account code
'' 09/16/2014   DAJ         Core Application functions for DLL's; Database Objects;
''                          Move account objects out of NavSuite into NavBroker.DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lAccountID As Long                  ' Database identifier of the account
    strAccountNumber As String          ' User defined account number
    strName As String                   ' User defined account name
    AccountType As eTT_AccountType      ' Type of the account
    strUserName As String               ' User name if tied to broker
    strPassword As String               ' Password if tied to broker
    strBroker As String                 ' Broker name
    dComms As Double                    ' Commissions for this account
    dStockFees As Double                ' Stock commissions
    lSecTypeMask As Long                ' Security Type mask
    bFillRT As Boolean                  ' Have TradeNav fill in Real-Time mode?
    StartingDate As Date                ' Date the account was started
    nFillMatchMode As eTT_FillMatchMode ' Fill matching mode for the account
    bHoldOcoAtBroker As Boolean         ' Have OCO's held at broker for this account?
    strFcmAccountNumber As String       ' FCM account number
    
    dStartingBalance As Double          ' Starting balance of the account
    dClosedProfit As Double             ' Closed profit stored in the database
    dTotalFees As Double                ' Total amount of fees
    dTotalAdjustments As Double         ' Total amount of all adjustments
    dCurrentBalance As Double           ' Current closed balance of the account
    dOpenProfit As Double               ' Open profit in the account right now
    dCurrentValue As Double             ' Current value of the account
    
    nStatus As eGDConnectionStatus      ' Connection status for this account
    nTypeOfAccount As eGDTypeOfAccount  ' Type of account
End Type
Private m As mPrivate

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal lAccountID As Long)
    m.lAccountID = lAccountID
End Property

Public Property Get AccountNumber() As String
    AccountNumber = m.strAccountNumber
End Property
Public Property Let AccountNumber(ByVal pData As String)
    m.strAccountNumber = pData
End Property

Public Property Get Name() As String
    Name = m.strName
End Property
Public Property Let Name(ByVal pData As String)
    m.strName = pData
End Property

Public Property Get AccountType() As eTT_AccountType
    AccountType = m.AccountType
End Property
Public Property Let AccountType(ByVal pData As eTT_AccountType)
    m.AccountType = pData
End Property

Public Property Get UserName() As String
    UserName = m.strUserName
End Property
Public Property Let UserName(ByVal pData As String)
    m.strUserName = pData
End Property

Public Property Get Password() As String
    Password = m.strPassword
End Property
Public Property Let Password(ByVal pData As String)
    m.strPassword = pData
End Property

Public Property Get Broker() As String
    Broker = m.strBroker
End Property
Public Property Let Broker(ByVal pData As String)
    m.strBroker = pData
End Property

Public Property Get Comms() As Double
    Comms = m.dComms
End Property
Public Property Let Comms(ByVal pData As Double)
    m.dComms = pData
End Property

Public Property Get StockFees() As Double
    StockFees = m.dStockFees
End Property
Public Property Let StockFees(ByVal dStockFees As Double)
    m.dStockFees = dStockFees
End Property

Public Property Get SecTypeMask() As Long
    SecTypeMask = m.lSecTypeMask
End Property
Public Property Let SecTypeMask(ByVal lSecTypeMask As Long)
    m.lSecTypeMask = lSecTypeMask
End Property

Public Property Get FillRT() As Boolean
    FillRT = m.bFillRT
End Property
Public Property Let FillRT(ByVal bFillRT As Boolean)
    m.bFillRT = bFillRT
End Property

Public Property Get StartingDate() As Date
    StartingDate = m.StartingDate
End Property
Public Property Let StartingDate(ByVal pData As Date)
    m.StartingDate = pData
End Property

Public Property Get FillMatchMode() As eTT_FillMatchMode
    FillMatchMode = m.nFillMatchMode
End Property
Public Property Let FillMatchMode(ByVal nFillMatchMode As eTT_FillMatchMode)
    m.nFillMatchMode = nFillMatchMode
End Property

Public Property Get HoldOcoAtBroker() As Boolean
    HoldOcoAtBroker = m.bHoldOcoAtBroker
End Property
Public Property Let HoldOcoAtBroker(ByVal bHoldOcoAtBroker As Boolean)
    m.bHoldOcoAtBroker = bHoldOcoAtBroker
End Property

Public Property Get StartingBalance() As Double
    StartingBalance = m.dStartingBalance
End Property
Public Property Let StartingBalance(ByVal pData As Double)
    m.dStartingBalance = pData
    m.dCurrentBalance = m.dStartingBalance + m.dTotalAdjustments + m.dClosedProfit - m.dTotalFees
    m.dCurrentValue = m.dCurrentBalance + m.dOpenProfit
End Property

Public Property Get ClosedProfit() As Double
    ClosedProfit = m.dClosedProfit
End Property
Public Property Let ClosedProfit(ByVal dClosedProfit As Double)
    m.dClosedProfit = dClosedProfit
    m.dCurrentBalance = m.dStartingBalance + m.dTotalAdjustments + m.dClosedProfit - m.dTotalFees
    m.dCurrentValue = m.dCurrentBalance + m.dOpenProfit
End Property

Public Property Get TotalAdjustments() As Double
    TotalAdjustments = m.dTotalAdjustments
End Property
Public Property Let TotalAdjustments(ByVal dTotalAdjustments As Double)
    m.dTotalAdjustments = dTotalAdjustments
    m.dCurrentBalance = m.dStartingBalance + m.dTotalAdjustments + m.dClosedProfit - m.dTotalFees
    m.dCurrentValue = m.dCurrentBalance + m.dOpenProfit
End Property

Public Property Get TotalFees() As Double
    TotalFees = m.dTotalFees
End Property
Public Property Let TotalFees(ByVal dTotalFees As Double)
    m.dTotalFees = dTotalFees
    m.dCurrentBalance = m.dStartingBalance + m.dTotalAdjustments + m.dClosedProfit - m.dTotalFees
    m.dCurrentValue = m.dCurrentBalance + m.dOpenProfit
End Property

Public Property Get CurrentBalance() As Double
    CurrentBalance = m.dCurrentBalance
End Property
Public Property Let CurrentBalance(ByVal pData As Double)
    m.dCurrentBalance = pData
End Property
Public Property Get CurrentClosedBalance() As Double
    CurrentClosedBalance = m.dCurrentBalance
End Property

Public Property Get OpenProfit() As Double
    OpenProfit = m.dOpenProfit
End Property
Public Property Let OpenProfit(ByVal dOpenProfit As Double)
    m.dOpenProfit = dOpenProfit
    m.dCurrentValue = m.dCurrentBalance + m.dOpenProfit
End Property

Public Property Get CurrentValue() As Double
    CurrentValue = m.dCurrentValue
End Property

Public Property Get ConnectionStatus() As eGDConnectionStatus
    ConnectionStatus = m.nStatus
End Property
Public Property Let ConnectionStatus(ByVal nStatus As eGDConnectionStatus)
    m.nStatus = nStatus
End Property

Public Property Get TypeOfAccount() As eGDTypeOfAccount
    TypeOfAccount = m.nTypeOfAccount
End Property
Public Property Let TypeOfAccount(ByVal nTypeOfAccount As eGDTypeOfAccount)
    m.nTypeOfAccount = nTypeOfAccount
End Property

Public Property Get FcmAccountNumber() As String
    FcmAccountNumber = m.strFcmAccountNumber
End Property
Public Property Let FcmAccountNumber(ByVal strFcmAccountNumber As String)
    m.strFcmAccountNumber = strFcmAccountNumber
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the account from the database
'' Inputs:      Account ID to load
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal lAccountID As Long) As Boolean
On Error GoTo ErrSection:

    Load = g.BrokerDB.LoadAccount(Me, lAccountID)
    
ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cPtAccount.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the account to the database
'' Inputs:      None
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Save() As Boolean
On Error GoTo ErrSection:
    
    Save = g.BrokerDB.SaveAccount(Me)
    
ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cPtAccount.Save"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Delete
'' Description: Delete the account out of the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete()
On Error GoTo ErrSection:

    g.BrokerDB.DeleteAccount AccountID
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cPtAccount.Delete"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateFees
'' Description: Recalculate the fees associated with fills
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateFees()
On Error GoTo ErrSection:

    TotalFees = g.BrokerDB.FeesForAccount(AccountID)
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cPtAccount.RecalculateFees"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the order
'' Inputs:      None
'' Returns:     Copy of the order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cPtAccount
On Error GoTo ErrSection:

    Dim aCopy As New cPtAccount         ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cPtAccount.MakeCopy"

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
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cPtAccount.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialization for the class
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    m.nStatus = eGDConnectionStatus_Disconnected
    m.nTypeOfAccount = eGDTypeOfAccount_Simulated
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cPtAccount.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up for the class
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cPtAccount.Class_Terminate"
    
End Sub
