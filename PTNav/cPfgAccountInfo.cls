VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPfgAccountInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPfgAccountInfo.cls
'' Description: Class to handle PFG account information
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/25/2009   DAJ         Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strAccountNumber As String          ' Account number
    dCashValue As Double                ' Cash value
    dLiquidValue As Double              ' Liquid value
    dMarginRequired As Double           ' Margin required
    dMarginExcess As Double             ' Margin excess
    dOpenPnl As Double                  ' Open Profit/Loss
    dClosedPnl As Double                ' Closed Profit/Loss
    dNetPnl As Double                   ' Net Profit/Loss
    dLovSov As Double                   ' Long Option Value/Short Option Value
    dBeginningCash As Double            ' Beginning Cash
    dSecuritiesOnDeposit As Double      ' Securities on Deposit
    dOvernightMarketValue As Double     ' Overnight market value
    dCashBalance As Double              ' Cash balance
    dAvailableEquity As Double          ' Available equity
    dOvernightEquity As Double          ' Overnight equity
    dOpeningExcessEquity As Double      ' Opening excess equity
    dScalpedProfit As Double            ' Scalped Profit
    dOpeningMargin As Double            ' Opening margin
    dIntradayMargin As Double           ' Intraday margin
    dMaintenanceMargin As Double        ' Maintenance margin
    dMaintenanceMarginExcess As Double  ' Maintenance margin excess
End Type
Private m As mPrivate

Public Property Get AccountNumber() As String
    AccountNumber = m.strAccountNumber
End Property
Public Property Let AccountNumber(ByVal strAccountNumber As String)
    m.strAccountNumber = strAccountNumber
End Property

Public Property Get CashValue() As Double
    CashValue = m.dCashValue
End Property
Public Property Let CashValue(ByVal dCashValue As Double)
    m.dCashValue = dCashValue
End Property

Public Property Get LiquidValue() As Double
    LiquidValue = m.dLiquidValue
End Property
Public Property Let LiquidValue(ByVal dLiquidValue As Double)
    m.dLiquidValue = dLiquidValue
End Property

Public Property Get MarginRequired() As Double
    MarginRequired = m.dMarginRequired
End Property
Public Property Let MarginRequired(ByVal dMarginRequired As Double)
    m.dMarginRequired = dMarginRequired
End Property

Public Property Get MarginExcess() As Double
    MarginExcess = m.dMarginExcess
End Property
Public Property Let MarginExcess(ByVal dMarginExcess As Double)
    m.dMarginExcess = dMarginExcess
End Property

Public Property Get OpenPnl() As Double
    OpenPnl = m.dOpenPnl
End Property
Public Property Let OpenPnl(ByVal dOpenPnl As Double)
    m.dOpenPnl = dOpenPnl
End Property

Public Property Get ClosedPnl() As Double
    ClosedPnl = m.dClosedPnl
End Property
Public Property Let ClosedPnl(ByVal dClosedPnl As Double)
    m.dClosedPnl = dClosedPnl
End Property

Public Property Get NetPnl() As Double
    NetPnl = m.dNetPnl
End Property
Public Property Let NetPnl(ByVal dNetPnl As Double)
    m.dNetPnl = dNetPnl
End Property

Public Property Get LovSov() As Double
    LovSov = m.dLovSov
End Property
Public Property Let LovSov(ByVal dLovSov As Double)
    m.dLovSov = dLovSov
End Property

Public Property Get BeginningCash() As Double
    BeginningCash = m.dBeginningCash
End Property
Public Property Let BeginningCash(ByVal dBeginningCash As Double)
    m.dBeginningCash = dBeginningCash
End Property

Public Property Get SecuritiesOnDeposit() As Double
    SecuritiesOnDeposit = m.dSecuritiesOnDeposit
End Property
Public Property Let SecuritiesOnDeposit(ByVal dSecuritiesOnDeposit As Double)
    m.dSecuritiesOnDeposit = dSecuritiesOnDeposit
End Property

Public Property Get OvernightMarketValue() As Double
    OvernightMarketValue = m.dOvernightMarketValue
End Property
Public Property Let OvernightMarketValue(ByVal dOvernightMarketValue As Double)
    m.dOvernightMarketValue = dOvernightMarketValue
End Property

Public Property Get CashBalance() As Double
    CashBalance = m.dCashBalance
End Property
Public Property Let CashBalance(ByVal dCashBalance As Double)
    m.dCashBalance = dCashBalance
End Property

Public Property Get AvailableEquity() As Double
    AvailableEquity = m.dAvailableEquity
End Property
Public Property Let AvailableEquity(ByVal dAvailableEquity As Double)
    m.dAvailableEquity = dAvailableEquity
End Property

Public Property Get OvernightEquity() As Double
    OvernightEquity = m.dOvernightEquity
End Property
Public Property Let OvernightEquity(ByVal dOvernightEquity As Double)
    m.dOvernightEquity = dOvernightEquity
End Property

Public Property Get OpeningExcessEquity() As Double
    OpeningExcessEquity = m.dOpeningExcessEquity
End Property
Public Property Let OpeningExcessEquity(ByVal dOpeningExcessEquity As Double)
    m.dOpeningExcessEquity = dOpeningExcessEquity
End Property

Public Property Get ScalpedProfit() As Double
    ScalpedProfit = m.dScalpedProfit
End Property
Public Property Let ScalpedProfit(ByVal dScalpedProfit As Double)
    m.dScalpedProfit = dScalpedProfit
End Property

Public Property Get OpeningMargin() As Double
    OpeningMargin = m.dOpeningMargin
End Property
Public Property Let OpeningMargin(ByVal dOpeningMargin As Double)
    m.dOpeningMargin = dOpeningMargin
End Property

Public Property Get IntradayMargin() As Double
    IntradayMargin = m.dIntradayMargin
End Property
Public Property Let IntradayMargin(ByVal dIntradayMargin As Double)
    m.dIntradayMargin = dIntradayMargin
End Property

Public Property Get MaintenanceMargin() As Double
    MaintenanceMargin = m.dMaintenanceMargin
End Property
Public Property Let MaintenanceMargin(ByVal dMaintenanceMargin As Double)
    m.dMaintenanceMargin = dMaintenanceMargin
End Property

Public Property Get MaintenanceMarginExcess() As Double
    MaintenanceMarginExcess = m.dMaintenanceMarginExcess
End Property
Public Property Let MaintenanceMarginExcess(ByVal dMaintenanceMarginExcess As Double)
    m.dMaintenanceMarginExcess = dMaintenanceMarginExcess
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Handle account information back from the PFG server
'' Inputs:      Pfg Message
'' Returns:     None
''
'' Fields:      Account, Cash, Liquid Value, Margin Required, Margin Excess,
''              Open PNL, Closed PNL, Net PNL, LOVSOV, Beginning Cash,
''              Securities on Deposit, Overnight Market Value, Cash Balance,
''              Available Equity, Overnight Equity, Opening Excess Equity,
''              Scaped Profit, Opening Margin, Intraday Margin, Maintenance
''              Margin, MME
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strPfgMessage As String)
On Error GoTo ErrSection:

    Dim astrMessage As New cGdArray     ' Message split out into an array
    
    astrMessage.SplitFields strPfgMessage
    
    m.strAccountNumber = astrMessage(0)
    m.dCashValue = Val(astrMessage(1))
    m.dLiquidValue = Val(astrMessage(2))
    m.dMarginRequired = Val(astrMessage(3))
    m.dMarginExcess = Val(astrMessage(4))
    m.dOpenPnl = Val(astrMessage(5))
    m.dClosedPnl = Val(astrMessage(6))
    m.dNetPnl = Val(astrMessage(7))
    m.dLovSov = Val(astrMessage(8))
    m.dBeginningCash = Val(astrMessage(9))
    m.dSecuritiesOnDeposit = Val(astrMessage(10))
    m.dOvernightMarketValue = Val(astrMessage(11))
    m.dCashBalance = Val(astrMessage(12))
    m.dAvailableEquity = Val(astrMessage(13))
    m.dOvernightEquity = Val(astrMessage(14))
    m.dOpeningExcessEquity = Val(astrMessage(15))
    m.dScalpedProfit = Val(astrMessage(16))
    m.dOpeningMargin = Val(astrMessage(17))
    m.dIntradayMargin = Val(astrMessage(18))
    m.dMaintenanceMargin = Val(astrMessage(19))
    m.dMaintenanceMarginExcess = Val(astrMessage(20))
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPfgAccountInfo.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Store the information in the member variables in a string
'' Inputs:      None
'' Returns:     String of Information
''
'' Fields:      Account, Cash, Liquid Value, Margin Required, Margin Excess,
''              Open PNL, Closed PNL, Net PNL, LOVSOV, Beginning Cash,
''              Securities on Deposit, Overnight Market Value, Cash Balance,
''              Available Equity, Overnight Equity, Opening Excess Equity,
''              Scaped Profit, Opening Margin, Intraday Margin, Maintenance
''              Margin, MME
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrMessage As New cGdArray     ' Message split out into an array
    
    astrMessage.Create eGDARRAY_Strings

    astrMessage.Add m.strAccountNumber
    astrMessage.Add Str(m.dCashValue)
    astrMessage.Add Str(m.dLiquidValue)
    astrMessage.Add Str(m.dMarginRequired)
    astrMessage.Add Str(m.dMarginExcess)
    astrMessage.Add Str(m.dOpenPnl)
    astrMessage.Add Str(m.dClosedPnl)
    astrMessage.Add Str(m.dNetPnl)
    astrMessage.Add Str(m.dLovSov)
    astrMessage.Add Str(m.dBeginningCash)
    astrMessage.Add Str(m.dSecuritiesOnDeposit)
    astrMessage.Add Str(m.dOvernightMarketValue)
    astrMessage.Add Str(m.dCashBalance)
    astrMessage.Add Str(m.dAvailableEquity)
    astrMessage.Add Str(m.dOvernightEquity)
    astrMessage.Add Str(m.dOpeningExcessEquity)
    astrMessage.Add Str(m.dScalpedProfit)
    astrMessage.Add Str(m.dOpeningMargin)
    astrMessage.Add Str(m.dIntradayMargin)
    astrMessage.Add Str(m.dMaintenanceMargin)
    astrMessage.Add Str(m.dMaintenanceMarginExcess)
    
    ToString = astrMessage.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPfgAccountInfo.ToString"
    
End Function
