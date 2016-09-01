VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTrade.cls
'' Description: Class to hold information about an entry and an optional exit
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' 10/24/2013   DAJ         Pass account number to g.Profit.Profit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strAccountNumber As String          ' Account number
    strBrokerSymbol As String           ' Broker symbol
    strGenesisSymbol As String          ' Genesis symbol
    dEntryTime As Double                ' Time of the entry fill
    strEntryOrderID As String           ' Order ID of the entry fill
    strEntryFillID As String            ' ID of the entry fill
    bEntryIsBuy As Boolean              ' Was the entry fill a buy?
    lEntryQuantity As Long              ' Quantity of the entry
    dEntryPriceBroker As Double         ' Broker Price of the entry
    dEntryPriceGenesis As Double        ' Genesis Price of the entry
    dExitTime As Double                 ' Time of the exit fill
    strExitOrderID As String            ' Order ID of the exit fill
    strExitFillID As String             ' ID of the exit fill
    bExitIsBuy As Boolean               ' Was the exit fill a buy?
    lExitQuantity As Long               ' Quantity of the exit
    dExitPriceBroker As Double          ' Broker Price of the exit
    dExitPriceGenesis As Double         ' Genesis Price of the exit
    dClosedProfit As Double             ' Closed profit
    
    Bars As cGdBars                     ' Bars object
End Type
Private m As mPrivate

Public Property Get AccountNumber() As String
    AccountNumber = m.strAccountNumber
End Property
Public Property Let AccountNumber(ByVal strAccountNumber As String)
    m.strAccountNumber = strAccountNumber
End Property

Public Property Get BrokerSymbol() As String
    BrokerSymbol = m.strBrokerSymbol
End Property
Public Property Let BrokerSymbol(ByVal strBrokerSymbol As String)
    m.strBrokerSymbol = strBrokerSymbol
End Property

Public Property Get GenesisSymbol() As String
    GenesisSymbol = m.strGenesisSymbol
End Property
Public Property Let GenesisSymbol(ByVal strGenesisSymbol As String)
    m.strGenesisSymbol = strGenesisSymbol
    SetBarProperties m.Bars, strGenesisSymbol
End Property

Public Property Get EntryTime() As Double
    EntryTime = m.dEntryTime
End Property
Public Property Let EntryTime(ByVal dEntryTime As Double)
    m.dEntryTime = dEntryTime
End Property

Public Property Get EntryOrderID() As String
    EntryOrderID = m.strEntryOrderID
End Property
Public Property Let EntryOrderID(ByVal strEntryOrderID As String)
    m.strEntryOrderID = strEntryOrderID
End Property

Public Property Get EntryFillID() As String
    EntryFillID = m.strEntryFillID
End Property
Public Property Let EntryFillID(ByVal strEntryFillID As String)
    m.strEntryFillID = strEntryFillID
End Property

Public Property Get EntryIsBuy() As Boolean
    EntryIsBuy = m.bEntryIsBuy
End Property
Public Property Let EntryIsBuy(ByVal bEntryIsBuy As Boolean)
    m.bEntryIsBuy = bEntryIsBuy
End Property

Public Property Get EntryQuantity() As Long
    EntryQuantity = m.lEntryQuantity
End Property
Public Property Let EntryQuantity(ByVal lEntryQuantity As Long)
    m.lEntryQuantity = lEntryQuantity
End Property

Public Property Get EntryPriceBroker() As Double
    EntryPriceBroker = m.dEntryPriceBroker
End Property
Public Property Let EntryPriceBroker(ByVal dEntryPriceBroker As Double)
    m.dEntryPriceBroker = dEntryPriceBroker
End Property

Public Property Get EntryPriceGenesis() As Double
    EntryPriceGenesis = m.dEntryPriceGenesis
End Property
Public Property Let EntryPriceGenesis(ByVal dEntryPriceGenesis As Double)
    m.dEntryPriceGenesis = dEntryPriceGenesis
End Property

Public Property Get ExitTime() As Double
    ExitTime = m.dExitTime
End Property
Public Property Let ExitTime(ByVal dExitTime As Double)
    m.dExitTime = dExitTime
End Property

Public Property Get ExitOrderID() As String
    ExitOrderID = m.strExitOrderID
End Property
Public Property Let ExitOrderID(ByVal strExitOrderID As String)
    m.strExitOrderID = strExitOrderID
End Property

Public Property Get ExitFillID() As String
    ExitFillID = m.strExitFillID
End Property
Public Property Let ExitFillID(ByVal strExitFillID As String)
    m.strExitFillID = strExitFillID
End Property

Public Property Get ExitIsBuy() As Boolean
    ExitIsBuy = m.bExitIsBuy
End Property
Public Property Let ExitIsBuy(ByVal bExitIsBuy As Boolean)
    m.bExitIsBuy = bExitIsBuy
End Property

Public Property Get ExitQuantity() As Long
    ExitQuantity = m.lExitQuantity
End Property
Public Property Let ExitQuantity(ByVal lExitQuantity As Long)
    m.lExitQuantity = lExitQuantity
End Property

Public Property Get ExitPriceBroker() As Double
    ExitPriceBroker = m.dExitPriceBroker
End Property
Public Property Let ExitPriceBroker(ByVal dExitPriceBroker As Double)
    m.dExitPriceBroker = dExitPriceBroker
End Property

Public Property Get ExitPriceGenesis() As Double
    ExitPriceGenesis = m.dExitPriceGenesis
End Property
Public Property Let ExitPriceGenesis(ByVal dExitPriceGenesis As Double)
    m.dExitPriceGenesis = dExitPriceGenesis
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClosedProfit
'' Description: Determine the closed profit for the trade
'' Inputs:      None
'' Returns:     Closed Profit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ClosedProfit() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function

    If (m.lExitQuantity = 0&) Or (Len(m.strGenesisSymbol) = 0) Then
        dReturn = 0#
    Else
        If m.dClosedProfit = kNullData Then
            If m.bEntryIsBuy Then
                m.dClosedProfit = g.Profit.Profit(m.strGenesisSymbol, m.dExitPriceGenesis - m.dEntryPriceGenesis, m.lEntryQuantity, , , , AccountNumber)
            Else
                m.dClosedProfit = g.Profit.Profit(m.strGenesisSymbol, m.dEntryPriceGenesis - m.dExitPriceGenesis, m.lEntryQuantity, , , , AccountNumber)
            End If
        End If
        dReturn = m.dClosedProfit
    End If
    
    ClosedProfit = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTrade.ClosedProfit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OpenProfit
'' Description: Determine the open profit for the trade
'' Inputs:      Current Price
'' Returns:     Open Profit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenProfit(ByVal dCurrentPrice As Double) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim lOpenQuantity As Long           ' Open quantity for the trade
    
    If (Len(m.strGenesisSymbol) = 0) Or (m.lExitQuantity = m.lEntryQuantity) Then
        dReturn = 0#
    Else
        If m.bEntryIsBuy Then
            dReturn = g.Profit.Profit(m.strGenesisSymbol, dCurrentPrice - m.dEntryPriceGenesis, m.lEntryQuantity, , , , AccountNumber)
        Else
            dReturn = g.Profit.Profit(m.strGenesisSymbol, m.dEntryPriceGenesis - dCurrentPrice, m.lEntryQuantity, , , , AccountNumber)
        End If
    End If
    
    OpenProfit = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTrade.OpenProfit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the trade
'' Inputs:      None
'' Returns:     New Copy
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cTrade
On Error GoTo ErrSection:

    Dim aCopy As New cTrade
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTrade.MakeCopy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private properties for a copy of the trade
'' Inputs:      Copy From
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    m = mCopyFrom
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrade.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Setup and intialize class members when the class is intialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Bars = New cGdBars
    m.dClosedProfit = kNullData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrade.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Cleanup class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Bars = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrade.Class_Terminate"
    
End Sub
