VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPtOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPtOrder.cls
'' Description: Wrapper object for an order
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 03/19/2009   DAJ         Set Broker and Account on Order from Option Nav
'' 08/13/2009   DAJ         Added the IsAutoExit flag for orders
'' 09/01/2009   DAJ         Use new Parked order status, OCO fixes
'' 10/07/2009   DAJ         Added support for Linked Orders at broker
'' 10/08/2009   DAJ         Fix for advanced order trailing stop
'' 12/01/2009   DAJ         Added support for automatic commissions on fills
'' 12/04/2009   DAJ         Changed automatic commissions for stocks
'' 12/09/2009   DAJ         Fixes for broker held OCOs
'' 01/29/2010   DAJ         Added numbars required stuff for conditional orders (#5580)
'' 05/06/2010   DAJ         Use Triggered price for wrong side of market check if applicable,
''                          Handle Parking/Submitting order with triggered by order (#5715)
'' 05/12/2010   DAJ         Fixed infinite loops in HasTriggered orders and HasOcoOrders
'' 05/25/2010   DAJ         Enhanced logging for advanced orders
'' 09/16/2010   DAJ         More information in e-mail order alerts (#5891)
'' 01/24/2011   DAJ         Send order date to Option Navigator with an order
'' 01/26/2011   DAJ         Added Previous Broker ID to Option Nav order message
'' 01/26/2011   DAJ         Added Message field to Option Nav order message
'' 01/28/2011   DAJ         Implemented 'ChangeOrderStatus' and 'BrokerDate' functions
'' 02/04/2011   DAJ         Added the 'FillsText' function
'' 03/07/2011   DAJ         Changed some Option Nav calls, Don't update alerts when filled
'' 03/16/2011   DAJ         When building from Option Nav string, clear out the legs
'' 05/27/2011   DAJ         Added IsAutomated flag for orders
'' 11/21/2011   DAJ         Don't overwrite Genesis ID with Access ID from Option Nav
'' 01/30/2012   DAJ         Option Nav Journal Image
'' 02/14/2012   DAJ         Added multi-leg order support
'' 03/01/2012   DAJ         Fix for loading bars for MIT orders
'' 03/21/2012   DAJ         Fix for amending a partially filled order
'' 05/15/2012   DAJ         Option Nav fixes
'' 05/30/2012   DAJ         Added check for all realtime data available for conditional orders
'' 08/08/2012   DAJ         If the Order Date or Session Date aren't set on Save, set them
'' 10/09/2012   DAJ         Added PreviousOrderID property
'' 10/23/2012   DAJ         Always set expiration date for a SimTradeMessageString
'' 10/24/2012   DAJ         Only add TIF to order text if it has symbol ( not for Trade Console )
'' 12/11/2012   DAJ         Contingency Orders, OcaGroup Name
'' 01/18/2013   TLB         Init stop/limit prices to Null, since spread orders can be at negative prices
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 04/16/2013   DAJ         Log when auto trade id changes, update broker info object upon save
'' 06/24/2013   DAJ         Send default symbol to MarketsInExpressions ( #6868 )
'' 07/15/2013   DAJ         Allow 'Of Monthly' in conditional order expression
'' 07/30/2013   DAJ         Data Pending status for Conditional Orders
'' 08/01/2013   DAJ         Change to whether or not order is considered conditional
'' 08/01/2013   DAJ         Fix for whether conditional orders are data pending/trigger pending/working
'' 10/10/2013   DAJ         Added more logging for conditional orders
'' 11/04/2013   DAJ         More logging for conditional orders; Add bars to stream
'' 01/31/2014   DAJ         Have user specify TIF on contingecy orders
'' 03/03/2014   DAJ         Got rid of the DajLog function in mTradeTracker
'' 04/10/2014   DAJ         Added ability to use remaining quantity in Engine String
'' 08/29/2014   DAJ         Added SetSessionDateForFeedTime function
'' 09/04/2014   DAJ         Pulled Option Navigator, SimTrade conversions out of trade objects
'' 10/03/2014   DAJ         Added some logging for OCO
'' 10/29/2014   DAJ         Remove old synthetic order/MIT code
'' 10/31/2014   DAJ         Added Market-On-Open order type
'' 11/18/2014   DAJ         Fix for building engine string for Market-On-Close orders
'' 12/29/2014   DAJ         Fix for loading markets in LoadAllMarkets if no data loaded for main market
'' 09/09/2015   DAJ         Dump date for TIF for GTD instead of 'GTD' in OrderText
'' 09/14/2015   DAJ         Added Tradier
'' 11/06/2015   DAJ         Moved trailing stop code from order into conditional orders; New move with trigger flag on triggered orders
'' 11/18/2015   DAJ         Set the bar properties again when the order legs have been set
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum eGDConditionTypes
    eGDCondition_Price = 0
    eGDCondition_Time
    eGDCondition_Custom
End Enum

Private Type mPrivate
    lOrderID As Long
    OrderDate As Date
    dFees As Double
    Status As eTT_OrderStatus
    dStatusDate As Double
    strBrokerID As String
    strExchangeID As String
    lExpiration As Long
    lAccountID As Long
    dDateSent As Double
    dDateHostRecd As Double
    dDateExchRecd As Double
    dDateExchAckn As Double
    strMessage As String
    strGenesisOrderID As String
    lAutoTradeItemID As Long
    lCancelOrderID As Long              ' ID of corresponding OCO order
    lTriggerOrderID As Long             ' ID of corresponding trigger order
    strTriggerOptions As String         ' Options for order trigger
    strConditionOptions As String       ' Options for conditional orders
    strCondition As String              ' Coded text condition
    strConditionSymbol As String        ' Symbol for the price condition
    dTrailAmount As Double              ' Amount to trail HH or LL by
    strTrailOptions As String           ' Options for trail orders
    Contingency As cContingencyOrders   ' Options for contingency orders
    lExitPos As Long                    ' Exit percent of position
    lPrevQty As Long                    ' Previous quantity
    strSession As String                ' Session code for the order
    bIsSnapshot As Boolean              ' Is this a snapshot order?
    lSessionDate As Long                ' Session date for the order
    strPreviousBrokerID As String       ' Previous broker order ID
    lGroupID As Long                    ' ID of the order group (for Option Navigator bundling)
    strGroupName As String              ' Name of the order group (for Option Navigator bundling)
    bIsAutoExit As Boolean              ' Is this an auto exit order?
    lBrokerCancelOrderID As Long        ' Broker Cancel Order ID
    bIsAutomated As Boolean             ' Was the order created in an automated fashion?
    lQuantity As Long                   ' Quantity for the order
    nOrderType As eTT_OrderType         ' Order Type for the order
    dLimitPrice As Double               ' Limit Price for the order
    dStopPrice As Double                ' Stop Price for the order
    strUnderlyingSymbol As String       ' Underlying symbol
    lUnderlyingSymbolID As Long         ' Underlying symbol ID
    nDebitCredit As eGDDebitCredit      ' Debit/Credit
    strOcaGroup As String               ' OCA Group for Interactive Brokers
    
    strSpreadSymbol As String           ' Spread symbol
    strOptionNavImageFile As String     ' Option Navigator image file
    
    OrderLegs As cOrderLegs             ' Order leg information
    Fills As cGdTree                    ' Collection of fills for this order
    SpreadFills As cGdTree              ' Collection of spread fills for this order
    History As cGdTree                  ' Collection of order history for this order
    
    Bars As cGdBars                     ' Bars object for running expression
    SecondaryMarkets As cGdTree         ' Collection of secondary markets
    PreviousOrder As cPtOrder           ' Order that was amended to make this one
    
    nBroker As eTT_AccountType          ' Broker for the order
    bRefreshed As Boolean               ' Has this order been refreshed?
    strAdvanced As String               ' Advanced order text
    lPreviousOrderID As Long            ' Previous order ID
End Type
Private m As mPrivate

Public Property Get OrderID() As Long
    OrderID = m.lOrderID
End Property
Public Property Let OrderID(ByVal pData As Long)
    m.lOrderID = pData
End Property

Public Property Get Quantity() As Long
    Quantity = m.lQuantity
End Property
Public Property Let Quantity(ByVal lQuantity As Long)
    m.lPrevQty = m.lQuantity
    m.lQuantity = lQuantity
End Property

Public Property Get OrderDate() As Date
    OrderDate = m.OrderDate
End Property
Public Property Let OrderDate(ByVal pData As Date)
    m.OrderDate = pData
End Property

Public Property Get Buy() As Boolean
    Buy = m.OrderLegs(1).IsBuy
End Property
Public Property Let Buy(ByVal pData As Boolean)
    m.OrderLegs(1).IsBuy = pData
End Property

Public Property Get OrderType() As eTT_OrderType
    OrderType = m.nOrderType
End Property
Public Property Let OrderType(ByVal nOrderType As eTT_OrderType)
    m.nOrderType = nOrderType
End Property

Public Property Get StopPrice() As Double
    StopPrice = m.dStopPrice
End Property
Public Property Get StopPriceString() As String
    If m.dStopPrice = kNullData Then
        StopPriceString = ""
    Else
        StopPriceString = m.Bars.PriceDisplay(m.dStopPrice)
    End If
End Property
Public Property Let StopPrice(ByVal dStopPrice As Double)
    m.dStopPrice = dStopPrice
End Property

Public Property Get LimitPrice() As Double
    LimitPrice = m.dLimitPrice
End Property
Public Property Get LimitPriceString() As String
    If m.dLimitPrice = kNullData Then
        LimitPriceString = ""
    Else
        LimitPriceString = m.Bars.PriceDisplay(m.dLimitPrice)
    End If
End Property
Public Property Let LimitPrice(ByVal dLimitPrice As Double)
    m.dLimitPrice = dLimitPrice
End Property

Public Property Get MitPrice() As Double
    MitPrice = m.dLimitPrice
End Property
Public Property Get MitPriceString() As String
    If m.dLimitPrice = kNullData Then
        MitPriceString = ""
    Else
        MitPriceString = m.Bars.PriceDisplay(m.dLimitPrice)
    End If
End Property
Public Property Let MitPrice(ByVal dMITPrice As Double)
    m.dLimitPrice = dMITPrice
End Property

Public Property Get Enter() As Boolean
    Enter = m.OrderLegs(1).IsEntry
End Property
Public Property Let Enter(ByVal pData As Boolean)
    m.OrderLegs(1).IsEntry = pData
End Property

Public Property Get Fees() As Double
    Fees = m.dFees
End Property
Public Property Let Fees(ByVal pData As Double)
    m.dFees = pData
End Property

Public Property Get Symbol() As String
    Symbol = m.OrderLegs(1).Symbol
End Property
Public Property Get SymbolID() As Long
    SymbolID = m.OrderLegs(1).SymbolID
End Property

Public Property Get Status() As eTT_OrderStatus
    Status = m.Status
End Property
Public Property Let Status(ByVal nNewStatus As eTT_OrderStatus)
    m.Status = nNewStatus
End Property

Public Property Get StatusDate() As Double
    StatusDate = m.dStatusDate
End Property
Public Property Let StatusDate(ByVal pData As Double)
    m.dStatusDate = pData
End Property

Public Property Get BrokerID() As String
    BrokerID = m.strBrokerID
End Property
Public Property Let BrokerID(ByVal pData As String)
    m.strBrokerID = pData
End Property

Public Property Get ExchangeID() As String
    ExchangeID = m.strExchangeID
End Property
Public Property Let ExchangeID(ByVal pData As String)
    m.strExchangeID = pData
End Property

Public Property Get Expiration() As Long
    Expiration = m.lExpiration
End Property
Public Property Let Expiration(ByVal pData As Long)
    m.lExpiration = pData
End Property

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal pData As Long)
    m.lAccountID = pData
    m.nBroker = g.Broker.AccountTypeForID(m.lAccountID)
End Property

Public Property Get DateSent() As Double
    DateSent = m.dDateSent
End Property
Public Property Let DateSent(ByVal pData As Double)
    m.dDateSent = pData
End Property

Public Property Get DateHostRecd() As Double
    DateHostRecd = m.dDateHostRecd
End Property
Public Property Let DateHostRecd(ByVal pData As Double)
    m.dDateHostRecd = pData
End Property

Public Property Get DateExchRecd() As Double
    DateExchRecd = m.dDateExchRecd
End Property
Public Property Let DateExchRecd(ByVal pData As Double)
    m.dDateExchRecd = pData
End Property

Public Property Get DateExchAckn() As Double
    DateExchAckn = m.dDateExchAckn
End Property
Public Property Let DateExchAckn(ByVal pData As Double)
    m.dDateExchAckn = pData
End Property

Public Property Get Message() As String
    Message = m.strMessage
End Property
Public Property Let Message(ByVal pData As String)
    m.strMessage = pData
End Property

Public Property Get GenesisOrderID() As String
    GenesisOrderID = m.strGenesisOrderID
End Property
Public Property Let GenesisOrderID(ByVal strGenesisOrderID As String)
    m.strGenesisOrderID = strGenesisOrderID
End Property

Public Property Get AutoTradeItemID() As Long
    AutoTradeItemID = m.lAutoTradeItemID
End Property
Public Property Let AutoTradeItemID(ByVal lAutoTradeItemID As Long)
    m.lAutoTradeItemID = lAutoTradeItemID
End Property

Public Property Get CancelOrderID() As Long
    CancelOrderID = m.lCancelOrderID
End Property
Public Property Let CancelOrderID(ByVal lCancelOrderID As Long)
    If lCancelOrderID <> m.lCancelOrderID Then
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "( " & OrderIdText & " ): Cancel Order ID changed from " & Str(m.lCancelOrderID) & " to " & Str(lCancelOrderID)
        m.lCancelOrderID = lCancelOrderID
    End If
End Property

Public Property Get TriggerOrderID() As Long
    TriggerOrderID = m.lTriggerOrderID
End Property
Public Property Let TriggerOrderID(ByVal lTriggerOrderID As Long)
    m.lTriggerOrderID = lTriggerOrderID
End Property

Public Property Get TriggerOptions() As String
    TriggerOptions = m.strTriggerOptions
End Property
Public Property Let TriggerOptions(ByVal strTriggerOptions As String)
    m.strTriggerOptions = strTriggerOptions
End Property

Public Property Get ConditionOptions() As String
    ConditionOptions = m.strConditionOptions
End Property
Public Property Let ConditionOptions(ByVal strConditionOptions As String)
    m.strConditionOptions = strConditionOptions
    m.strCondition = ""
End Property

Public Property Get TrailAmount() As Double
    TrailAmount = m.dTrailAmount
End Property
Public Property Let TrailAmount(ByVal dTrailAmount As Double)
    m.dTrailAmount = dTrailAmount
End Property

Public Property Get TrailOptions() As String
    TrailOptions = m.strTrailOptions
End Property
Public Property Let TrailOptions(ByVal strTrailOptions As String)
    If strTrailOptions <> m.strTrailOptions Then
        g.Broker.BrokerDebug Broker, vbTab & vbTab & "Trail options changed from '" & m.strTrailOptions & "' to '" & strTrailOptions & "'"
        m.strTrailOptions = strTrailOptions
    End If
End Property

Public Property Get Contingency() As cContingencyOrders
    Set Contingency = m.Contingency
End Property
Public Property Let Contingency(ByVal Contingency As cContingencyOrders)
    Set m.Contingency = Contingency
End Property

Public Property Get ExitPos() As Long
    ExitPos = m.lExitPos
End Property
Public Property Let ExitPos(ByVal lExitPos As Long)
    m.lExitPos = lExitPos
End Property

Public Property Get Session() As String
    Session = m.strSession
End Property
Public Property Let Session(ByVal strSession As String)
    m.strSession = strSession
End Property

Public Property Get PrevQty() As Long
    PrevQty = m.lPrevQty
End Property
Public Property Let PrevQty(ByVal lPrevQty As Long)
    m.lPrevQty = lPrevQty
End Property

Public Property Get Fills() As cGdTree
    Set Fills = m.Fills
End Property
Public Property Let Fills(ByVal pData As cGdTree)
    Set m.Fills = pData
End Property

Public Property Get SpreadFills() As cGdTree
    Set SpreadFills = m.SpreadFills
End Property
Public Property Let SpreadFills(ByVal pData As cGdTree)
    Set m.SpreadFills = pData
End Property

Public Property Get History() As cGdTree
    Set History = m.History
End Property
Public Property Let History(ByVal tHistory As cGdTree)
    Set m.History = tHistory
End Property

Public Property Get TriggerOnPartial() As Boolean
    TriggerOnPartial = False
    If Len(m.strTriggerOptions) > 0 Then
        If Parse(m.strTriggerOptions, ",", 1) = "1" Then TriggerOnPartial = True
    End If
End Property

Public Property Get TriggerConfirm() As Boolean
    TriggerConfirm = False
    If Len(m.strTriggerOptions) > 0 Then
        If Parse(m.strTriggerOptions, ",", 2) = "1" Then TriggerConfirm = True
    End If
End Property

Public Property Get MoveWithTrigger() As Boolean
    MoveWithTrigger = CBool(Val(Parse(m.strTriggerOptions, ",", 4)))
End Property

Public Property Get SymbolOrSymbolID() As Variant
    SymbolOrSymbolID = m.OrderLegs(1).SymbolOrSymbolID
End Property

Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.OrderLegs(1).SymbolOrSymbolID = ConvertSymbol(vSymbolOrSymbolID)
    SetBarProperties m.Bars, m.OrderLegs(1).SymbolOrSymbolID
End Property

Public Property Get TimeInForce() As eTT_TimeInForce
    If m.lExpiration < 0 Then
        TimeInForce = eTT_TimeInForce_Day
    ElseIf m.lExpiration = 0 Then
        TimeInForce = eTT_TimeInForce_GTC
    Else
        TimeInForce = eTT_TimeInForce_GTD
    End If
End Property

Public Property Get IsSnapshot() As Boolean
    IsSnapshot = m.bIsSnapshot
End Property
Public Property Let IsSnapshot(ByVal bIsSnapshot As Boolean)
    m.bIsSnapshot = bIsSnapshot
End Property

Public Property Get SessionDate() As Long
    SessionDate = m.lSessionDate
End Property
Public Property Let SessionDate(ByVal lSessionDate As Long)
    m.lSessionDate = lSessionDate
End Property

Public Property Get PreviousBrokerID() As String
    PreviousBrokerID = m.strPreviousBrokerID
End Property
Public Property Let PreviousBrokerID(ByVal strPreviousBrokerID As String)
    m.strPreviousBrokerID = strPreviousBrokerID
End Property

Public Property Get Exchange() As String
    Exchange = m.OrderLegs(1).Exchange
End Property
Public Property Let Exchange(ByVal strExchange As String)
    m.OrderLegs(1).Exchange = strExchange
End Property

Public Property Get OrderLegs() As cOrderLegs
    Set OrderLegs = m.OrderLegs
End Property
Public Property Let OrderLegs(ByVal Legs As cOrderLegs)
    Set m.OrderLegs = Legs
    SetBarProperties m.Bars, m.OrderLegs(1).SymbolOrSymbolID
End Property

Public Property Get NumberOfLegs() As Long
    NumberOfLegs = m.OrderLegs.Count
End Property

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property

Public Property Get Refreshed() As Boolean
    Refreshed = m.bRefreshed
End Property
Public Property Let Refreshed(ByVal bRefreshed As Boolean)
    m.bRefreshed = bRefreshed
End Property

Public Property Get BarsHandle() As Long
    If m.Bars Is Nothing Then
        BarsHandle = -99999
    Else
        BarsHandle = m.Bars.BarsHandle
    End If
End Property

Public Property Get GroupID() As Long
    GroupID = m.lGroupID
End Property
Public Property Let GroupID(ByVal lGroupID As Long)
    m.lGroupID = lGroupID
End Property

Public Property Get GroupName() As String
    GroupName = m.strGroupName
End Property
Public Property Let GroupName(ByVal strGroupName As String)
    m.strGroupName = strGroupName
End Property

Public Property Get IsAutoExit() As Boolean
    IsAutoExit = m.bIsAutoExit
End Property
Public Property Let IsAutoExit(ByVal bIsAutoExit As Boolean)
    m.bIsAutoExit = bIsAutoExit
End Property

Public Property Get BrokerCancelOrderID() As Long
    BrokerCancelOrderID = m.lBrokerCancelOrderID
End Property
Public Property Let BrokerCancelOrderID(ByVal lBrokerCancelOrderID As Long)
    If lBrokerCancelOrderID <> m.lBrokerCancelOrderID Then
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "( " & OrderIdText & " ): Broker Cancel Order ID changed from " & Str(m.lBrokerCancelOrderID) & " to " & Str(lBrokerCancelOrderID)
        m.lBrokerCancelOrderID = lBrokerCancelOrderID
    End If
End Property

Public Property Get IsAutomated() As Boolean
    IsAutomated = m.bIsAutomated
End Property
Public Property Let IsAutomated(ByVal bIsAutomated As Boolean)
    m.bIsAutomated = bIsAutomated
End Property

Public Property Get UnderlyingSymbol() As String
    UnderlyingSymbol = m.strUnderlyingSymbol
End Property

Public Property Get UnderlyingSymbolID() As Long
    UnderlyingSymbolID = m.lUnderlyingSymbolID
End Property

Public Property Get UnderlyingSymbolOrSymbolID() As Variant
    If m.lUnderlyingSymbolID = 0 Then
        UnderlyingSymbolOrSymbolID = m.strUnderlyingSymbol
    Else
        UnderlyingSymbolOrSymbolID = m.lUnderlyingSymbolID
    End If
End Property
Public Property Let UnderlyingSymbolOrSymbolID(ByVal vUnderlyingSymbolOrSymbolID As Variant)
    m.lUnderlyingSymbolID = GetSymbolID(vUnderlyingSymbolOrSymbolID)
    m.strUnderlyingSymbol = GetSymbol(vUnderlyingSymbolOrSymbolID)
End Property

Public Property Get DebitCredit() As eGDDebitCredit
    DebitCredit = m.nDebitCredit
End Property
Public Property Let DebitCredit(ByVal nDebitCredit As eGDDebitCredit)
    m.nDebitCredit = nDebitCredit
End Property

Public Property Get SpreadSymbol() As String
    SpreadSymbol = m.strSpreadSymbol
End Property

Public Property Get OptionNavImageFile() As String
    OptionNavImageFile = m.strOptionNavImageFile
End Property
Public Property Let OptionNavImageFile(ByVal strOptionNavImageFile As String)
    m.strOptionNavImageFile = strOptionNavImageFile
End Property

Public Property Get PreviousOrder() As cPtOrder
    Set PreviousOrder = m.PreviousOrder
End Property
Public Property Let PreviousOrder(ByVal PrevOrder As cPtOrder)
    Set m.PreviousOrder = PrevOrder
End Property

Public Property Get PreviousOrderID() As Long
    PreviousOrderID = m.lPreviousOrderID
End Property
Public Property Let PreviousOrderID(ByVal lPreviousOrderID As Long)
    m.lPreviousOrderID = lPreviousOrderID
End Property

Public Property Get OcaGroup() As String
    OcaGroup = m.strOcaGroup
End Property
Public Property Let OcaGroup(ByVal strOcaGroup As String)
    m.strOcaGroup = strOcaGroup
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the order from the database
'' Inputs:      Order ID, Recordset, Fills Recordset
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal lOrderID As Long, Optional rs As Recordset = Nothing, Optional rsFills As Recordset = Nothing) As Boolean
On Error GoTo ErrSection

    Dim rs2 As Recordset                ' Recordset into the database
    Dim Fill As cPtFill                 ' Temporary fill object
    Dim bReturn As Boolean              ' Return value for the function
    
    If rs Is Nothing Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                    "WHERE [OrderID]=" & Str(lOrderID) & ";", dbOpenDynaset)
    End If
    
    If Not rs.EOF Then
        m.lOrderID = rs!OrderID
        m.OrderDate = rs!OrderDate
        m.dFees = rs!Fees
        m.Status = NullChk(rs!Status, 0)
        m.dStatusDate = NullChk(rs!StatusDate, 0)
        m.strBrokerID = NullChk(rs!BrokerOrderID)
        m.strExchangeID = NullChk(rs!ExchangeOrderID)
        m.lExpiration = NullChk(rs!Expiration, 0)
        m.lAccountID = NullChk(rs!AccountID, 0)
        m.nBroker = g.Broker.AccountTypeForID(m.lAccountID)
        m.dDateSent = NullChk(rs!DateSent, 0)
        m.dDateHostRecd = NullChk(rs!DateHostRecd, 0)
        m.dDateExchRecd = NullChk(rs!DateExchRecd, 0)
        m.dDateExchAckn = NullChk(rs!DateExchAckn, 0)
        m.strMessage = NullChk(rs!Message)
        m.strGenesisOrderID = NullChk(rs!GenesisOrderID)
        m.lAutoTradeItemID = NullChk(rs!AutoTradeItemID, 0&)
        m.lCancelOrderID = NullChk(rs!CancelOrderID, 0&)
        m.lTriggerOrderID = NullChk(rs!TriggerOrderID, 0&)
        m.strTriggerOptions = NullChk(rs!TriggerOptions, "")
        m.strConditionOptions = NullChk(rs!ConditionOptions, "")
        m.dTrailAmount = NullChk(rs!TrailAmount, 0#)
        m.strTrailOptions = NullChk(rs!TrailOptions, "")
        m.lExitPos = NullChk(rs!ExitPos, 0&)
        m.strSession = NullChk(rs!Session, "")
        m.bIsSnapshot = NullChk(rs!IsSnapshot, False)
        m.lSessionDate = NullChk(rs!SessionDate, 0&)
        m.strPreviousBrokerID = NullChk(rs!PreviousBrokerID, "")
        m.lGroupID = NullChk(rs!GroupID, 0&)
        m.strGroupName = NullChk(rs!GroupName, "")
        m.bIsAutoExit = NullChk(rs!IsAutoExit, False)
        m.lBrokerCancelOrderID = NullChk(rs!BrokerCancelOrderID, 0)
        m.bIsAutomated = NullChk(rs!IsAutomated, 0)
        m.lQuantity = rs!Quantity
        m.nOrderType = rs!OrderType
        m.dLimitPrice = rs!LimitPrice
        m.dStopPrice = rs!StopPrice
        m.strUnderlyingSymbol = rs!UnderlyingSymbol
        m.lUnderlyingSymbolID = rs!UnderlyingSymbolID
        m.nDebitCredit = rs!DebitCredit
        m.strOcaGroup = NullChk(rs!OcaGroup)
        
        Set m.Contingency = New cContingencyOrders
        m.Contingency.FromString NullChk(rs!ContingencyOptions, "")
    
        ' Load the information for the order legs...
        m.OrderLegs.Load m.lOrderID
        
        ' Load the information for the fills...
        If rsFills Is Nothing Then
            Set rsFills = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [OrderID]=" & Str(lOrderID) & " ORDER BY [FillID];", dbOpenDynaset)
        End If
        If Not (rsFills.BOF And rsFills.EOF) Then
            rsFills.MoveFirst
            Do While Not rsFills.EOF
                If rsFills!OrderID = m.lOrderID Then
                    Set Fill = New cPtFill
                    If Fill.Load(rsFills!FillID, rsFills) Then
                        If Fill.IsSpread Then
                            m.SpreadFills.Add Fill, Str(Fill.FillID)
                        Else
                            m.Fills.Add Fill, Str(Fill.FillID)
                        End If
                    End If
                End If
                
                rsFills.MoveNext
            Loop
        End If
        
        ' Do not allow a continuous contract in an order...
        If InStr(Symbol, "-0") <> 0 Then
            SymbolOrSymbolID = RollSymbolForDate(Symbol, m.OrderDate)
            Save
        End If
        
        ' Allow negative prices only for a SpreadSymbol
        If Not IsSpreadSymbol(Symbol) Then
            If m.dStopPrice <= 0 Then m.dStopPrice = kNullData
            If m.dLimitPrice <= 0 Then m.dLimitPrice = kNullData
        End If
        
        ' Due to a bug, we were storing the long date instead of the julian in the expiration
        ' field when it came back from the trade server.  This should fix this problem...
        If m.lExpiration > 20000000 Then
            m.lExpiration = JulFromLong(m.lExpiration)
            Save
        End If
        
        SetBarProperties m.Bars, SymbolOrSymbolID
        
        m.strAdvanced = AdvancedOrderText
        
        LoadHistory
        
        If NumberOfLegs > 1 Then
            BuildSpreadSymbol
        End If
        
        If Len(m.strPreviousBrokerID) > 0 Then
            Set m.PreviousOrder = New cPtOrder
            m.PreviousOrder.LoadForBrokerID m.strPreviousBrokerID, rsFills
        Else
            Set m.PreviousOrder = Nothing
        End If
        
        bReturn = True
    End If
    
    Load = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadForBrokerID
'' Description: Load the order from the database for the given broker ID
'' Inputs:      Broker ID, Fills recordset
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadForBrokerID(ByVal strBrokerID As String, Optional rsFills As Recordset = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    
    bReturn = False
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [BrokerOrderID]='" & strBrokerID & "';", dbOpenDynaset)
    If Not rs.BOF And rs.EOF Then
        bReturn = Load(rs!OrderID, rs, rsFills)
    End If
    
    LoadForBrokerID = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.LoadForBrokerID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reload
'' Description: Reload the order from the database
'' Inputs:      None
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Reload() As Boolean
On Error GoTo ErrSection:

    Reload = Load(m.lOrderID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.Reload"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the order to the database
'' Inputs:      Submit Contingency Orders?, Update Broker Info?
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Save(Optional ByVal bSubmitContingency As Boolean = True, Optional ByVal bUpdateBrokerInfo As Boolean = True) As Boolean
On Error GoTo ErrSection

    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Temporary fill object
    Dim nPrevStatus As eTT_OrderStatus  ' Previous order status
    Dim lOldOrderID As Long             ' Old order ID
    Dim strAdvanced As String           ' Advanced order text
    Dim dFees As Double                 ' Automatic fees to generate if necessary
    Dim bAllFeesNull As Boolean         ' Are all of the fill fees null?
    Dim lOldAutoTradeID As Long         ' Previous auto trade ID
        
    ' Do this first so that the contingency string gets updated before we write
    ' to the database...
    If (bSubmitContingency = True) Then
        If (m.Status = eTT_OrderStatus_Working) Or (m.Status = eTT_OrderStatus_Partial) Then
            SubmitContingencyOrders
        End If
    End If
        
    ' Save the Order information...
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                "WHERE [OrderID]=" & Str(m.lOrderID) & ";", dbOpenDynaset)
                
    If rs.EOF = False Then
        rs.Edit
        lOldAutoTradeID = rs!AutoTradeItemID
    Else
        rs.AddNew
        lOldAutoTradeID = kNullData
    End If
    
    nPrevStatus = rs!Status
    lOldOrderID = m.lOrderID
    
    If SecurityType(SymbolOrSymbolID, True) = "S" Then
        dFees = g.Broker.AutomaticCommission(AccountID, SymbolOrSymbolID)
    Else
        dFees = kNullData
    End If
    
    If m.OrderDate = 0 Then
        m.OrderDate = BrokerDate(CurrentTime("", Symbol, True))
    End If
    SetSessionDate
    
    ' 11/05/2015 DAJ: When a trailing stop order goes working, set the submit time if it isn't already set...
    If (m.Status = eTT_OrderStatus_Working) And (Len(m.strTrailOptions) > 0) Then
        If Val(Parse(m.strTrailOptions, ",", 2)) = 0 Then
            TrailOptions = Parse(m.strTrailOptions, ",", 1) & "," & Str(CurrentTime(m.Bars.Prop(eBARS_ExchangeTimeZoneInf), Symbol, True))
        End If
    End If
        
    rs!OrderDate = m.OrderDate
    rs!Fees = m.dFees
    rs!Status = m.Status
    rs!StatusDate = m.dStatusDate
    rs!BrokerOrderID = m.strBrokerID
    rs!ExchangeOrderID = m.strExchangeID
    rs!Expiration = m.lExpiration
    rs!AccountID = m.lAccountID
    rs!DateSent = m.dDateSent
    rs!DateHostRecd = m.dDateHostRecd
    rs!DateExchRecd = m.dDateExchRecd
    rs!DateExchAckn = m.dDateExchAckn
    If Len(m.strMessage) > 255 Then m.strMessage = Left(m.strMessage, 255)
    rs!Message = m.strMessage
    rs!GenesisOrderID = m.strGenesisOrderID
    rs!AutoTradeItemID = m.lAutoTradeItemID
    rs!CancelOrderID = m.lCancelOrderID
    rs!TriggerOrderID = m.lTriggerOrderID
    rs!TriggerOptions = m.strTriggerOptions
    rs!ConditionOptions = m.strConditionOptions
    rs!TrailAmount = m.dTrailAmount
    rs!TrailOptions = m.strTrailOptions
    rs!ContingencyOptions = m.Contingency.ToString
    rs!ExitPos = m.lExitPos
    rs!Session = m.strSession
    rs!IsSnapshot = m.bIsSnapshot
    rs!SessionDate = m.lSessionDate
    rs!PreviousBrokerID = m.strPreviousBrokerID
    rs!GroupID = m.lGroupID
    rs!GroupName = m.strGroupName
    rs!IsAutoExit = m.bIsAutoExit
    rs!BrokerCancelOrderID = m.lBrokerCancelOrderID
    rs!IsAutomated = m.bIsAutomated
    rs!Quantity = m.lQuantity
    rs!OrderType = m.nOrderType
    rs!LimitPrice = m.dLimitPrice
    rs!StopPrice = m.dStopPrice
    rs!UnderlyingSymbol = m.strUnderlyingSymbol
    rs!UnderlyingSymbolID = m.lUnderlyingSymbolID
    rs!DebitCredit = m.nDebitCredit
    rs!OcaGroup = m.strOcaGroup
    
    m.lOrderID = rs!OrderID
    rs.Update
    
    ' Save the order legs...
    m.OrderLegs.Save m.lOrderID
    
    ' Save the Fills...
    bAllFeesNull = True
    For lIndex = 1 To m.Fills.Count
        If m.Fills(lIndex).Fees(False) <> kNullData Then
            bAllFeesNull = False
        End If
        
        m.Fills(lIndex).OrderID = m.lOrderID
        m.Fills(lIndex).IsSnapshot = m.bIsSnapshot
        m.Fills(lIndex).Save
        
        If m.Fills.Key(lIndex) <> Str(m.Fills(lIndex).FillID) Then
            m.Fills.Key(lIndex) = Str(m.Fills(lIndex).FillID)
        End If
    Next lIndex
    
    ' If we just got an Access Order ID, send an order ID changed event to Option Nav
    ' changing from the Genesis Order ID to the new Access Order ID...
    If (lOldOrderID = 0&) And (Len(m.strGenesisOrderID) > 0) Then
        SendOrderIdChangeToOptNav m.lAccountID, m.strGenesisOrderID, Str(m.lOrderID)
    End If
    
    ' If this is a stock and all of the fees on the fills are null and there is a default
    ' commission rate set for stocks, set the fees on the first fill to the commission...
    If (bAllFeesNull = True) And (m.Fills.Count > 0) And (dFees <> kNullData) Then
        m.Fills(1).Fees = dFees
        m.Fills(1).Save
    End If
    
    ' Delete any Fills that no longer exist...
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                "WHERE [OrderID]=" & Str(m.lOrderID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        If Not m.Fills.Exists(Str(rs!FillID)) Then
            If Not m.SpreadFills.Exists(Str(rs!FillID)) Then
                Set Fill = New cPtFill
                If Fill.Load(rs!FillID, rs) Then
                    Fill.Delete "No Longer Exists in Order"
                End If
            End If
        End If
        
        rs.MoveNext
    Loop
    
    SaveHistory
    
    ' Save the Cancel Order ID on the other order as well...
    UpdateOCO
    UpdateBrokerOCO
    
    ' Handle an order status change...
    If (lOldOrderID <> 0&) And (m.Status <> nPrevStatus) Then
        ' Let the FillCallback routine call g.Alerts.OrderStatusChange as the order
        ' fills, but call it here for any other order status change...
        If (m.Status <> eTT_OrderStatus_Filled) And (m.Status <> eTT_OrderStatus_Partial) Then
            g.Alerts.OrderStatusChange Me
        End If
        CheckTriggerByOrders Me
    End If
    
    ' If the advanced order information has changed, dump it to the appropriate log...
    strAdvanced = AdvancedOrderText
    If strAdvanced <> m.strAdvanced Then
        m.strAdvanced = strAdvanced
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "Advanced Order Info (" & Str(m.lOrderID) & ", '" & m.strGenesisOrderID & "', '" & m.strBrokerID & "'): " & m.strAdvanced
    End If
    
    If (m.lAutoTradeItemID <> lOldAutoTradeID) And (lOldAutoTradeID <> kNullData) Then
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & OrderText(True, True, True) & ": Auto Trade Item ID changed from " & Str(lOldAutoTradeID) & " to " & Str(m.lAutoTradeItemID)
    End If
    
    If bUpdateBrokerInfo = True Then
        g.Broker.RefreshOrder Me
    End If
        
    Save = True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.Save"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Delete
'' Description: Delete this order out of the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [OrderID]=" & Str(m.lOrderID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        rs.Delete
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.Delete"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetFill
'' Description: Either Add or Replace the given fill in the fills collection
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetFill(Fill As cPtFill)
On Error GoTo ErrSection:

    Set Fills(Str(Fill.FillID)) = Fill

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPtOrder.SetFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EarliestFillDate
'' Description: Determine the date/time of the first fill
'' Inputs:      None
'' Returns:     Date/Time of the first fill
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EarliestFillDate() As Double
On Error GoTo ErrSection:

    Dim dDate As Double                 ' Date to return from the function
    Dim lIndex As Long                  ' Index into a for loop
    
    If Not m.PreviousOrder Is Nothing Then
        dDate = m.PreviousOrder.EarliestFillDate
    Else
        dDate = 0
    End If
    
    For lIndex = 1 To m.Fills.Count
        If dDate = 0 Or m.Fills(lIndex).FillDate < dDate Then
            dDate = m.Fills(lIndex).FillDate
        End If
    Next lIndex
    
    EarliestFillDate = dDate

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.EarliestFillDate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LatestFillDate
'' Description: Determine the date/time of the last fill
'' Inputs:      None
'' Returns:     Date/Time of the last fill
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LatestFillDate() As Double
On Error GoTo ErrSection:

    Dim dDate As Double                 ' Date to return from the function
    Dim lIndex As Long                  ' Index into a for loop
    
    If Not m.PreviousOrder Is Nothing Then
        dDate = m.PreviousOrder.LatestFillDate
    Else
        dDate = 0
    End If
    
    For lIndex = 1 To m.Fills.Count
        If m.Fills(lIndex).FillDate > dDate Then
            dDate = m.Fills(lIndex).FillDate
        End If
    Next lIndex
    
    LatestFillDate = dDate

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.LatestFillDate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AvgFillPrice
'' Description: Determine the average price of all of the fills
'' Inputs:      None
'' Returns:     Average price of all of the fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AvgFillPrice() As Double
On Error GoTo ErrSection:

    Dim dPrice As Double                ' Sum of the prices of the fills
    Dim lIndex As Long                  ' Index into a for loop
    Dim lQuantity As Long               ' Sum of the quantities of the fills
    Dim Fill As New cPtFill             ' Temporary fill object
    
    If Not m.PreviousOrder Is Nothing Then
        dPrice = m.PreviousOrder.AvgFillPrice
        lQuantity = m.PreviousOrder.FillQuantity
    Else
        dPrice = 0#
        lQuantity = 0&
    End If
    
    If IsSpreadSymbol(Symbol) Then
        For lIndex = 1 To m.SpreadFills.Count
            Set Fill = m.SpreadFills(lIndex)
            dPrice = dPrice + (Fill.Price * Fill.Quantity)
            lQuantity = lQuantity + Fill.Quantity
        Next lIndex
    Else
        For lIndex = 1 To m.Fills.Count
            Set Fill = m.Fills(lIndex)
            dPrice = dPrice + (Fill.Price * Fill.Quantity)
            lQuantity = lQuantity + Fill.Quantity
        Next lIndex
    End If
    
    If lQuantity > 0& Then AvgFillPrice = dPrice / lQuantity

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.AvgFillPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillQuantity
'' Description: Total Quantity of all of the fills for this order
'' Inputs:      Order Leg
'' Returns:     Sum of Quantity of fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillQuantity(Optional ByVal lOrderLeg As Long = 1&) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lQuantity As Long               ' Cumulative Quantity of the fills
    Dim Fill As New cPtFill             ' Temporary fill object
    Dim vSymbolOrSymbolID As Variant    ' Symbol or symbol ID of the leg
    Dim lDivisor As Long                ' Quantity Divisor
       
    If Not m.PreviousOrder Is Nothing Then
        lQuantity = m.PreviousOrder.FillQuantity(lOrderLeg)
    Else
        lQuantity = 0&
    End If
    vSymbolOrSymbolID = m.OrderLegs(lOrderLeg).SymbolOrSymbolID
    
    If IsSpreadSymbol(m.OrderLegs(lOrderLeg).Symbol) Then
        For lIndex = 1 To m.SpreadFills.Count
            Set Fill = m.SpreadFills(lIndex)
            If Fill.SymbolOrSymbolID = vSymbolOrSymbolID Then
                lQuantity = lQuantity + Fill.Quantity
            End If
        Next lIndex
    Else
        For lIndex = 1 To m.Fills.Count
            Set Fill = m.Fills(lIndex)
            If Fill.SymbolOrSymbolID = vSymbolOrSymbolID Then
                lQuantity = lQuantity + Fill.Quantity
            End If
        Next lIndex
    End If
    
    If NumberOfLegs > 1 Then
        lDivisor = 0&
        For lIndex = 1 To NumberOfLegs
            lDivisor = lDivisor + OrderLegs(lIndex).Multiplier
        Next lIndex
        If lDivisor > 0 Then
            lQuantity = lQuantity / lDivisor
        End If
    End If
    
    FillQuantity = lQuantity

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.FillQuantity"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillFees
'' Description: Total amount of commissions paid on all fills
'' Inputs:      None
'' Returns:     Total commissions
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillFees() As Double
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim dFees As Double                 ' Total fees
    Dim Fill As New cPtFill             ' Temporary fill object
    
    If Not m.PreviousOrder Is Nothing Then
        dFees = m.PreviousOrder.Fees
    Else
        dFees = 0#
    End If
    
    For lIndex = 1 To m.Fills.Count
        Set Fill = m.Fills(lIndex)
        dFees = dFees + Fill.Fees
    Next lIndex
    
    FillFees = dFees

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.FillFees"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddHistory
'' Description: Add an order history item for this order
'' Inputs:      Date, Text, Status, Sequence
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddHistory(ByVal dDate As Double, ByVal strText As String, ByVal strStatus As String, ByVal lSeq As Long)
On Error GoTo ErrSection:

    Dim strHistory As String            ' History string
    
    strHistory = "0" & vbTab & Str(dDate) & vbTab & strText & vbTab & strStatus & vbTab
    strHistory = strHistory & Str(lSeq)
    
    If Not m.History.Exists(Str(lSeq)) Then
        m.History.Add strHistory, Str(lSeq)
    Else
        m.History(Str(lSeq)) = strHistory
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.AddHistory"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LinkStatus
'' Description: Return a string of the current link status of the order
'' Inputs:      None
'' Returns:     Link Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LinkStatus() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If m.lCancelOrderID <> 0 Then
        strReturn = "Linked"
    ElseIf m.lBrokerCancelOrderID < 0 Then
        If m.Status = eTT_OrderStatus_Parked Then
            strReturn = "Link Parked"
        Else
            strReturn = "Linking at " & g.Broker.BrokerName(m.nBroker)
        End If
    ElseIf m.lBrokerCancelOrderID > 0 Then
        strReturn = "Linked at " & g.Broker.BrokerName(m.nBroker)
    End If
    
    LinkStatus = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.LinkStatus"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerDate
'' Description: Convert the local date/time given to the broker date/time
'' Inputs:      Local Time
'' Returns:     Converted Time
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerDate(ByVal dLocalDateTime As Double) As Double
On Error GoTo ErrSection:

    BrokerDate = ConvertToBrokerDate(dLocalDateTime, m.nBroker, Symbol, True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.BrokerDate"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeOrderStatus
'' Description: Change the order status on the order
'' Inputs:      New Order Status, Broker Info Object, Check Trigger Order?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeOrderStatus(ByVal nNewStatus As eTT_OrderStatus, Optional ByVal BInfo As cBrokerInfo = Nothing, Optional ByVal bCheckTrigger As Boolean = True)
On Error GoTo ErrSection:

    Dim nPrevStatus As eTT_OrderStatus  ' Previous order status
    
    nPrevStatus = m.Status
    
    m.Status = nNewStatus
    m.dStatusDate = BrokerDate(CurrentTime)
    Save False
    
    If BInfo Is Nothing Then
        Set BInfo = g.Broker.BrokerInfo(m.nBroker)
    End If
    BInfo.AddOrder Me, False
    
    OrderCallback Me, , , bCheckTrigger
    SendOrderToOptionNav Me, False
    
    If nNewStatus <> nPrevStatus Then
        g.ActivityLogs.AddOrderToActivityLog m.nBroker, Me
        g.Broker.BrokerDebug m.nBroker, OrderText(True, True, True) & ": Order Status changed from '" & OrderStatus(nPrevStatus) & "' to '" & OrderStatus(nNewStatus) & "'"
        
        If nNewStatus = eTT_OrderStatus_Working Then
            If SubmitContingencyOrders Then
                Save False
            End If
        ElseIf (IsOpenOrder(nNewStatus, True) = False) And (nNewStatus <> eTT_OrderStatus_Filled) Then
            CancelContingencyOrders
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.ChangeOrderStatus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddToStream
'' Description: Add all of the bars for this order to the stream
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddToStream()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    If Not m.SecondaryMarkets Is Nothing Then
        For lIndex = 1 To m.SecondaryMarkets.Count
            g.RealTime.SpliceBars m.SecondaryMarkets(lIndex)
            g.RealTime.AddTickBuffer m.SecondaryMarkets(lIndex)
        Next lIndex
    End If
    
    If Not m.Bars Is Nothing Then
        g.RealTime.SpliceBars m.Bars
        g.RealTime.AddTickBuffer m.Bars
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.AddToStream"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetSessionDateForFeedTime
'' Description: Set the session date for the order given the current feed time
'' Inputs:      Feed Time
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetSessionDateForFeedTime(ByVal dFeedTime As Double)
On Error GoTo ErrSection:

    SessionDate = m.Bars.SessionDateForTradeTime(ConvertTimeZone(dFeedTime, "NY", m.Bars.Prop(eBARS_ExchangeTimeZoneInf)))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.SetSessionDateForFeedTime"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MarketPrice
'' Description: Determine the market price for the order taking into account
''              triggering orders if applicable
'' Inputs:      Called from a triggering order?
'' Returns:     Market Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MarketPrice(ByVal bTriggered As Boolean) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim Order As cPtOrder               ' Triggering order
    
    ' If there is no triggering order and we are not being called from a triggered order, then
    ' use the current market price for the symbol...
    If (TriggerOrderID = 0) And (bTriggered = False) Then
        dReturn = g.RealTime.LastKnownPrice(SymbolOrSymbolID)
        
    ' Otherwise if this is a limit order and we are being called from a triggered order, then
    ' use the limit price since that is where the market will be if this order fills...
    ElseIf (OrderType = eTT_OrderType_Limit) And (bTriggered = True) Then
        dReturn = LimitPrice
        
    ' Otherwise if this is a MIT order and we are being called from a triggered order, then
    ' use the MIT price since that is where the market will be if this order fills...
    ElseIf (OrderType = eTT_OrderType_MIT) And (bTriggered = True) Then
        dReturn = MitPrice
        
    ' Otherwise if this is a non-trailing stop order and we are being called from a triggered
    ' order, then use the stop price since that is where the market will be if this order fills...
    ElseIf (OrderType = eTT_OrderType_Stop) And (TrailAmount = 0) And (bTriggered = True) Then
        dReturn = StopPrice
        
    ' Otherwise if there is a triggering order, use the market price for the triggering order...
    ElseIf (TriggerOrderID <> 0) Then
        Set Order = New cPtOrder
        If Order.Load(TriggerOrderID) Then
            dReturn = Order.MarketPrice(True)
        End If
        
    ' Otherwise, use the current market price for the symbol ( this must be a market order or
    ' something like that )...
    Else
        dReturn = g.RealTime.LastKnownPrice(SymbolOrSymbolID)
    End If
    
    MarketPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.MarketPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveHistory
'' Description: Save the order history items to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveHistory()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrLine As New cGdArray        ' Temporary array of information
    Dim rs As Recordset                 ' Recordset into the database
    Dim lHistoryID As Long              ' ID of the history item
    Dim dDateTime As Double             ' Date/Time of the history item
    Dim strText As String               ' Text of the history item
    Dim strStatus As String             ' Status of the history item
    Dim lSeq As Long                    ' Sequence of the history item
        
    For lIndex = 1 To m.History.Count
        astrLine.SplitFields m.History(lIndex), vbTab
        lHistoryID = CLng(Val(astrLine(0)))
        dDateTime = Val(astrLine(1))
        strText = astrLine(2)
        strStatus = astrLine(3)
        lSeq = CLng(Val(astrLine(4)))
        
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderHistory] " & _
                    "WHERE [OrderID]=" & Str(m.lOrderID) & " AND [HistoryID]=" & Str(lHistoryID) & ";", dbOpenDynaset)
        If rs.EOF And rs.BOF Then
            rs.AddNew
        Else
            rs.Edit
        End If
        
        rs!OrderID = m.lOrderID
        rs!Seq = lSeq
        rs!StatusDate = dDateTime
        rs!Status = strStatus
        rs!Message = strText
        astrLine(0) = Str(rs!HistoryID)
        rs.Update
        m.History(lIndex) = astrLine.JoinFields(vbTab)
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.SaveHistory"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFill
'' Description: Update the appropriate fill
'' Inputs:      Fill ID, Fill
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UpdateFill(ByVal lFillID As Long, Fill As cPtFill) As Boolean
On Error GoTo ErrSection:

    Dim dFillDate As Double             ' Earliest fill date for the order
    Dim lFillQuantity As Long           ' Total quantity of fills for the order

    Set m.Fills(Str(lFillID)) = Fill
    
    dFillDate = EarliestFillDate
    If dFillDate < m.OrderDate Then OrderDate = dFillDate
    
    Buy = Fill.Buy
    
    lFillQuantity = FillQuantity
    If lFillQuantity > Quantity Then
        Quantity = lFillQuantity
        m.Status = eTT_OrderStatus_Filled
        m.dStatusDate = LatestFillDate
        
    ElseIf lFillQuantity < Quantity Then
        If m.Status = eTT_OrderStatus_Filled Then
            m.Status = eTT_OrderStatus_Partial
            m.dStatusDate = LatestFillDate
        End If
    End If
    
    UpdateFill = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.UpdateFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadHistory
'' Description: Load the order history from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadHistory()
On Error GoTo ErrSection:

    Dim strHistory As String            ' String of history item information
    Dim rs As Recordset                 ' Recordset from the database
    
    m.History.Clear
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderHistory] " & _
                "WHERE [OrderID]=" & Str(m.lOrderID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        strHistory = Str(rs!HistoryID) & vbTab & Str(rs!StatusDate) & vbTab
        strHistory = strHistory & rs!Message & vbTab & rs!Status & vbTab
        strHistory = strHistory & Str(rs!Seq)
        
        m.History(Str(rs!Seq)) = strHistory
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.LoadHistory"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderText
'' Description: Compile a descriptive string of the order
'' Inputs:      Include Symbol?, Include Account?, Include ID's?
'' Returns:     Descriptive String of the Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderText(Optional ByVal bWithSymbol As Boolean = True, Optional ByVal bWithAccount As Boolean = False, Optional ByVal bWithID As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' String to return
    Dim lQuantity As Long               ' Quantity
    
    If NumberOfLegs = 0 Then
        strReturn = ""
    Else
        If NumberOfLegs = 1 Then
            lQuantity = Quantity * OrderLegs(1).Multiplier
            
            If Buy = True Then
                If m.lExitPos = 0& Or ((m.Status <> eTT_OrderStatus_TriggerPending) And (m.Status <> eTT_OrderStatus_Open)) Then
                    strReturn = "Buy " & Str(lQuantity)
                Else
                    strReturn = "Buy to Exit"
                End If
            Else
                If m.lExitPos = 0& Or ((m.Status <> eTT_OrderStatus_TriggerPending) And (m.Status <> eTT_OrderStatus_Open)) Then
                    strReturn = "Sell " & Str(lQuantity)
                Else
                    strReturn = "Sell to Exit"
                End If
            End If
            
            If bWithSymbol = True Then
                strReturn = strReturn & " '" & Symbol & "' at "
            Else
                strReturn = strReturn & " at "
            End If
        Else
            strReturn = Str(Quantity)
            Select Case DebitCredit
                Case eGDDebitCredit_Credit
                    strReturn = strReturn & " Credit Spread at "
                Case eGDDebitCredit_Debit
                    strReturn = strReturn & " Debit Spread at "
                Case eGDDebitCredit_Even
                    strReturn = strReturn & " Even Spread at "
            End Select
        End If
        
        Select Case OrderType
            Case eTT_OrderType_Market
                strReturn = strReturn & "MARKET"
                
            Case eTT_OrderType_Stop
                If Len(Parse(m.strTriggerOptions, ",", 3)) = 0 Then
                    strReturn = strReturn & Trim(m.Bars.PriceDisplay(StopPrice)) & " STOP"
                ElseIf Buy Then
                    strReturn = strReturn & "Triggered Entry Price + " & Trim(PriceDisplay(Val(Parse(m.strTriggerOptions, ",", 3)), Symbol)) & " STOP"
                Else
                    strReturn = strReturn & "Triggered Entry Price - " & Trim(PriceDisplay(Val(Parse(m.strTriggerOptions, ",", 3)), Symbol)) & " STOP"
                End If
                
            Case eTT_OrderType_Limit
                If Len(Parse(m.strTriggerOptions, ",", 3)) = 0 Then
                    strReturn = strReturn & Trim(m.Bars.PriceDisplay(LimitPrice)) & " LIMIT"
                ElseIf Buy Then
                    strReturn = strReturn & "Triggered Entry Price - " & Trim(PriceDisplay(Val(Parse(m.strTriggerOptions, ",", 3)), Symbol)) & " LIMIT"
                Else
                    strReturn = strReturn & "Triggered Entry Price + " & Trim(PriceDisplay(Val(Parse(m.strTriggerOptions, ",", 3)), Symbol)) & " LIMIT"
                End If
            
            Case eTT_OrderType_StopWithLimit
                strReturn = strReturn & Trim(m.Bars.PriceDisplay(StopPrice)) & " STOP with a " & Trim(m.Bars.PriceDisplay(LimitPrice)) & " LIMIT"
        
            Case eTT_OrderType_MarketOnClose
                strReturn = strReturn & "MARKET ON CLOSE"
            
            Case eTT_OrderType_StopCloseOnly
                strReturn = strReturn & Trim(m.Bars.PriceDisplay(StopPrice)) & " STOP CLOSE ONLY"
            
            Case eTT_OrderType_LimitCloseOnly
                strReturn = strReturn & Trim(m.Bars.PriceDisplay(LimitPrice)) & " LIMIT CLOSE ONLY"
            
            Case eTT_OrderType_StopWithLimitCloseOnly
                strReturn = strReturn & Trim(m.Bars.PriceDisplay(StopPrice)) & " STOP with a " & Trim(m.Bars.PriceDisplay(LimitPrice)) & " LIMIT CLOSE ONLY"
                
            Case eTT_OrderType_MIT
                strReturn = strReturn & Trim(m.Bars.PriceDisplay(MitPrice)) & " MIT"
        
            Case eTT_OrderType_MarketOnOpen
                strReturn = strReturn & "MARKET ON OPEN"
            
        End Select
        
        If bWithSymbol Then
            If m.lExpiration > 0 Then
                strReturn = strReturn & " ( TIF = " & DateFormat(m.lExpiration, MM_DD_YYYY) & " )"
            Else
                strReturn = strReturn & " ( TIF = " & mTradeTracker.TimeInForce(TimeInForce) & " )"
            End If
        End If
        
        If bWithAccount Then
            strReturn = strReturn & " in account '" & g.Broker.AccountNumberForID(m.lAccountID) & "'"
        End If
        
        If bWithID Then
            strReturn = strReturn & " (" & OrderIdText & ")"
        End If
    End If
    
    OrderText = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.OrderText"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillsText
'' Description: Build a text string to describe the fills
'' Inputs:      None
'' Returns:     Text about the fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillsText() As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strReturn As String             ' Return value for the function
    Dim Fill As cPtFill                 ' Fill object
    Dim FillPrices As cGdTree           ' Collection of fill prices
    Dim strFillPrice As String          ' Display fill price
    
    If m.Fills.Count = 0 Then
        strReturn = "No Fills"
    Else
        Set FillPrices = New cGdTree
        
        For lIndex = 1 To m.Fills.Count
            Set Fill = m.Fills(lIndex)
            strFillPrice = m.Bars.PriceDisplay(Fill.Price)
            If m.Fills.Exists(strFillPrice) Then
                FillPrices(strFillPrice) = FillPrices(strFillPrice) + Fill.Quantity
            Else
                FillPrices.Add Fill.Quantity, strFillPrice
            End If
        Next lIndex
        
        If Buy = True Then
            strReturn = "Bought "
        Else
            strReturn = "Sold "
        End If
        
        For lIndex = 1 To FillPrices.Count
            If lIndex = 1 Then
                strReturn = strReturn & Str(FillPrices(lIndex)) & " @ " & FillPrices.Key(lIndex)
            Else
                strReturn = strReturn & ", " & Str(FillPrices(lIndex)) & " @ " & FillPrices.Key(lIndex)
            End If
        Next lIndex
    End If
    
    FillsText = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.FillsText"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AdvancedOrderText
'' Description: Compile a descriptive string of the order
'' Inputs:      Include Symbol?, Include Advanced Information?
'' Returns:     Descriptive String of the Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AdvancedOrderText() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of information
    Dim strTemp As String               ' Temporary string
    Dim strReturn As String             ' String to return
    Dim strCondition As String          ' Condition for order
    Dim rs As Recordset                 ' Recordset into the database
    Dim strContingency As String        ' Contingency options
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    If m.lCancelOrderID <> 0& Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [OrderID]=" & Str(m.lCancelOrderID) & ";", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            If Len(rs!BrokerOrderID) > 0 Then
                astrReturn.Add "OCO: " & rs!GenesisOrderID & ", " & rs!BrokerOrderID
            Else
                astrReturn.Add "OCO: " & rs!GenesisOrderID
            End If
        End If
    End If
    
    If m.lTriggerOrderID <> 0& Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [OrderID]=" & Str(m.lTriggerOrderID) & ";", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            If Len(rs!BrokerOrderID) > 0 Then
                strTemp = "OTO: " & rs!GenesisOrderID & ", " & rs!BrokerOrderID
            Else
                strTemp = "OTO: " & rs!GenesisOrderID
            End If
            
            If TriggerConfirm Then
                strTemp = strTemp & ", Confirm = True"
            Else
                strTemp = strTemp & ", Confirm = False"
            End If
            If TriggerOnPartial Then
                strTemp = strTemp & ", Partial = True"
            Else
                strTemp = strTemp & ", Partial = False"
            End If
            astrReturn.Add strTemp
        End If
    End If
    
    If TriggerTime <> 0 Then
        astrReturn.Add "Trigger Time: " & DateFormat(TriggerTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
    End If
    
    If ExpireTime <> 0 Then
        astrReturn.Add "Expire Time: " & DateFormat(ExpireTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
    End If
    
    strCondition = ConditionText
    If Len(strCondition) > 0 Then
        astrReturn.Add "Condition: " & strCondition
    End If
    
    If m.dTrailAmount <> 0# Then
        If Parse(m.strTrailOptions, ",", 1) = "0" Then
            astrReturn.Add "Trail By: " & Format(m.dTrailAmount, "$#,##0.00")
        Else
            astrReturn.Add "Trail By: " & m.Bars.PriceDisplay(m.dTrailAmount)
        End If
    End If
    
    strContingency = m.Contingency.ToString
    If Len(strContingency) > 0 Then
        astrReturn.Add "Contingency: " & strContingency
    End If
    
    If astrReturn.Size = 0 Then
        strReturn = "No Advanced Order Settings"
    Else
        strReturn = astrReturn.JoinFields(";")
    End If
    
    AdvancedOrderText = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.AdvancedOrderText"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConditionText
'' Description: Get the english text condition
'' Inputs:      None
'' Returns:     English Text String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConditionText() As String
On Error GoTo ErrSection:

    Dim astrCondOpts As New cGdArray    ' Array of condition options
    Dim strCondition As String          ' Condition english text
    Dim astrCondition As New cGdArray   ' Condition
    Dim strField As String              ' Field
    
    astrCondOpts.Create eGDARRAY_Strings
    astrCondition.Create eGDARRAY_Strings
    
    If Len(m.strConditionOptions) > 0 Then
        astrCondOpts.SplitFields m.strConditionOptions, vbTab
        
        If astrCondOpts(4) = "1" Then
            astrCondition.SplitFields astrCondOpts(5), ";"
            Select Case UCase(astrCondition(0))
                Case "SESSION OPEN"
                    strField = "Open"
                Case "SESSION HIGH"
                    strField = "High"
                Case "SESSION LOW"
                    strField = "Low"
                Case "LAST PRICE"
                    strField = "Close"
            End Select
            strCondition = strField & " Of " & Chr(34) & astrCondition(1) & ",Daily" & Chr(34)
            strCondition = strCondition & " " & astrCondition(2) & " " & astrCondition(3)
        End If
        
        If astrCondOpts(6) = "1" Then
            If Len(strCondition) > 0 Then
                strCondition = "(" & strCondition & ") AND (" & astrCondOpts(7) & ")"
            Else
                strCondition = astrCondOpts(7)
            End If
        End If
    End If
    
    ConditionText = strCondition

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.ConditionText"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EngineString
'' Description: Generate an engine style string from the order
'' Inputs:      Use Remaining Quantity?
'' Returns:     Engine String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EngineString(Optional ByVal bUseRemainingQuantity As Boolean = False) As String
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array to turn into a delimited string
    
    astrOrder.Create eGDARRAY_Strings
    astrOrder(0) = "  1"
    If Enter Then
        If Buy Then
            astrOrder(1) = "EL"
        Else
            astrOrder(1) = "ES"
        End If
    Else
        If Buy Then
            astrOrder(1) = "XS"
        Else
            astrOrder(1) = "XL"
        End If
    End If
    astrOrder(2) = "0"
    Select Case OrderType
        Case eTT_OrderType_Market
            astrOrder(3) = "M"
            astrOrder(4) = ""
            astrOrder(5) = ""
            astrOrder(6) = ""
            astrOrder(7) = ""
        Case eTT_OrderType_Stop
            astrOrder(3) = "S"
            astrOrder(4) = Str(StopPrice)
            astrOrder(5) = ""
            astrOrder(6) = ""
            astrOrder(7) = ""
        Case eTT_OrderType_Limit
            astrOrder(3) = "L"
            astrOrder(4) = Str(LimitPrice)
            astrOrder(5) = ""
            astrOrder(6) = ""
            astrOrder(7) = ""
        Case eTT_OrderType_MarketOnClose
            astrOrder(3) = "MOC"
            astrOrder(4) = "-999999"    ' DAJ 11/17/2014: The engine is putting a null price here
            astrOrder(5) = ""
            astrOrder(6) = ""
            astrOrder(7) = ""
        Case eTT_OrderType_StopCloseOnly
            astrOrder(3) = "SCO"
            astrOrder(4) = Str(StopPrice)
            astrOrder(5) = ""
            astrOrder(6) = ""
            astrOrder(7) = ""
        Case eTT_OrderType_LimitCloseOnly
            astrOrder(3) = "LCO"
            astrOrder(4) = Str(LimitPrice)
            astrOrder(5) = ""
            astrOrder(6) = ""
            astrOrder(7) = ""
        Case eTT_OrderType_StopWithLimit
            astrOrder(3) = "SWL"
            astrOrder(4) = Str(StopPrice)
            astrOrder(5) = ""
            astrOrder(6) = Str(LimitPrice)
            astrOrder(7) = ""
        Case eTT_OrderType_StopWithLimitCloseOnly
            astrOrder(3) = "SWLCO"
            astrOrder(4) = Str(StopPrice)
            astrOrder(5) = ""
            astrOrder(6) = Str(LimitPrice)
            astrOrder(7) = ""
        Case eTT_OrderType_MIT
            astrOrder(3) = "MIT"
            astrOrder(4) = Str(MitPrice)
            astrOrder(5) = ""
            astrOrder(6) = ""
            astrOrder(7) = ""
    End Select
    
    If bUseRemainingQuantity Then
        astrOrder(8) = Str(RemainingQuantity)
    Else
        astrOrder(8) = Str(Quantity)
    End If
    
    EngineString = astrOrder.JoinFields(vbTab)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.EngineString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderPriceHit
'' Description: Has the order price been hit?
'' Inputs:      Current Price
'' Returns:     True if makes order true, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderPriceHit(ByVal dPrice As Double) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    Select Case OrderType
        Case eTT_OrderType_Market
            'bReturn = True
            
        Case eTT_OrderType_Stop
            If Buy Then
                If dPrice >= StopPrice Then bReturn = True
            Else
                If dPrice <= StopPrice Then bReturn = True
            End If
            
        Case eTT_OrderType_Limit
            If Buy Then
                If dPrice <= LimitPrice Then bReturn = True
            Else
                If dPrice >= LimitPrice Then bReturn = True
            End If
            
        Case eTT_OrderType_StopWithLimit
            If Buy Then
                If dPrice >= StopPrice And dPrice <= LimitPrice Then bReturn = True
            Else
                If dPrice <= StopPrice And dPrice >= LimitPrice Then bReturn = True
            End If
            
        Case eTT_OrderType_MIT
            If Buy Then
                If dPrice <= MitPrice Then bReturn = True
            Else
                If dPrice >= MitPrice Then bReturn = True
            End If
    
    End Select
    
    OrderPriceHit = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.OrderPriceHit"
    
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
    
    Set m.OrderLegs = mCopyFrom.OrderLegs.MakeCopy
    Set m.Fills = mCopyFrom.Fills.MakeCopy
    Set m.History = mCopyFrom.History.MakeCopy
    Set m.Bars = mCopyFrom.Bars.MakeCopy
    Set m.SecondaryMarkets = mCopyFrom.SecondaryMarkets.MakeCopy
    Set m.Contingency = mCopyFrom.Contingency.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the order
'' Inputs:      None
'' Returns:     Copy of the order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cPtOrder
On Error GoTo ErrSection:

    Dim aCopy As New cPtOrder           ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.MakeCopy"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConditionConfirm
'' Description: Is it necessary to confirm with the user on a condition?
'' Inputs:      None
'' Returns:     True if need to confirm, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConditionConfirm() As Boolean
On Error GoTo ErrSection:

    Dim bConditionConfirm As Boolean    ' Do we need to confirm on a condition?

    bConditionConfirm = False
    If Len(m.strConditionOptions) > 0 Then
        bConditionConfirm = (Val(Parse(m.strConditionOptions, vbTab, 1)) <> 0)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.ConditionConfirm"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TriggerTime
'' Description: Get the trigger time for the order
'' Inputs:      None
'' Returns:     Trigger Time or Zero if none
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TriggerTime() As Double
On Error GoTo ErrSection:

    Dim dTriggerTime As Double          ' Trigger Time

    dTriggerTime = 0#
    If Len(m.strConditionOptions) > 0 Then
        dTriggerTime = Val(Parse(m.strConditionOptions, vbTab, 2))
    End If
    TriggerTime = dTriggerTime

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.TriggerTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExpireTime
'' Description: Get the expire time for the order
'' Inputs:      None
'' Returns:     Expire Time or Zero if none
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExpireTime() As Double
On Error GoTo ErrSection:

    Dim dExpireTime As Double           ' Expire Time

    dExpireTime = 0#
    If Len(m.strConditionOptions) > 0 Then
        dExpireTime = Val(Parse(m.strConditionOptions, vbTab, 4))
    End If
    ExpireTime = dExpireTime

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.ExpireTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Condition
'' Description: Get the coded text condition to test
'' Inputs:      None
'' Returns:     Coded Text String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Condition() As String
On Error GoTo ErrSection:

    Dim astrCondOpts As New cGdArray    ' Array of condition options
    Dim strCondition As String          ' Condition english text
    Dim astrCondition As New cGdArray   ' Condition
    Dim strField As String              ' Field
    Dim Expr As New cExpression         ' Expression to get coded text from
    Dim astrBarNames As New cGdArray    ' Array of bar names
    Dim astrExpressions As New cGdArray ' Array of expressions
    Dim lIndex As Long                  ' Index into a for loop
    
    If Len(m.strCondition) = 0 Then
        astrCondOpts.Create eGDARRAY_Strings
        astrCondition.Create eGDARRAY_Strings
        
        If Len(m.strConditionOptions) > 0 Then
            astrCondOpts.SplitFields m.strConditionOptions, vbTab
            
            If astrCondOpts(4) = "1" Then
                astrCondition.SplitFields astrCondOpts(5), ";"
                Select Case UCase(astrCondition(0))
                    Case "SESSION OPEN"
                        strField = "Open"
                    Case "SESSION HIGH"
                        strField = "High"
                    Case "SESSION LOW"
                        strField = "Low"
                    Case "LAST PRICE"
                        strField = "Close"
                End Select
                m.strConditionSymbol = astrCondition(1)
                strCondition = strField & " Of " & Chr(34) & m.strConditionSymbol & ",Daily" & Chr(34)
                strCondition = strCondition & " " & astrCondition(2) & " " & astrCondition(3)
            Else
                m.strConditionSymbol = ""
            End If
            
            If astrCondOpts(6) = "1" Then
                If Len(strCondition) > 0 Then
                    strCondition = "(" & strCondition & ") AND (" & astrCondOpts(7) & ")"
                Else
                    strCondition = astrCondOpts(7)
                End If
            End If
            
            If Len(strCondition) > 0 Then
                With Expr
                    .PortfolioNavigator = False
                    .Functions = g.Functions
                    .ValidateFunctionRule strCondition
                    m.strCondition = .CodedText
                End With
                
                astrExpressions(0) = m.strCondition
                
                astrBarNames.Add "Market1"
                astrBarNames.Add "Daily"
                astrBarNames.Add "Weekly"
                astrBarNames.Add "Monthly"
        
                MarketsInExpressions astrExpressions, 0, False, astrBarNames, Nothing, astrCondOpts(8), Symbol
                LoadAllMarkets
            End If
        End If
    End If
    
    Condition = m.strCondition

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.Condition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunExpression
'' Description: Run the expression
'' Inputs:      None
'' Returns:     True if evaluated to True, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RunExpression() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrExpr As New cGdArray        ' Array of expressions to hand to engine
    Dim astrParms As New cGdArray       ' Array of parameters to hand to engine
    Dim astrBarNames As New cGdArray    ' Array of bar names to hand to engine
    Dim alResults As New cGdArray       ' Array of results back from engine
    Dim alBarHandles As New cGdArray    ' Array of bar handles to hand to engine
    Dim SecondaryMarkets As New cGdTree ' Secondary markets collection
    Dim Bars As New cGdBars             ' Bars object
    Dim Daily As New cGdBars            ' Daily version of the Bars object
    Dim Weekly As New cGdBars           ' Weekly version of the Bars object
    Dim Monthly As New cGdBars          ' Monthly version of the Bars object
    Dim strCodedText As String          ' Coded text expression
    Dim dStartDate As Double            ' Starting date to load data
    Dim hArray As Long                  ' Handle for an array
    Dim lBars As Long                   ' Index into a for loop
    Dim rc As Long                      ' Return code from a function call
    Dim dValue As Double                ' Value from the result array
    Dim lIndex As Long                  ' Index into a for loop
    Dim dNow As Double                  ' Current tick count
    Dim DumpBars As cGdBars             ' Bars to dump
    Dim strData As String               ' Data to dump
    Dim strBarName As String            ' Bar name
    Static bRunExpression As Boolean    ' Previous value of run expression
    Static dLastDump As Double          ' Last time we dumped the status
        
    bReturn = False
    strCodedText = Condition
    If Len(strCodedText) > 0 Then
        ' Create the arrays...
        astrExpr.Create eGDARRAY_Strings
        astrParms.Create eGDARRAY_Strings
        astrBarNames.Create eGDARRAY_Strings
        alResults.Create eGDARRAY_Longs
        alBarHandles.Create eGDARRAY_Longs
        
        If m.Bars.Size > 0 Then
            astrExpr.Add strCodedText
            dStartDate = m.Bars(eBARS_DateTime, 0)
            
            astrBarNames.Add "Market1"
            astrBarNames.Add "Daily"
            astrBarNames.Add "Weekly"
            astrBarNames.Add "Monthly"
            
            MarketsInExpressions astrExpr, dStartDate, True, astrBarNames, Nothing, "Daily", Symbol
            Daily.BuildBars "Daily", m.Bars.BarsHandle
            Weekly.BuildBars "Weekly", m.Bars.BarsHandle
            Monthly.BuildBars "Monthly", m.Bars.BarsHandle
        
            hArray = gdCreateArray(eGDARRAY_Doubles, Bars.Size)
            alResults.Add hArray
            
            astrParms(0) = "Order" & Str(m.lOrderID)
            If Not SetupExpressions(astrParms, astrBarNames, astrExpr) Then
                InfBox "An error exists in the Order Condition expression", "[]", , "Order Condition"
                Exit Function
            End If
        
            alBarHandles.Num(0) = m.Bars.BarsHandle
            alBarHandles.Num(1) = Daily.BarsHandle
            alBarHandles.Num(2) = Weekly.BarsHandle
            alBarHandles.Num(3) = Monthly.BarsHandle
            
            For lBars = 5 To m.SecondaryMarkets.Count
                alBarHandles.Num(lBars - 1) = m.SecondaryMarkets(lBars).BarsHandle
            Next lBars
            astrParms.Size = 1
        
            rc = RunExpressions(astrParms.ArrayHandle, astrBarNames.ArrayHandle, alBarHandles.ArrayHandle, alResults.ArrayHandle, ByVal 0&, ByVal 0&)
            If rc = 0 Then
                hArray = alResults.Num(0)
                dValue = gdGetNum(hArray, gdGetSize(hArray) - 1)
                If dValue <> gdNullValue(hArray) Then
                    ' if so, find first non-null item
                    For lIndex = gdGetSize(hArray) - 1 To 0 Step -1
                        dValue = gdGetNum(hArray, lIndex)
                        If dValue <> gdNullValue(hArray) Then
                            If dValue <> 0 Then
                                bReturn = True
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If
            
            ' destroy all the result arrays that got created with gdCreateArray
            For lIndex = 0 To alResults.Size - 1
                gdDestroyArray alResults(lIndex)
            Next
        End If
        
        SetupExpressions astrParms
    
        If bReturn <> bRunExpression Then
            g.Broker.BrokerDebug m.nBroker, vbTab & "RunExpression changed from " & Str(bRunExpression) & " to " & Str(bReturn) & ": " & OrderText(True, True, True) & " ( " & m.strConditionOptions & " )"
            bRunExpression = bReturn
        Else
            dNow = gdTickCount
            If dNow > (dLastDump + (5# * 60# * 1000#)) Then
                Set DumpBars = m.Bars
                strData = " ( " & DumpBars.Prop(eBARS_Symbol) & ";" & DumpBars.Prop(eBARS_PeriodicityStr) & ";Market1 -  o:" & Str(DumpBars(eBARS_Open, DumpBars.Size - 1)) & "; h:" & Str(DumpBars(eBARS_High, DumpBars.Size - 1)) & "; l:" & Str(DumpBars(eBARS_Low, DumpBars.Size - 1)) & "; c:" & Str(DumpBars(eBARS_Close, DumpBars.Size - 1)) & "; handle:" & Str(DumpBars.BarsHandle) & " )"
                
                If Len(m.strConditionSymbol) > 0 Then
                    strBarName = Chr(34) & m.strConditionSymbol & ",Daily" & Chr(34)
                    For lIndex = 0 To astrBarNames.Size - 1
                        If astrBarNames(lIndex) = UCase(strBarName) Then
                            Set DumpBars = m.SecondaryMarkets(lIndex + 1)
                            strData = strData & " ( " & strBarName & " - o:" & Str(DumpBars(eBARS_Open, DumpBars.Size - 1)) & "; h:" & Str(DumpBars(eBARS_High, DumpBars.Size - 1)) & "; l:" & Str(DumpBars(eBARS_Low, DumpBars.Size - 1)) & "; c:" & Str(DumpBars(eBARS_Close, DumpBars.Size - 1)) & "; handle:" & Str(DumpBars.BarsHandle) & " )"
                            
                            Exit For
                        End If
                    Next lIndex
                ElseIf m.SecondaryMarkets.Count > 4 Then
                    For lIndex = 5 To m.SecondaryMarkets.Count
                        strBarName = astrBarNames(lIndex - 1)
                        Set DumpBars = m.SecondaryMarkets(lIndex)
                        strData = strData & " (" & DumpBars.Prop(eBARS_Symbol) & ";" & DumpBars.Prop(eBARS_PeriodicityStr) & ";" & strBarName & " - c:" & Str(DumpBars(eBARS_Close, DumpBars.Size - 1)) & "; handle:" & Str(DumpBars.BarsHandle) & " )"
                    Next lIndex
                End If
                
                g.Broker.BrokerDebug m.nBroker, vbTab & "RunExpression = " & Str(bReturn) & ": " & OrderText(True, True, True) & strData
                dLastDump = dNow
            End If
        End If
    End If
    
    RunExpression = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.RunExpression"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBars
'' Description: Load up the bars
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadBars(Bars As cGdBars, ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String)
On Error GoTo ErrSection:

    Dim lNumBars As Long                ' Number of bars required
    Dim strNumBars As String            ' Number of bar information
    
    strNumBars = Parse(m.strConditionOptions, vbTab, 10)
    If Len(strNumBars) > 0 Then
        If CBool(Parse(strNumBars, "|", 1)) = True Then
            lNumBars = CLng(Val(Parse(strNumBars, "|", 3))) * -1&
        Else
            lNumBars = CLng(Val(Parse(strNumBars, "|", 2))) * -1&
        End If
    Else
        lNumBars = 0&
    End If

    DM_GetBars Bars, vSymbolOrSymbolID, strPeriod, lNumBars
    g.RealTime.AddTickBuffer Bars
    g.RealTime.SpliceBars Bars

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.LoadBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllRtDataAvailable
'' Description: Is all of the streaming data available?
'' Inputs:      None
'' Returns:     True if all data available, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllRtDataAvailable() As Boolean
On Error GoTo ErrSection:

    Dim bAvailable As Boolean           ' Bars available?
    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim bError As Boolean               ' Did the data request error?
    Dim strMessage As String            ' Message for the order
    Static bAlreadyLogged As Boolean    ' Have we already logged information?
    Static bLastKnown As Boolean        ' Last known value
    
    bReturn = False
    If HasExpression = False Then
        bReturn = True
    ElseIf Len(Condition) > 0 Then
        bAvailable = g.RealTime.RtDataAvailableForBars(m.Bars, bError)
        
        If bAvailable Then
            bReturn = True
            For lIndex = 5 To m.SecondaryMarkets.Count
                bAvailable = g.RealTime.RtDataAvailableForBars(m.SecondaryMarkets(lIndex), bError)
                
                If bAvailable = False Then
                    bReturn = False
                    
                    If bError = True Then
                        strMessage = "Error retrieving data for '" & m.SecondaryMarkets(lIndex).Prop(eBARS_Symbol) & "' from the server"
                        g.Broker.BrokerDebug m.nBroker, strMessage
                        Message = strMessage
                        ChangeOrderStatus eTT_OrderStatus_Error
                        
                        ShowAlertPopup "Order failed:||" & OrderText(True, True, False) & "||" & strMessage, "Order Failed", vbLeftJustify
                    End If
                    
                    Exit For
                End If
            Next lIndex
        Else
            bReturn = False
            If bError = True Then
                strMessage = "Error retrieving data for '" & m.Bars.Prop(eBARS_Symbol) & "' from the server"
                g.Broker.BrokerDebug m.nBroker, strMessage
                Message = strMessage
                ChangeOrderStatus eTT_OrderStatus_Error
                        
                ShowAlertPopup "Order failed:||" & OrderText(True, True, False) & "||" & strMessage, "Order Failed", vbLeftJustify
            End If
        End If
        
        If (bAlreadyLogged = False) Or (bReturn <> bLastKnown) Then
            g.Broker.BrokerDebug m.nBroker, "All Rt Data for Conditional Order ( " & Trim(ConditionText) & " ) = " & Str(bReturn)
            
            bAlreadyLogged = True
            bLastKnown = bReturn
        End If
    End If
        
    AllRtDataAvailable = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.AllRtDataAvailable"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBars
'' Description: Update the bars realtime
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateBars()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bNewBar As Boolean              ' Do we have a new bar?
    
    For lIndex = 5 To m.SecondaryMarkets.Count
        If g.RealTime.UpdateBars(m.SecondaryMarkets(lIndex), bNewBar) = True Then
            If bNewBar = True Then
                LoadBars m.SecondaryMarkets(lIndex), m.SecondaryMarkets(lIndex).SymbolOrSymbolID, m.SecondaryMarkets(lIndex).Prop(eBARS_PeriodicityStr)
            End If
        End If
    Next lIndex
    
    If g.RealTime.UpdateBars(m.Bars, bNewBar) = True Then
        If bNewBar = True Then
            LoadBars m.Bars, m.Bars.SymbolOrSymbolID, m.Bars.Prop(eBARS_PeriodicityStr)
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.UpdateBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAllMarkets
'' Description: Load up all of the markets
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadAllMarkets()
On Error GoTo ErrSection:

    Dim strCodedText As String          ' Coded text for the expression
    Dim strPeriod As String             ' Default Period for the bars
    Dim astrExpr As New cGdArray        ' Array of expressions
    Dim astrBarNames As New cGdArray    ' Bar names
    Dim Daily As New cGdBars            ' Temporary bars object
    Dim Weekly As New cGdBars           ' Temporary bars object
    Dim Monthly As New cGdBars          ' Temporary bars object
    Dim lIndex As Long                  ' Index into a for loop
    Dim dStartDate As Double            ' Start date for loading the data
    
    strCodedText = Condition
    strPeriod = Parse(m.strConditionOptions, vbTab, 9)
    If Len(strPeriod) = 0 Then strPeriod = "Daily"
    
    If Len(strCodedText) > 0 Then
        LoadBars m.Bars, SymbolOrSymbolID, strPeriod
        
        Daily.BuildBars "Daily", m.Bars.BarsHandle
        Weekly.BuildBars "Weekly", Daily.BarsHandle
        Monthly.BuildBars "Monthly", Weekly.BarsHandle
        
        Set m.SecondaryMarkets = New cGdTree
        m.SecondaryMarkets.Add m.Bars
        m.SecondaryMarkets.Add Daily
        m.SecondaryMarkets.Add Weekly
        m.SecondaryMarkets.Add Monthly
        
        astrBarNames.Add "Market1"
        astrBarNames.Add "Daily"
        astrBarNames.Add "Weekly"
        astrBarNames.Add "Monthly"
        
        dStartDate = m.Bars(eBARS_DateTime, 0)
        If dStartDate = kNullData Then
            dStartDate = 0#
        End If
        
        astrExpr.Add strCodedText
        MarketsInExpressions astrExpr, dStartDate, True, astrBarNames, m.SecondaryMarkets, strPeriod, Symbol, , False
        For lIndex = 5 To m.SecondaryMarkets.Count
            LoadBars m.SecondaryMarkets(lIndex), m.SecondaryMarkets(lIndex).Prop(eBARS_Symbol), m.SecondaryMarkets(lIndex).Prop(eBARS_PeriodicityStr)
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.LoadAllMarkets"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsConditional
'' Description: Determine whether the order has a condition
'' Inputs:      Include Expire?
'' Returns:     True if Conditional, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsConditional(ByVal bIncludeExpire As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim astrCondition As cGdArray       ' Conditional options broken into an array
    Dim bReturn As Boolean              ' Return value for the function
    
    ' Trigger Time On, Trigger Time, Expire Time On, Expire Time, Price Condition On,
    ' Price Condition, TradeSense On, TradeSense Condition, TradeSense Period
    
    Set astrCondition = New cGdArray
    astrCondition.Create eGDARRAY_Strings
    astrCondition.SplitFields m.strConditionOptions, vbTab
    
    bReturn = False
    If (astrCondition(0) = "1") And (Len(astrCondition(1)) > 0) Then
        bReturn = True
    ElseIf (bIncludeExpire = True) And ((astrCondition(2) = "1") And (Len(astrCondition(3)) > 0)) Then
        bReturn = True
    ElseIf (astrCondition(4) = "1") And (Len(astrCondition(5)) > 0) Then
        bReturn = True
    ElseIf (astrCondition(6) = "1") And (Len(astrCondition(7)) > 0) Then
        bReturn = True
    End If

    IsConditional = bReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cPtOrder.IsConditional"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasExpression
'' Description: Does the order have a conditional expression turned on?
'' Inputs:      None
'' Returns:     True if has expression, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasExpression() As Boolean
On Error GoTo ErrSection:

    Dim astrCondition As cGdArray       ' Conditional options broken into an array
    Dim bReturn As Boolean              ' Return value for the function
    
    ' Trigger Time On, Trigger Time, Expire Time On, Expire Time, Price Condition On,
    ' Price Condition, TradeSense On, TradeSense Condition, TradeSense Period
    
    Set astrCondition = New cGdArray
    astrCondition.Create eGDARRAY_Strings
    astrCondition.SplitFields m.strConditionOptions, vbTab
    
    bReturn = False
    If (astrCondition(4) = "1") And (Len(astrCondition(5)) > 0) Then
        bReturn = True
    ElseIf (astrCondition(6) = "1") And (Len(astrCondition(7)) > 0) Then
        bReturn = True
    End If

    HasExpression = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.HasExpression"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasTrigger
'' Description: Determine whether the order is pending a trigger
'' Inputs:      None
'' Returns:     True if Trigger, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasTrigger() As Boolean
On Error GoTo ErrSection:

    HasTrigger = (m.lTriggerOrderID <> 0) Or IsConditional(False)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.HasTrigger"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckTime
'' Description: Check to see if we are inside the proper times
'' Inputs:      None
'' Returns:     True if inside the times (or no times), False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckTime() As Boolean
On Error GoTo ErrSection:

    Dim dTimeNow As Double              ' Time right now
    Dim dTriggerTime As Double          ' Trigger time
    Dim bTriggerTimeHit As Boolean      ' Trigger time hit?
    Dim dExpireTime As Double           ' Expire time
    Dim bExpireTimeHit As Boolean       ' Expire time hit?
    Static bTriggerTime As Boolean      ' Last state of the trigger time check
    Static bExpireTime As Boolean       ' Last state of the expire time check
    
    dTimeNow = CurrentTime(, , True)
    
    bTriggerTimeHit = True
    dTriggerTime = TriggerTime
    If dTriggerTime <> 0 Then
        bTriggerTimeHit = (dTimeNow >= dTriggerTime)
        If bTriggerTimeHit <> bTriggerTime Then
            g.Broker.BrokerDebug m.nBroker, vbTab & "Trigger time ( " & DateFormat(dTriggerTime, MM_DD_YYYY, HH_MM_SS) & " ) hit changed from " & Str(bTriggerTime) & " to " & Str(bTriggerTimeHit) & ": " & OrderText(True, True, True)
            bTriggerTime = bTriggerTimeHit
        End If
    End If
    
    bExpireTimeHit = True
    dExpireTime = ExpireTime
    If dExpireTime <> 0 Then
        bExpireTimeHit = (dTimeNow >= dExpireTime)
        If bExpireTimeHit <> bExpireTime Then
            g.Broker.BrokerDebug m.nBroker, vbTab & "Expire time ( " & DateFormat(dExpireTime, MM_DD_YYYY, HH_MM_SS) & " ) hit changed from " & Str(bExpireTime) & " to " & Str(bExpireTimeHit) & ": " & OrderText(True, True, True)
            bExpireTime = bExpireTimeHit
        End If
    End If
    
    CheckTime = (bTriggerTimeHit And bExpireTimeHit)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.CheckTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TrailingStopValue
'' Description: Calculate the trailing stop value for this order given the
''              lowest low and highest high of the data
'' Inputs:      Lowest Low, Highest High
'' Returns:     Trailing Stop Value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TrailingStopValue(ByVal dLowestLow As Double, ByVal dHighestHigh As Double)
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim dNumTicks As Double             ' Number of ticks away from extreme
    
    dReturn = kNullData
    If Len(TrailOptions) > 0 Then
        ' Determine the number of ticks that the user wants to trail ( may be a dollar amount )...
        If Parse(TrailOptions, ",", 1) = "0" Then
            If m.Bars.Prop(eBARS_TickValue) <> 0 Then
                dNumTicks = (TrailAmount / m.Bars.Prop(eBARS_TickValue)) * m.Bars.Prop(eBARS_TickMove)
            End If
        Else
            dNumTicks = TrailAmount
        End If
    
        ' A Buy Stop trails the market the specified number of ticks above the lowest low...
        If Buy Then
            If dLowestLow <> kNullData Then
                dReturn = RoundToSigDigits(dLowestLow + dNumTicks)
            End If
            
        ' A Sell Stop trails the market the specified number of ticks below the highest high...
        Else
            If dHighestHigh <> kNullData Then
                dReturn = RoundToSigDigits(dHighestHigh - dNumTicks)
            End If
        End If
    End If
    
    TrailingStopValue = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.TrailingStopValue"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Property:    OrderPrice (Get)
'' Description: gets price of order
'' Inputs:      bStop to specify whether stop or limit price is wanted
''              for stop with limit order
'' Returns:     Returns stop or limit price based on type of order
''              or stop price if bStop is true for stop with limit order
''              or limit price if bStop is false for stop with limit order
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get OrderPrice(ByVal bStop As Boolean) As Double
On Error GoTo ErrSection:

    Dim dPrice As Double                ' Price to return from the function
    Dim dRelative As Double             ' Relative price to a triggering order
    Dim Order As New cPtOrder           ' Triggering order
    
    dPrice = kNullData
    dRelative = Val(Parse(m.strTriggerOptions, ",", 3))
    If (dRelative <> 0#) And (m.lTriggerOrderID <> 0&) Then
        If Order.Load(m.lTriggerOrderID) Then
            Select Case OrderType
                Case eTT_OrderType_Limit
                    If Buy Then
                        dPrice = Order.OrderPrice(bStop) - dRelative
                    Else
                        dPrice = Order.OrderPrice(bStop) + dRelative
                    End If
                Case eTT_OrderType_Stop
                    If Buy Then
                        dPrice = Order.OrderPrice(bStop) + dRelative
                    Else
                        dPrice = Order.OrderPrice(bStop) - dRelative
                    End If
                Case eTT_OrderType_MIT
                    If Buy Then
                        dPrice = Order.OrderPrice(bStop) - dRelative
                    Else
                        dPrice = Order.OrderPrice(bStop) + dRelative
                    End If
            End Select
        End If
    Else
        Select Case OrderType
            Case eTT_OrderType_Limit, eTT_OrderType_LimitCloseOnly
                dPrice = LimitPrice
            Case eTT_OrderType_Stop, eTT_OrderType_StopCloseOnly
                dPrice = StopPrice
            Case eTT_OrderType_StopWithLimit, eTT_OrderType_StopWithLimitCloseOnly
                If bStop Then
                    dPrice = StopPrice
                Else
                    dPrice = LimitPrice
                End If
            Case eTT_OrderType_MIT
                dPrice = MitPrice
        End Select
    End If
    
    OrderPrice = dPrice
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPtOrder.OrderPrice.Get"

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Property:    OrderPrice (Let)
'' Description: sets price of order
'' Inputs:      bStop to specify whether stop or limit price is to be set
''              for stop with limit order
'' Returns:     sets stop or limit price based on type of order
''              or sets stop price if bStop is true for stop with limit order
''              or sets limit price if bStop is false for stop with limit order
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let OrderPrice(ByVal bStop As Boolean, ByVal dPrice As Double)
On Error GoTo ErrSection:
    
    Dim dRelative As Double             ' Relative price to a triggering order
    Dim Order As New cPtOrder           ' Triggering order
    
    dRelative = Val(Parse(m.strTriggerOptions, ",", 3))
    If (dRelative <> 0#) And (m.lTriggerOrderID <> 0&) Then
        If Order.Load(m.lTriggerOrderID) Then
            m.strTriggerOptions = "1,0," & Str(Abs(dPrice - Order.OrderPrice(bStop)))
        End If
    Else
        Select Case OrderType
            Case eTT_OrderType_Limit, eTT_OrderType_LimitCloseOnly
                LimitPrice = dPrice
            Case eTT_OrderType_Stop, eTT_OrderType_StopCloseOnly
                StopPrice = dPrice
            Case eTT_OrderType_StopWithLimit, eTT_OrderType_StopWithLimitCloseOnly
                If bStop Then
                    StopPrice = dPrice
                Else
                    LimitPrice = dPrice
                End If
            Case eTT_OrderType_MIT
                MitPrice = dPrice
        End Select
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPtOrder.OrderPrice.Let"

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasOpenTriggerOrder
'' Description: Is this order triggered by another order that has not filled yet?
'' Inputs:      None
'' Returns:     True if has open trigger order, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasOpenTriggerOrder() As Boolean
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Triggering order
    
    HasOpenTriggerOrder = False
    If m.lTriggerOrderID <> 0 Then
        If Order.Load(m.lTriggerOrderID) Then
            If (Order.Status <> eTT_OrderStatus_Filled) And (Not ((Order.Status = eTT_OrderStatus_Partial) And (TriggerOnPartial = True))) Then
                HasOpenTriggerOrder = True
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.HasOpenTriggerOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemainingQuantity
'' Description: Determine the amount of the original quantity that is unfilled
'' Inputs:      Order Leg
'' Returns:     Remaining quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemainingQuantity(Optional ByVal lOrderLeg As Long = 1&) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function

#If 0 Then
    Dim lIndex As Long                  ' Index into a for loop
    Dim lFilledQuantity As Long         ' Quantity filled
    Dim vSymbolOrSymbolID As Variant    ' Symbol or Symbol ID for the leg
    
    vSymbolOrSymbolID = m.OrderLegs(lOrderLeg).SymbolOrSymbolID
    For lIndex = 1 To m.Fills.Count
        If m.Fills(lIndex).SymbolOrSymbolID = vSymbolOrSymbolID Then
            lFilledQuantity = lFilledQuantity + m.Fills(lIndex).Quantity
        End If
    Next lIndex

    lReturn = m.lQuantity - lFilledQuantity
#Else
    lReturn = m.lQuantity - FillQuantity(lOrderLeg)
#End If
    
    RemainingQuantity = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.RemainingQuantity"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearConditionOptions
'' Description: Clear out the condition options
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearConditionOptions()
On Error GoTo ErrSection:

    Dim astrCondition As cGdArray       ' Conditional options broken into an array
    
    ' Trigger Time On, Trigger Time, Expire Time On, Expire Time, Price Condition On,
    ' Price Condition, TradeSense On, TradeSense Condition, TradeSense Period
        
    If Len(m.strConditionOptions) > 0 Then
        Set astrCondition = New cGdArray
        astrCondition.Create eGDARRAY_Strings
        astrCondition.SplitFields m.strConditionOptions, vbTab
        
        ' Turn off everything but the expire time -- if it is turned on, we want it to
        ' stay on so that it can expire at the appropriate time...
        astrCondition(0) = "0"
        astrCondition(4) = "0"
        astrCondition(6) = "0"
        
        m.strConditionOptions = astrCondition.JoinFields(vbTab)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.ClearConditionOptions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConvertSymbol
'' Description: Convert a symbol or symbol ID to the non-synthetic version of
''              the symbol or symbol ID (and take the front month if it is a
''              continuous contract)
'' Inputs:      Symbol or Symbol ID, Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ConvertSymbol(ByVal vSymbolOrID As Variant, Optional ByVal dDate# = 0) As Variant
On Error GoTo ErrSection:
       
    ' if a synthetic or continuous symbol, then convert it
    If dDate = 0 Then
        dDate = m.OrderDate
        If dDate = 0 Then dDate = Date
    End If
    ConvertSymbol = ConvertToTradeSymbol(vSymbolOrID, dDate)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.ConvertSymbol"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsPrimaryOrder
'' Description: Determine based on the order status if this is a primary order
'' Inputs:      None
'' Returns:     True if Primary Order, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get IsPrimaryOrder() As Boolean
On Error GoTo ErrSection:
    
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If IsOpenOrder(m.Status) = True Then
        If (m.Status <> eTT_OrderStatus_Parked) And (m.Status <> eTT_OrderStatus_TriggerPending) And (m.Status <> eTT_OrderStatus_DataPending) Then
            bReturn = True
        End If
    End If
    
    IsPrimaryOrder = bReturn

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.IsPrimaryOrder"

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerInfoString
'' Description: Convert the order into a broker info string
'' Inputs:      None
'' Returns:     Broker Info String
''
'' Fields:      Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date, At ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerInfoString() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of order information to join together
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    astrReturn(0) = m.strBrokerID
    astrReturn(1) = m.strGenesisOrderID
    astrReturn(2) = ""
    astrReturn(3) = g.Broker.AccountNumberForID(m.lAccountID)
    astrReturn(4) = Symbol
    astrReturn(5) = Str(OrderType)
    astrReturn(6) = Str(CLng(Buy))
    astrReturn(7) = Str(Quantity)
    astrReturn(8) = Str(LimitPrice)
    astrReturn(9) = Str(StopPrice)
    astrReturn(10) = Str(m.lExpiration)
    astrReturn(11) = Str(m.Status)
    astrReturn(12) = Str(RemainingQuantity)
    astrReturn(13) = Str(m.dStatusDate)
    astrReturn(14) = Str(m.lAutoTradeItemID)
    
    BrokerInfoString = astrReturn.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.BrokerInfoString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CopyAdvancedInfo
'' Description: Copy advanced order information from the given order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CopyAdvancedInfo(ByVal OrderFrom As cPtOrder)
On Error GoTo ErrSection:

    m.lCancelOrderID = OrderFrom.CancelOrderID
    m.lTriggerOrderID = OrderFrom.TriggerOrderID
    m.strTriggerOptions = OrderFrom.TriggerOptions
    m.strConditionOptions = OrderFrom.ConditionOptions
    m.dTrailAmount = OrderFrom.TrailAmount
    m.strTrailOptions = OrderFrom.TrailOptions
    Set m.Contingency = OrderFrom.Contingency.MakeCopy

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.CopyAdvancedInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsExpired
'' Description: Determine whether the order should be expired or not
'' Inputs:      None
'' Returns:     True if Expired, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsExpired() As Boolean
On Error GoTo ErrSection:

    Dim lSessionDate As Long            ' Session date for the current time
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    lSessionDate = m.Bars.SessionDateForTradeTime(CurrentTime(m.Bars.Prop(eBARS_ExchangeTimeZoneInf)))
    
    ' If this is a day order, then check the current session date against the session date for
    ' the order date...
    If m.lExpiration < 0 Then
        bReturn = (lSessionDate > m.Bars.SessionDateForTradeTime(ConvertBrokerDate(m.OrderDate, m.nBroker, Symbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))))
        
    ' Otherwise, if this is a GTD, then check the current session date against the given expiration date...
    ElseIf m.lExpiration > 0 Then
        bReturn = (lSessionDate > m.lExpiration)
    End If
    
    IsExpired = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.IsExpired"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    WrongSideOfMarket
'' Description: Is this order on the wrong side of the market?
'' Inputs:      Current Price
'' Returns:     True if on wrong side of the market, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WrongSideOfMarket(Optional ByVal dCurrentPrice As Double = kNullData) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    ' If no price is passed in, then if this order is triggered by another working
    ' order, use that order price, otherwise use the last known price for the
    ' symbol...
    If dCurrentPrice = kNullData Then
        dCurrentPrice = TriggeredByPrice
        If dCurrentPrice = kNullData Then
            dCurrentPrice = g.RealTime.LastKnownPrice(SymbolOrSymbolID)
        End If
    End If

    Select Case OrderType
        Case eTT_OrderType_Market, eTT_OrderType_MarketOnClose
            bReturn = False
            
        Case eTT_OrderType_Stop, eTT_OrderType_StopWithLimit, eTT_OrderType_StopCloseOnly, eTT_OrderType_StopWithLimitCloseOnly
            If Buy = True Then
                bReturn = (StopPrice < dCurrentPrice)
            Else
                bReturn = (StopPrice > dCurrentPrice)
            End If
        
        Case eTT_OrderType_Limit, eTT_OrderType_LimitCloseOnly
            If Buy = True Then
                bReturn = (LimitPrice > dCurrentPrice)
            Else
                bReturn = (LimitPrice < dCurrentPrice)
            End If
    
    End Select
    
    WrongSideOfMarket = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.WrongSideOfMarket"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TriggeredByPrice
'' Description: Price of the triggering order if applicable
'' Inputs:      None
'' Returns:     Triggering Order price (Null if not applicable)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TriggeredByPrice() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim TOrder As New cPtOrder          ' Triggering order
    
    dReturn = kNullData
    If TriggerOrderID <> 0 Then
        If TOrder.Load(TriggerOrderID) Then
            If IsOpenOrder(TOrder.Status) Then
                dReturn = TOrder.OrderPrice(True)
            End If
        End If
    End If
    
    TriggeredByPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.TriggeredByPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasTriggeredOrders
'' Description: Does this order have orders that are triggered by it?
'' Inputs:      None
'' Returns:     True if triggered orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasTriggeredOrders() As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim bReturn As Boolean              ' Return value for the function
    Dim Order As cPtOrder               ' Triggered order
    
    bReturn = False
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE Abs([TriggerOrderID])=" & Str(m.lOrderID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        Set Order = New cPtOrder
        If Order.Load(rs!OrderID, rs) Then
            If IsOpenOrder(Order.Status) Then
                bReturn = True
                Exit Do
            End If
        End If
        
        rs.MoveNext
    Loop
    
    HasTriggeredOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.HasTriggeredOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasContingencyOrders
'' Description: Does this order have contingency orders?
'' Inputs:      None
'' Returns:     True if has contingency orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasContingencyOrders() As Boolean
On Error GoTo ErrSection:

    HasContingencyOrders = (m.Contingency.ProfitOrderId > 0&) Or (m.Contingency.StopOrderId > 0&)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.HasContingencyOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasOcoOrders
'' Description: Does this order have an order OCO'ed with it?
'' Inputs:      None
'' Returns:     True if OCO link, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasOcoOrders(Optional ByVal bIncludeBrokerOco As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim bReturn As Boolean              ' Return value for the function
    Dim Order As cPtOrder               ' Triggered order
    
    bReturn = False
    If bIncludeBrokerOco Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [CancelOrderID]=" & Str(m.lOrderID) & " OR [BrokerCancelOrderID]=" & Str(m.lOrderID) & ";", dbOpenDynaset)
    Else
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [CancelOrderID]=" & Str(m.lOrderID) & ";", dbOpenDynaset)
    End If
    Do While Not rs.EOF
        Set Order = New cPtOrder
        If Order.Load(rs!OrderID, rs) Then
            If IsOpenOrder(Order.Status) Then
                bReturn = True
                Exit Do
            End If
        End If
        
        rs.MoveNext
    Loop
    
    HasOcoOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.HasTriggeredOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateContingentOrder
'' Description: Update the contingent order information
'' Inputs:      Order
'' Returns:     True if Changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UpdateContingentOrder(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bChanged As Boolean             ' Did the contingency order information change?
    Dim dPointsAmount As Double         ' Point difference between orders

    bChanged = False
    If Order.OrderID = m.Contingency.ProfitOrderId Then
        If Order.Status = eTT_OrderStatus_TriggerPending Then
            If OrderType = eTT_OrderType_Limit Then
                dPointsAmount = Abs(LimitPrice - Order.LimitPrice)
            ElseIf (OrderType = eTT_OrderType_Stop) Or (OrderType = eTT_OrderType_StopWithLimit) Then
                dPointsAmount = Abs(StopPrice - Order.LimitPrice)
            End If
            
            If dPointsAmount <> m.Contingency.ProfitPointsAmount Then
                bChanged = True
                
                m.Contingency.ProfitPointsAmount = dPointsAmount
                m.Contingency.ProfitDollarAmount = dPointsAmount * m.Bars.TickValuePerMove
            End If
        Else
            m.Contingency.UseProfitTarget = False
            m.Contingency.ProfitOrderId = 0&
            bChanged = True
        End If
    ElseIf Order.OrderID = m.Contingency.StopOrderId Then
        If Order.Status = eTT_OrderStatus_TriggerPending Then
            If OrderType = eTT_OrderType_Limit Then
                dPointsAmount = Abs(LimitPrice - Order.StopPrice)
            ElseIf (OrderType = eTT_OrderType_Stop) Or (OrderType = eTT_OrderType_StopWithLimit) Then
                dPointsAmount = Abs(StopPrice - Order.StopPrice)
            End If
            
            If dPointsAmount <> m.Contingency.StopPointsAmount Then
                bChanged = True
                
                m.Contingency.StopPointsAmount = dPointsAmount
                m.Contingency.StopDollarAmount = dPointsAmount * m.Bars.TickValuePerMove
            End If
        Else
            m.Contingency.UseStopLoss = False
            m.Contingency.StopOrderId = 0&
            bChanged = True
        End If
    End If
    
    If bChanged Then
        Save False
    End If
    
    UpdateContingentOrder = bChanged

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.UpdateContingentOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildSpreadSymbol
'' Description: Build the spread symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildSpreadSymbol()
On Error GoTo ErrSection:

    Dim astrLegs As cGdArray            ' Leg information
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrLegs = New cGdArray
    astrLegs.Create eGDARRAY_Strings, m.OrderLegs.Count
    
    For lIndex = 1 To m.OrderLegs.Count
        astrLegs(lIndex - 1) = m.OrderLegs(lIndex).LegInfo
    Next lIndex
    
    m.strSpreadSymbol = astrLegs.JoinFields("|")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.BuildSpreadSymbol"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateOCO
'' Description: Update the other order in an OCO situation as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateOCO()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lOldCancelID As Long            ' Old Cancel Order ID
    Dim Order As New cPtOrder           ' Order on the other side of the OCO

    If m.lCancelOrderID <> 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [OrderID]=" & Str(m.lCancelOrderID) & ";", dbOpenDynaset)
        If Not (rs.EOF And rs.BOF) Then
            lOldCancelID = rs!CancelOrderID
            If lOldCancelID <> m.lOrderID Then
                g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "( " & OrderIdText & " ): OCO order was linked to order " & Str(rs!CancelOrderID)
                
                If Order.Load(rs!OrderID) Then
                    Order.CancelOrderID = m.lOrderID
                    Order.Save
                    
                    g.Broker.AddOrder Order
                    OrderCallback Order
                End If
                
                If lOldCancelID <> 0 Then
                    Set Order = New cPtOrder
                    If Order.Load(lOldCancelID) Then
                        Order.CancelOrderID = 0&
                        Order.Save
                    
                        g.Broker.AddOrder Order
                        OrderCallback Order
                    End If
                End If
            ElseIf (m.Status = eTT_OrderStatus_Cancelled) Or (m.Status = eTT_OrderStatus_Expired) Or (m.Status = eTT_OrderStatus_Rejected) Then
                g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "( " & OrderIdText & " ): Reseting OCO on order " & Str(rs!OrderID) & " because this order is closed"
                
                Set Order = New cPtOrder
                If Order.Load(rs!OrderID) Then
                    Order.CancelOrderID = 0&
                    Order.Save
                    
                    g.Broker.AddOrder Order
                    OrderCallback Order
                End If
            End If
        End If
    Else
        ' If this is no longer set to cancel an order, make sure that the order that was to be
        ' cancelled by this order is also set to not cancel anymore...
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [CancelOrderID]=" & Str(m.lOrderID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "( " & OrderIdText & " ): Reseting OCO on order " & Str(rs!OrderID) & " because this order is no longer set to OCO"
            
            If Order.Load(rs!OrderID) Then
                Order.CancelOrderID = 0&
                Order.Save
                    
                g.Broker.AddOrder Order
                OrderCallback Order
            End If
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.UpdateOCO"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBrokerOCO
'' Description: Update the other order in an Broker held OCO situation as
''              appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateBrokerOCO()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lOldCancelID As Long            ' Old Cancel Order ID
    Dim Order As New cPtOrder           ' Order on the other side of the OCO

    If m.lBrokerCancelOrderID <> 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [OrderID]=" & Str(Abs(m.lBrokerCancelOrderID)) & ";", dbOpenDynaset)
        If Not (rs.EOF And rs.BOF) Then
            lOldCancelID = rs!BrokerCancelOrderID
            frmTest2.AddList "UpdateBrokerOCO(" & Str(m.lOrderID) & "): Old BrokerCancelOrderID for " & Str(m.lBrokerCancelOrderID) & " is " & Str(lOldCancelID)
            If lOldCancelID <> m.lOrderID Then
                If Order.Load(rs!OrderID) Then
                    If m.lBrokerCancelOrderID < 0 Then
                        Order.BrokerCancelOrderID = m.lOrderID * -1&
                    Else
                        Order.BrokerCancelOrderID = m.lOrderID
                    End If
                    
                    If Order.BrokerCancelOrderID <> lOldCancelID Then
                        Order.Save
                    
                        g.Broker.AddOrder Order
                        OrderCallback Order
                    End If
                ElseIf lOldCancelID <> 0 Then
                    Set Order = New cPtOrder
                    If Order.Load(Abs(lOldCancelID)) Then
                        Order.BrokerCancelOrderID = 0&
                        Order.Save
                    
                        g.Broker.AddOrder Order
                        OrderCallback Order
                    End If
                End If
            
            ' 12/09/2009 DAJ: Don't include Amened on this list because that case will take care of itself...
            ElseIf (m.Status = eTT_OrderStatus_Cancelled) Or (m.Status = eTT_OrderStatus_Expired) Or (m.Status = eTT_OrderStatus_Rejected) Then ' Or (m.Status = eTT_OrderStatus_Amended) Then
                Set Order = New cPtOrder
                If Order.Load(rs!OrderID) Then
                    Order.BrokerCancelOrderID = 0&
                    Order.Save
                    
                    g.Broker.AddOrder Order
                    OrderCallback Order
                End If
            End If
        End If
    Else
        ' If this is no longer set to cancel an order, make sure that the order that was to be
        ' cancelled by this order is also set to not cancel anymore...
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE ABS([BrokerCancelOrderID])=" & Str(m.lOrderID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            If Order.Load(rs!OrderID) Then
                Order.BrokerCancelOrderID = 0&
                Order.Save
                    
                g.Broker.AddOrder Order
                OrderCallback Order
            End If
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.UpdateBrokerOCO"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetSessionDate
'' Description: Set the session date for the order if not already set
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetSessionDate()
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol for the order

    If m.lSessionDate = 0& Then
        strSymbol = Symbol
        
        If Len(strSymbol) > 0 Then
            If m.dStatusDate = 0& Then
                m.lSessionDate = m.Bars.SessionDateForTradeTime(ConvertBrokerDate(m.OrderDate, m.nBroker, strSymbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf)))
            Else
                m.lSessionDate = m.Bars.SessionDateForTradeTime(ConvertBrokerDate(m.dStatusDate, m.nBroker, strSymbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf)))
            End If
            
            g.Broker.BrokerDebug m.nBroker, vbTab & "Session Date (" & Str(m.lOrderID) & ", '" & m.strGenesisOrderID & "', '" & m.strBrokerID & "'): " & DateFormat(m.lSessionDate, MM_DD_YYYY)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.SetSessionDate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitContingencyOrders
'' Description: Submit the contingency order(s) if applicable
'' Inputs:      None
'' Returns:     True if Order ID's set, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SubmitContingencyOrders() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim ProfitTargetOrder As cPtOrder   ' Profit target order
    Dim StopLossOrder As cPtOrder       ' Stop loss order
    Dim strAccountNumber As String      ' Account number
    Dim dAdjustment As Double           ' Adjustment to make on the price
    Dim bCreateProfitTarget As Boolean  ' Are we creating a profit target order?
    Dim bModifyProfitTarget As Boolean  ' Are we modifying the profit target order?
    Dim bCreateStopLoss As Boolean      ' Are we creating a stop loss order?
    Dim bModifyStopLoss As Boolean      ' Are we modifying the stop loss order?
    Dim strAction As String             ' Action for the order

    bReturn = False

    Set ProfitTargetOrder = Nothing
    bCreateProfitTarget = False
    bModifyProfitTarget = False
    
    Set StopLossOrder = Nothing
    bCreateStopLoss = False
    bModifyStopLoss = False

    strAccountNumber = g.Broker.AccountNumberForID(m.lAccountID)

    If m.Contingency.UseProfitTarget Then
        If Buy Then
            dAdjustment = m.Contingency.ProfitPointsAmount
        Else
            dAdjustment = m.Contingency.ProfitPointsAmount * -1#
        End If
        
        Set ProfitTargetOrder = New cPtOrder
        If m.Contingency.ProfitOrderId > 0& Then
            bModifyProfitTarget = True
            ProfitTargetOrder.Load m.Contingency.ProfitOrderId
            strAction = "Modified"
        Else
            bCreateProfitTarget = True
            strAction = "Created"
        End If
        
        With ProfitTargetOrder
            .AccountID = m.lAccountID
            If bCreateProfitTarget Then
                .GenesisOrderID = NextGenesisOrderID(strAccountNumber, m.nBroker)
            End If
            .AutoTradeItemID = m.lAutoTradeItemID
            .Buy = Not Buy
            .Quantity = m.lQuantity
            .SymbolOrSymbolID = SymbolOrSymbolID
            Select Case m.Contingency.ProfitTif
                Case eTT_TimeInForce_Day
                    .Expiration = -1&
                Case eTT_TimeInForce_GTC
                    .Expiration = 0&
            End Select
            .TriggerOrderID = m.lOrderID
            
            .OrderType = eTT_OrderType_Limit
            If OrderType = eTT_OrderType_Limit Then
                .LimitPrice = LimitPrice + dAdjustment
            ElseIf (OrderType = eTT_OrderType_Stop) Or (OrderType = eTT_OrderType_StopWithLimit) Then
                .LimitPrice = StopPrice + dAdjustment
            End If
            .StopPrice = kNullData
        End With
        
        g.Broker.BrokerDebug m.nBroker, vbTab & "Contingency Profit Target " & strAction & ": '" & ProfitTargetOrder.OrderText(True, True, True) & "'"
    Else
        If m.Contingency.ProfitOrderId > 0& Then
            g.Broker.BrokerDebug m.nBroker, vbTab & "Contingency Profit Target Cancelled because it was turned off"
            CancelOrderFromID m.Contingency.ProfitOrderId, "Submit Contingency"
            m.Contingency.ProfitOrderId = 0&
        End If
    End If
    
    If m.Contingency.UseStopLoss Then
        If Buy Then
            dAdjustment = m.Contingency.StopPointsAmount * -1#
        Else
            dAdjustment = m.Contingency.StopPointsAmount
        End If
        
        Set StopLossOrder = New cPtOrder
        If m.Contingency.StopOrderId > 0& Then
            bModifyStopLoss = True
            StopLossOrder.Load m.Contingency.StopOrderId
            strAction = "Modified"
        Else
            bCreateStopLoss = True
            strAction = "Created"
        End If
        
        With StopLossOrder
            .AccountID = m.lAccountID
            If bCreateStopLoss Then
                .GenesisOrderID = NextGenesisOrderID(strAccountNumber, m.nBroker)
            End If
            .AutoTradeItemID = m.lAutoTradeItemID
            .Buy = Not Buy
            .Quantity = m.lQuantity
            .SymbolOrSymbolID = SymbolOrSymbolID
            Select Case m.Contingency.StopTif
                Case eTT_TimeInForce_Day
                    .Expiration = -1&
                Case eTT_TimeInForce_GTC
                    .Expiration = 0&
            End Select
            .TriggerOrderID = m.lOrderID
            
            .OrderType = eTT_OrderType_Stop
            If OrderType = eTT_OrderType_Limit Then
                .StopPrice = LimitPrice + dAdjustment
            ElseIf (OrderType = eTT_OrderType_Stop) Or (OrderType = eTT_OrderType_StopWithLimit) Then
                .StopPrice = StopPrice + dAdjustment
            End If
            .LimitPrice = kNullData
        End With
        
        g.Broker.BrokerDebug m.nBroker, vbTab & "Contingency Stop Loss " & strAction & ": '" & StopLossOrder.OrderText(True, True, True) & "'"
    Else
        If m.Contingency.StopOrderId > 0& Then
            g.Broker.BrokerDebug m.nBroker, vbTab & "Contingency Stop Loss Cancelled because it was turned off"
            CancelOrderFromID m.Contingency.StopOrderId, "Submit Contingency"
            m.Contingency.StopOrderId = 0&
        End If
    End If
        
    If bCreateProfitTarget And bCreateStopLoss Then
        g.Broker.SubmitOrdersAsOco ProfitTargetOrder, StopLossOrder, m.lOrderID, True
        m.Contingency.ProfitOrderId = ProfitTargetOrder.OrderID
        m.Contingency.StopOrderId = StopLossOrder.OrderID
        bReturn = True
    Else
        If bCreateProfitTarget Then
            m.Contingency.ProfitOrderId = SubmitOrder(ProfitTargetOrder)
            bReturn = True
        ElseIf bModifyProfitTarget Then
            ProfitTargetOrder.ChangeOrderStatus eTT_OrderStatus_TriggerPending, , False
        End If
        
        If bCreateStopLoss Then
            m.Contingency.StopOrderId = SubmitOrder(StopLossOrder)
            bReturn = True
        ElseIf bModifyStopLoss Then
            StopLossOrder.ChangeOrderStatus eTT_OrderStatus_TriggerPending, , False
        End If
    End If
    
    g.Broker.BrokerDebug m.nBroker, vbTab & "Contingency Information (" & OrderIdText & ") after submitted: '" & m.Contingency.ToString & "'"
    
    SubmitContingencyOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.SubmitContingencyOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelContingencyOrders
'' Description: Cancel the contingency order(s) if applicable
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelContingencyOrders()
On Error GoTo ErrSection:

    If m.Contingency.ProfitOrderId > 0& Then
        g.Broker.BrokerDebug m.nBroker, vbTab & "Cancelling Contingency Profit Target because Order Status went to '" & OrderStatus(m.Status) & "'"
        CancelOrderFromID m.Contingency.ProfitOrderId, "Contingency Orders", False
    End If
    
    If m.Contingency.StopOrderId > 0& Then
        g.Broker.BrokerDebug m.nBroker, vbTab & "Cancelling Contingency Stop Loss because Order Status went to '" & OrderStatus(m.Status) & "'"
        CancelOrderFromID m.Contingency.StopOrderId, "Contingency Orders", False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.CancelContingencyOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIdText
'' Description: String with the various identifiers for this order
'' Inputs:      None
'' Returns:     Descriptive String of the Order Ids
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderIdText() As String
On Error GoTo ErrSection:

    OrderIdText = Str(m.lOrderID) & ", '" & m.strGenesisOrderID & "', '" & m.strBrokerID & "'"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtOrder.OrderIdText"
    
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

    Set m.OrderLegs = New cOrderLegs
    Set m.Fills = New cGdTree
    Set m.SpreadFills = New cGdTree
    Set m.History = New cGdTree
    Set m.SecondaryMarkets = New cGdTree
    Set m.PreviousOrder = Nothing
    Set m.Contingency = New cContingencyOrders
    
    Set m.Bars = New cGdBars
    
    ' Default a new order to a day order...
    m.lExpiration = -1&
    
    ' Default a new order to a non-auto exit order...
    m.bIsAutoExit = False
    
    ' Default is that the order was not created in an automated fashion...
    m.bIsAutomated = False
    
    m.lPreviousOrderID = 0&
    
    m.dStopPrice = kNullData
    m.dLimitPrice = kNullData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.Class_Initialize"
    
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

    Set m.OrderLegs = Nothing
    Set m.Fills = Nothing
    Set m.SpreadFills = Nothing
    Set m.History = Nothing
    Set m.SecondaryMarkets = Nothing
    Set m.PreviousOrder = Nothing
    Set m.Contingency = Nothing
    
    Set m.Bars = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtOrder.Class_Terminate"
    
End Sub

