VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerKeyValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerKeyValue.cls
'' Description: Class to manange key-value pair messages from stand-alone apps
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/25/2011   DAJ         Fixed calls to convert to Genesis price
'' 08/29/2011   DAJ         Fixed calls to convert from Genesis price
'' 09/07/2011   DAJ         Fixes to OrderToMessage, call Replaced orders Working
'' 09/20/2011   DAJ         Changed code in PositionFromMessage
'' 10/07/2011   DAJ         Handle a previous broker fill ID in a fill message
'' 10/25/2011   DAJ         Create account for order/fill/position if not exist
'' 12/02/2011   DAJ         Added direction to order message
'' 12/09/2011   DAJ         Multi-leg order support
'' 01/31/2012   DAJ         Keep dumping OrderIsEntry info to the log
'' 02/14/2012   DAJ         Added multi-leg order support
'' 04/12/2012   DAJ         Mods for broker view mode
'' 05/31/2012   DAJ         Turnkey implementation
'' 06/12/2012   DAJ         Added optional AccountID argument to "ToInfo" messages
'' 07/27/2012   DAJ         Use the FcmAccount as account nickname if passed in
'' 08/21/2012   DAJ         Added FcmAccountNumber for accounts
'' 09/11/2012   DAJ         Set PreviousBrokerID for order if applicable
'' 11/14/2012   DAJ         Fix for calculating session date for a fill
'' 10/16/2013   DAJ         Tweaks for Open E-Cry
'' 10/24/2013   DAJ         Handle 'Cancel succeeded' status from Open E-Cry
'' 11/14/2014   DAJ         Added support for MIT, On-Close, and On-Open orders
'' 09/14/2015   DAJ         Added Tradier
'' 04/07/2016   DAJ         Fixed Leg information in MultiLegOrderToMessage function
'' 04/12/2016   DAJ         Added DebitCredit flag and Underlying symbol to multi-leg order routines
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountFromMessage
'' Description: Fill an account object from the given message
'' Inputs:      Message, Broker, User Name, Password
'' Returns:     Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountFromMessage(ByVal brokerMessage As cBrokerMessage, ByVal Broker As cBroker, ByVal strUserName As String, ByVal strPassword As String) As cPtAccount
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account object to return
    
    Set Account = New cPtAccount
    With Account
        .AccountNumber = brokerMessage("Account")
        .AccountType = Broker.Broker
        .Broker = ""
        .ClosedProfit = 0#
        .Comms = 0#
        .CurrentBalance = Val(brokerMessage("Balance"))
        .FillRT = False
        .HoldOcoAtBroker = False
        If Len(brokerMessage("FcmAccount")) > 0 Then
            .Name = brokerMessage("FcmAccount")
        Else
            .Name = .AccountNumber
        End If
        .Password = strPassword
        .SecTypeMask = 31
        .StartingBalance = Val(brokerMessage("Balance"))
        .StartingDate = Date
        .UserName = strUserName
        .FcmAccountNumber = brokerMessage("FcmAccount")
    End With
    
    Set AccountFromMessage = Account

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.AccountFromMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromMessage
'' Description: Fill an order object from the given message
'' Inputs:      Message, Broker, Genesis symbol, Symbol information object, Account ID
'' Returns:     Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderFromMessage(ByVal brokerMessage As cBrokerMessage, ByVal Broker As cBroker, ByVal strGenesisSymbol As String, ByVal BrokerSym As cBrokerSymbol, Optional ByVal lAccountID As Long = kNullData) As cPtOrder
On Error GoTo ErrSection:

    Dim lQuantityFilled As Long         ' Quantity of the order that has filled
    Dim Bars As New cGdBars             ' Bars object
    Dim Order As cPtOrder               ' Order object to return
    
    Set Bars = New cGdBars
    SetBarProperties Bars, strGenesisSymbol
    
    If lAccountID = kNullData Then
        If (g.Broker.AccountExists(brokerMessage("Account")) = False) And (Len(brokerMessage("Account")) > 0) Then
            lAccountID = mTradeTracker.CreateAccountFromNumber(brokerMessage("Account"), Broker.Broker)
        Else
            lAccountID = g.Broker.AccountIDForNumber(brokerMessage("Account"))
        End If
    End If
    
    Set Order = New cPtOrder
    With Order
        .AccountID = lAccountID
        .BrokerID = brokerMessage("BrokerID")
        .Buy = (Left(UCase(brokerMessage("Side")), 3) = "BUY")
        .ExchangeID = brokerMessage("ExchangeID")
        .GenesisOrderID = brokerMessage("GenesisID")
        .LimitPrice = Broker.GenesisPrice(brokerMessage("LimitPrice"), BrokerSym.PriceMult)
        .Message = brokerMessage("Message")
        .OrderType = GenesisOrderType(brokerMessage("Type"), brokerMessage("Inst"))
        .Quantity = CLng(Val(brokerMessage("Quantity")))
        .Status = GenesisOrderStatus(brokerMessage("Status"))
        .StatusDate = DateToDouble(brokerMessage("StatusDate"))
        .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.StatusDate, Broker.Broker, strGenesisSymbol, False))
        .StopPrice = Broker.GenesisPrice(brokerMessage("StopPrice"), BrokerSym.PriceMult)
        .SymbolOrSymbolID = strGenesisSymbol
        lQuantityFilled = CLng(Val(brokerMessage("QuantityFilled")))
        
        Select Case GenesisTIF(brokerMessage("TIF"))
            Case eTT_TimeInForce_Day
                .Expiration = .SessionDate * -1&
            Case eTT_TimeInForce_GTC
                .Expiration = 0
            Case eTT_TimeInForce_GTD
                .Expiration = CLng(Val(brokerMessage("Expiration")))
        End Select
        
        If (Len(brokerMessage("PreviousBrokerID")) > 0) And (brokerMessage("PreviousBrokerID") <> brokerMessage("BrokerID")) Then
            .PreviousBrokerID = brokerMessage("PreviousBrokerID")
        End If
        
        If lQuantityFilled > 0 Then
            If (lQuantityFilled < .Quantity) And (UCase(brokerMessage("Status")) = "COMPLETED") Then
                .Status = eTT_OrderStatus_BalCancelled
            ElseIf (.Status = eTT_OrderStatus_Working) Then
                .Status = eTT_OrderStatus_Partial
            ElseIf (.Status = eTT_OrderStatus_Cancelled) Then
                .Status = eTT_OrderStatus_BalCancelled
            End If
        End If
    End With
    
    Set OrderFromMessage = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.OrderFromMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MultiLegOrderFromMessage
'' Description: Fill a multi-leg order object from the given message
'' Inputs:      Message, Broker, Genesis symbols, Symbol information object, Account ID
'' Returns:     Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MultiLegOrderFromMessage(ByVal brokerMessage As cBrokerMessage, ByVal Broker As cBroker, ByVal astrGenesisSymbols As cGdArray, ByVal BrokerSym As cBrokerSymbol, Optional ByVal lAccountID As Long = kNullData) As cPtOrder
On Error GoTo ErrSection:

    Dim lQuantityFilled As Long         ' Quantity of the order that has filled
    Dim Bars As New cGdBars             ' Bars object
    Dim Order As cPtOrder               ' Order object to return
    Dim lNumLegs As Long                ' Number of legs for the order
    Dim lIndex As Long                  ' Index into the for loop
    Dim OrderLeg As cOrderLeg           ' Order leg object
    
    Set Bars = New cGdBars
    SetBarProperties Bars, astrGenesisSymbols(0)
    
    If lAccountID = kNullData Then
        If (g.Broker.AccountExists(brokerMessage("Account")) = False) And (Len(brokerMessage("Account")) > 0) Then
            lAccountID = mTradeTracker.CreateAccountFromNumber(brokerMessage("Account"), Broker.Broker)
        Else
            lAccountID = g.Broker.AccountIDForNumber(brokerMessage("Account"))
        End If
    End If
    
    Set Order = New cPtOrder
    With Order
        .AccountID = lAccountID
        .BrokerID = brokerMessage("BrokerID")
        .ExchangeID = brokerMessage("ExchangeID")
        .GenesisOrderID = brokerMessage("GenesisID")
        .LimitPrice = Broker.GenesisPrice(brokerMessage("LimitPrice"), BrokerSym.PriceMult)
        .Message = brokerMessage("Message")
        .OrderType = GenesisOrderType(brokerMessage("Type"), brokerMessage("Inst"))
        .Quantity = brokerMessage("Quantity")
        .Status = GenesisOrderStatus(brokerMessage("Status"))
        .StatusDate = DateToDouble(brokerMessage("StatusDate"))
        .StopPrice = Broker.GenesisPrice(brokerMessage("StopPrice"), BrokerSym.PriceMult)
        .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.StatusDate, Broker.Broker, astrGenesisSymbols(0), False))
        
        Select Case GenesisTIF(brokerMessage("TIF"))
            Case eTT_TimeInForce_Day
                .Expiration = .SessionDate * -1&
            Case eTT_TimeInForce_GTC
                .Expiration = 0
            Case eTT_TimeInForce_GTD
                .Expiration = CLng(Val(brokerMessage("Expiration")))
        End Select
        
        .DebitCredit = mTradeTracker.DebitCreditFromString(brokerMessage("DebitCredit"))
        .UnderlyingSymbolOrSymbolID = brokerMessage("Underlying")
        
        lNumLegs = CLng(Val(brokerMessage("NumLegs")))
        For lIndex = 1 To lNumLegs
            Set OrderLeg = New cOrderLeg
            With OrderLeg
                .IsBuy = (Left(UCase(brokerMessage("Side" & Str(lIndex))), 3) = "BUY")
                .SymbolOrSymbolID = astrGenesisSymbols(lIndex - 1)
                .Multiplier = CLng(Val(brokerMessage("Multiplier" & Str(lIndex))))
            End With
            
            .OrderLegs.Add OrderLeg
        Next lIndex
    End With
    
    Set MultiLegOrderFromMessage = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.MultiLegOrderFromMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillFromMessage
'' Description: Fill a fill object from the given message
'' Inputs:      Message, Broker, Genesis symbol, Symbol information object, Account ID
'' Returns:     Fill
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillFromMessage(ByVal brokerMessage As cBrokerMessage, ByVal Broker As cBroker, ByVal strGenesisSymbol As String, ByVal BrokerSym As cBrokerSymbol, Optional ByVal lAccountID As Long = kNullData) As cPtFill
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Bars object
    Dim Fill As cPtFill                 ' Fill object to return
    
    Set Bars = New cGdBars
    SetBarProperties Bars, strGenesisSymbol
    
    If lAccountID = kNullData Then
        If (g.Broker.AccountExists(brokerMessage("Account")) = False) And (Len(brokerMessage("Account")) > 0) Then
            lAccountID = mTradeTracker.CreateAccountFromNumber(brokerMessage("Account"), Broker.Broker)
        Else
            lAccountID = g.Broker.AccountIDForNumber(brokerMessage("Account"))
        End If
    End If
    
    Set Fill = New cPtFill
    With Fill
        .AccountID = lAccountID
        .BrokerID = brokerMessage("FillID")
        .PreviousBrokerID = brokerMessage("PreviousFillID")
        .BrokerOrderID = brokerMessage("BrokerID")
        If Len(brokerMessage("Side")) > 0 Then
            .Buy = (Left(UCase(brokerMessage("Side")), 3) = "BUY")
        Else
            .Buy = (CLng(Val(brokerMessage("Quantity"))) > 0)
        End If
        .FillDate = DateToDouble(brokerMessage("FillDate"))
        .Price = Broker.GenesisPrice(brokerMessage("FillPrice"), BrokerSym.PriceMult)
        .Quantity = Abs(CLng(Val(brokerMessage("Quantity"))))
        
        ' 11/14/2012 DAJ: First of all, we need to set the symbol before the session date so that
        ' the correct symbol gets passed into the "ConvertTime" function.  Secondly, for efficiency
        ' purposes, we can pass in the exchange information so that we don't have to get it again...
        .SymbolOrSymbolID = strGenesisSymbol
        .SessionDate = Bars.SessionDateForTradeTime(.FillDateConverted(Bars.Prop(eBARS_ExchangeTimeZoneInf)))
    End With
    
    Set FillFromMessage = Fill

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.FillFromMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionFromMessage
'' Description: Fill a position object from the given message
'' Inputs:      Message, Broker, Genesis symbol, Account ID
'' Returns:     Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionFromMessage(ByVal brokerMessage As cBrokerMessage, ByVal Broker As cBroker, ByVal strGenesisSymbol As String, Optional ByVal lAccountID As Long = kNullData) As cPtPosition
On Error GoTo ErrSection:

    Dim Position As cPtPosition         ' Position object to return
    
    If lAccountID = kNullData Then
        If (g.Broker.AccountExists(brokerMessage("Account")) = False) And (Len(brokerMessage("Account")) > 0) Then
            lAccountID = mTradeTracker.CreateAccountFromNumber(brokerMessage("Account"), Broker.Broker)
        Else
            lAccountID = g.Broker.AccountIDForNumber(brokerMessage("Account"))
        End If
    End If
    
    Set Position = New cPtPosition
    With Position
        .AccountID = lAccountID
        .SymbolOrSymbolID = strGenesisSymbol
        If Len(brokerMessage("Current")) > 0 Then
            If Len(brokerMessage("CurrentSide")) > 0 Then
                If Left(UCase(brokerMessage("CurrentSide")), 3) = "BUY" Then
                    .CurrentPosition = CLng(Val(brokerMessage("Current")))
                Else
                    .CurrentPosition = CLng(Val(brokerMessage("Current"))) * -1&
                End If
            Else
                .CurrentPosition = CLng(Val(brokerMessage("Current")))
            End If
        End If
        If Len(brokerMessage("Carried")) > 0 Then
            If Len(brokerMessage("CarriedSide")) > 0 Then
                If Left(UCase(brokerMessage("CarriedSide")), 3) = "BUY" Then
                    .CarriedPosition = CLng(Val(brokerMessage("Carried")))
                Else
                    .CarriedPosition = CLng(Val(brokerMessage("Carried"))) * -1&
                End If
            Else
                .CarriedPosition = CLng(Val(brokerMessage("Carried")))
            End If
        End If
    End With
    
    Set PositionFromMessage = Position

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.PositionFromMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToMessage
'' Description: Build an order string to send to the broker
'' Inputs:      Order, New Order?, Broker Symbol, Symbol information object,
''              Broker object, Previous Genesis ID, Order Instruction
'' Returns:     Broker Order String
''
'' Fields:      GenesisID, BrokerID, Account, Symbol, Exchange, Order Type,
''              Buy/Sell, Quantity, Stop, Limit, TIF, Expiration
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderToMessage(ByVal Order As cPtOrder, ByVal bNewOrder As Boolean, ByVal strBrokerSymbol As String, ByVal BrokerSym As cBrokerSymbol, ByVal Broker As cBroker, Optional ByVal strPreviousGenesisID As String = "", Optional ByVal strInstruction As String = "") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker message
    
    strReturn = ""
    If (Len(strBrokerSymbol) > 0) And (Not BrokerSym Is Nothing) Then
        Set brokerMsg = New cBrokerMessage
        
        brokerMsg.Add "GenesisID", Order.GenesisOrderID
        If bNewOrder = False Then
            If g.Broker.IsOecBroker(Broker.Broker) Then
                brokerMsg.Add "BrokerID", Order.ExchangeID & ";" & Order.BrokerID
            Else
                brokerMsg.Add "BrokerID", Order.BrokerID
            End If
        End If
        brokerMsg.Add "Account", g.Broker.AccountNumberForID(Order.AccountID)
        brokerMsg.Symbol = strBrokerSymbol
        brokerMsg.Add "GenesisSymbol", Order.Symbol
        If (UCase(strInstruction) = "MIT") And (Order.OrderType = eTT_OrderType_MIT) Then
            brokerMsg.Add "Type", BrokerOrderType(eTT_OrderType_Limit)
        Else
            brokerMsg.Add "Type", BrokerOrderType(Order.OrderType)
        End If
        brokerMsg.Add "Inst", strInstruction
        brokerMsg.Add "Side", BrokerBuySell(Order.Buy)
        brokerMsg.Add "Quantity", Str(Order.Quantity)
        Select Case Order.OrderType
            Case eTT_OrderType_Stop
                brokerMsg.Add "StopPrice", Broker.BrokerPrice(Order.StopPrice, BrokerSym.PriceMult, BrokerSym.PriceFormat)
                brokerMsg.Add "GenesisStopPrice", Str(Order.StopPrice)
            Case eTT_OrderType_Limit
                brokerMsg.Add "LimitPrice", Broker.BrokerPrice(Order.LimitPrice, BrokerSym.PriceMult, BrokerSym.PriceFormat)
                brokerMsg.Add "GenesisLimitPrice", Str(Order.LimitPrice)
            Case eTT_OrderType_StopWithLimit
                brokerMsg.Add "StopPrice", Broker.BrokerPrice(Order.StopPrice, BrokerSym.PriceMult, BrokerSym.PriceFormat)
                brokerMsg.Add "GenesisStopPrice", Str(Order.StopPrice)
                brokerMsg.Add "LimitPrice", Broker.BrokerPrice(Order.LimitPrice, BrokerSym.PriceMult, BrokerSym.PriceFormat)
                brokerMsg.Add "GenesisLimitPrice", Str(Order.LimitPrice)
            Case eTT_OrderType_MIT
                brokerMsg.Add "LimitPrice", Broker.BrokerPrice(Order.LimitPrice, BrokerSym.PriceMult, BrokerSym.PriceFormat)
                brokerMsg.Add "GenesisLimitPrice", Str(Order.LimitPrice)
        End Select
        brokerMsg.Add "TIF", BrokerTIF(Order.TimeInForce)
        If Order.TimeInForce = eTT_TimeInForce_GTD Then
            brokerMsg.Add "Expiration", Format(Order.Expiration, "YYYYMMDD")
        End If
        If Order.IsAutomated Then
            brokerMsg.Add "Manual", "N"
        Else
            brokerMsg.Add "Manual", "Y"
        End If
        If Len(strPreviousGenesisID) > 0 Then
            brokerMsg.Add "PreviousGenesisID", strPreviousGenesisID
        End If
        If OrderIsEntry(Order, False, , True) Then
            brokerMsg.Add "Direction", "O"
        Else
            brokerMsg.Add "Direction", "C"
        End If
                
        strReturn = brokerMsg.ToString
    End If
    
    OrderToMessage = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.OrderToMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MultiLegOrderToMessage
'' Description: Build a multi-leg order string to send to the broker
'' Inputs:      Order, New Order?, Broker Symbol, Symbol information object,
''              Broker object, Previous Genesis ID, Order Instruction
'' Returns:     Broker Order String
''
'' Fields:      GenesisID, BrokerID, Account, Symbol, Exchange, Order Type,
''              Buy/Sell, Quantity, Stop, Limit, TIF, Expiration
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MultiLegOrderToMessage(ByVal Order As cPtOrder, ByVal bNewOrder As Boolean, ByVal astrBrokerSymbols As cGdArray, ByVal BrokerSym As cBrokerSymbol, ByVal Broker As cBroker, Optional ByVal strPreviousGenesisID As String = "", Optional ByVal strInstruction As String = "") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker message
    Dim lIndex As Long                  ' Index into a for loop
    
    strReturn = ""
    If (astrBrokerSymbols.Size > 0) And (Not BrokerSym Is Nothing) Then
        Set brokerMsg = New cBrokerMessage
        
        brokerMsg.Add "GenesisID", Order.GenesisOrderID
        If bNewOrder = False Then
            brokerMsg.Add "BrokerID", Order.BrokerID
        End If
        brokerMsg.Add "Account", g.Broker.AccountNumberForID(Order.AccountID)
        If (UCase(strInstruction) = "MIT") And (Order.OrderType = eTT_OrderType_MIT) Then
            brokerMsg.Add "Type", BrokerOrderType(eTT_OrderType_Limit)
        Else
            brokerMsg.Add "Type", BrokerOrderType(Order.OrderType)
        End If
        brokerMsg.Add "Inst", strInstruction
        Select Case Order.OrderType
            Case eTT_OrderType_Stop
                brokerMsg.Add "StopPrice", Broker.BrokerPrice(Order.StopPrice, BrokerSym.PriceMult, BrokerSym.PriceFormat)
            Case eTT_OrderType_Limit
                brokerMsg.Add "LimitPrice", Broker.BrokerPrice(Order.LimitPrice, BrokerSym.PriceMult, BrokerSym.PriceFormat)
            Case eTT_OrderType_StopWithLimit
                brokerMsg.Add "StopPrice", Broker.BrokerPrice(Order.StopPrice, BrokerSym.PriceMult, BrokerSym.PriceFormat)
                brokerMsg.Add "LimitPrice", Broker.BrokerPrice(Order.LimitPrice, BrokerSym.PriceMult, BrokerSym.PriceFormat)
            Case eTT_OrderType_MIT
                brokerMsg.Add "LimitPrice", Broker.BrokerPrice(Order.LimitPrice, BrokerSym.PriceMult, BrokerSym.PriceFormat)
        End Select
        brokerMsg.Add "TIF", BrokerTIF(Order.TimeInForce)
        If Order.TimeInForce = eTT_TimeInForce_GTD Then
            brokerMsg.Add "Expiration", Format(Order.Expiration, "YYYYMMDD")
        End If
        If Order.IsAutomated Then
            brokerMsg.Add "Manual", "N"
        Else
            brokerMsg.Add "Manual", "Y"
        End If
        If Len(strPreviousGenesisID) > 0 Then
            brokerMsg.Add "PreviousGenesisID", strPreviousGenesisID
        End If
        brokerMsg.Add "Quantity", Str(Order.Quantity)
        If Order.OrderType = eTT_OrderType_Market Then
            brokerMsg.Add "DebitCredit", mTradeTracker.DebitCreditString(eGDDebitCredit_Market)
        Else
            brokerMsg.Add "DebitCredit", mTradeTracker.DebitCreditString(Order.DebitCredit)
        End If
        brokerMsg.Add "Underlying", Order.UnderlyingSymbol
        
        brokerMsg.Add "NumLegs", Str(Order.NumberOfLegs)
        For lIndex = 1 To Order.OrderLegs.Count
            brokerMsg.LegSymbol(lIndex) = astrBrokerSymbols(lIndex - 1)
            brokerMsg.Add "Side" & Str(lIndex), BrokerBuySell(Order.OrderLegs(lIndex).IsBuy)
            brokerMsg.Add "Multiplier" & Str(lIndex), Str(Order.OrderLegs(lIndex).Multiplier)
            If mTradeTracker.OrderLegIsEntry(Order, lIndex, False) Then
                brokerMsg.Add "Direction" & Str(lIndex), "O"
            Else
                brokerMsg.Add "Direction" & Str(lIndex), "C"
            End If
        Next lIndex
                
        strReturn = brokerMsg.ToString
    End If
    
    MultiLegOrderToMessage = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.MultiLegOrderToMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderStatus
'' Description: Convert the broker order status to a Genesis order status
'' Inputs:      Broker Order Status
'' Returns:     Genesis Order Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderStatus(ByVal strBrokerOrderStatus As String) As eTT_OrderStatus
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderStatus      ' Return value for the function
    
    nReturn = -1&
    Select Case UCase(strBrokerOrderStatus)
        Case "NEW", "WORKING"
            nReturn = eTT_OrderStatus_Working
        
        Case "PARTIAL"
            nReturn = eTT_OrderStatus_Partial
        
        Case "FILLED", "COMPLETED"
            nReturn = eTT_OrderStatus_Filled
        
        Case "CANCELLED", "CANCELED", "CANCEL SUCCEEDED"
            nReturn = eTT_OrderStatus_Cancelled
        
        Case "REPLACED", "MODIFIED", "MODIFY FAILED"
            'nReturn = eTT_OrderStatus_Amended
            nReturn = eTT_OrderStatus_Working
        
        Case "CANCELPENDING", "PENDINGCANCEL", "CANCELLING"
            nReturn = eTT_OrderStatus_CancelPending
        
        Case "REJECTED"
            nReturn = eTT_OrderStatus_Rejected
        
        Case "SUSPENDED"
            nReturn = eTT_OrderStatus_Suspended
        
        Case "NEWPENDING", "PENDINGNEW"
            nReturn = eTT_OrderStatus_PreSubmitted
        
        Case "EXPIRED"
            nReturn = eTT_OrderStatus_Expired
        
        Case "REPLACEPENDING", "PENDINGREPLACE", "MODIFYING"
            nReturn = eTT_OrderStatus_AmendPending
        
        Case "PARKED"
            nReturn = eTT_OrderStatus_BrokerParked
            
        Case "SENT"
            nReturn = eTT_OrderStatus_Sent
            
        Case "ACCEPTED", "APPROVED"
            nReturn = eTT_OrderStatus_Approved
            
        Case "HELD"
            nReturn = eTT_OrderStatus_Held
    
    End Select
    
    GenesisOrderStatus = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.GenesisOrderStatus"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderType
'' Description: Convert the broker order type to a Genesis order type
'' Inputs:      Broker Order Type, Order Instruction
'' Returns:     Genesis Order Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderType(ByVal strBrokerOrderType As String, ByVal strInstruction As String) As eTT_OrderType
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderType        ' Return value for the function

    nReturn = -1&
    Select Case UCase(strBrokerOrderType)
        Case "MARKET"
            nReturn = eTT_OrderType_Market
        Case "STOP"
            nReturn = eTT_OrderType_Stop
        Case "LIMIT"
            If UCase(strInstruction) = "MIT" Then
                nReturn = eTT_OrderType_MIT
            Else
                nReturn = eTT_OrderType_Limit
            End If
        Case "STOPWITHLIMIT"
            nReturn = eTT_OrderType_StopWithLimit
        Case "MARKETONCLOSE"
            nReturn = eTT_OrderType_MarketOnClose
        Case "MARKETIFTOUCHED"
            nReturn = eTT_OrderType_MIT
        Case "MARKETONOPEN"
            nReturn = eTT_OrderType_MarketOnOpen
    End Select
    
    GenesisOrderType = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.GenesisOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisTIF
'' Description: Convert a broker Time-in-Force to a Genesis Time-in-Force
'' Inputs:      Broker Time-in-Force
'' Returns:     Genesis Time-in-Force
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisTIF(ByVal strBrokerTIF As String) As eTT_TimeInForce
On Error GoTo ErrSection:

    Dim nReturn As eTT_TimeInForce      ' Return value for the function

    Select Case UCase(strBrokerTIF)
        Case "DAY"
            nReturn = eTT_TimeInForce_Day
        Case "GTC"
            nReturn = eTT_TimeInForce_GTC
        Case "GTD"
            nReturn = eTT_TimeInForce_GTD
    End Select
    
    GenesisTIF = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.BrokerTIF"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerBuySell
'' Description: Convert a Genesis Buy/Sell to the broker buy/sell
'' Inputs:      Genesis Buy/Sell
'' Returns:     Broker Buy/Sell
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerBuySell(ByVal bBuySell As Boolean) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    If bBuySell Then
        strReturn = "Buy"
    Else
        strReturn = "Sell"
    End If
    
    BrokerBuySell = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.BrokerBuySell"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerOrderType
'' Description: Attempt to convert the Genesis order type to the broker order type
'' Inputs:      Genesis Order Type
'' Returns:     Broker Order Type (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerOrderType(ByVal nGenesisOrderType As eTT_OrderType) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    Select Case nGenesisOrderType
        Case eTT_OrderType_Market
            strReturn = "Market"
        Case eTT_OrderType_Stop
            strReturn = "Stop"
        Case eTT_OrderType_Limit
            strReturn = "Limit"
        Case eTT_OrderType_StopWithLimit
            strReturn = "StopWithLimit"
        Case eTT_OrderType_MarketOnClose
            strReturn = "MarketOnClose"
        Case eTT_OrderType_MIT
            strReturn = "MarketIfTouched"
        Case eTT_OrderType_MarketOnOpen
            strReturn = "MarketOnOpen"
    End Select
    
    BrokerOrderType = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.BrokerOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerTIF
'' Description: Convert a Genesis Time-in-Force to the broker Time-in-Force
'' Inputs:      Genesis Time-in-Force
'' Returns:     Broker Time-in-Force
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerTIF(ByVal nGenesisTIF As eTT_TimeInForce) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    Select Case nGenesisTIF
        Case eTT_TimeInForce_Day
            strReturn = "Day"
        Case eTT_TimeInForce_GTC
            strReturn = "GTC"
        Case eTT_TimeInForce_GTD
            strReturn = "GTD"
    End Select
    
    BrokerTIF = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.BrokerTIF"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateToDouble
'' Description: Convert a string version of a date/time to a Julian double
'' Inputs:      String Date/Time
'' Returns:     Julian Date/Time
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateToDouble(ByVal strDateTime As String) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value of the function
    Dim strDate As String               ' Date portion of the string
    Dim strTime As String               ' Time portion of the string
    Dim astrTime As cGdArray            ' Time broken out into an array
    Dim dTime As Double                 ' Time as fraction of the day
    
    dReturn = 0#
    strDate = Parse(strDateTime, "-", 1)
    strTime = Parse(strDateTime, "-", 2)
    
    Set astrTime = New cGdArray
    astrTime.SplitFields strTime, ":"
    
    dTime = (Val(astrTime(0)) / 24) + (Val(astrTime(1)) / 1440) + (Val(astrTime(2)) / 86400)
    dReturn = CDbl(JulFromLong(CLng(Val(strDate)))) + dTime
    
    DateToDouble = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerKeyValue.DateToDouble"
    
End Function
