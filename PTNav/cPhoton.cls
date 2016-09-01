VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPhoton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPhoton.cls
'' Description: Wrapper class for routines to access Photon Trader API
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 03/26/2009   DAJ         Use CommonFilesPath function instead of hard-coded
'' 09/01/2009   DAJ         Use new Parked order status
'' 03/11/2010   DAJ         Use global collections
'' 06/21/2011   DAJ         Last connection status change
'' 01/31/2012   DAJ         Handle unsolicted cancels in TradeSense order groups
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum ePhotonExchange
    ePhotonExchange_Eurex = 30
    ePhotonExchange_CME = 37
    ePhotonExchange_Liffe = 41
    ePhotonExchange_eCBOT = 96
End Enum

Private Enum ePhotonCommand
    ePhotonCommand_Order = 0
    ePhotonCommand_Cancel
    ePhotonCommand_Modify
End Enum

Private Type mPrivate
    strIniFile As String                ' Ini File to use for the Photon info
    nBroker As eTT_AccountType          ' Broker type for the Photon broker
    
    lUserID As Long                     ' ID for the user logging in
    strPassword As String               ' Password for the user to logon with
    strIniPassword As String            ' Password for the user to logon with
    strServerIP As String               ' IP for the server to logon to
    iServerPort As Integer              ' Port for the server to logon to
    strEuropeIP As String               ' IP address for the Europe server
    iEuropePort As Integer              ' Port for the Europe server
    
    nStatus As eGDConnectionStatus      ' Current overall connection status
    nDataStatus As eGDConnectionStatus  ' Current data connection status
    nCmeStatus As eGDConnectionStatus   ' Current Cme connection status
    nCbotStatus As eGDConnectionStatus  ' Current Cbot connection status
    nEurexStatus As eGDConnectionStatus ' Current Eurex connection status
    nLiffeStatus As eGDConnectionStatus ' Current Liffe connection status
    
    bConnectToMarket As Boolean         ' Does the user want to connect to the Market?
    lMarketNode As Long                 ' Port for the market data subnode
    bConnectToCME As Boolean            ' Does the user want to connect to the CME?
    lCmeNode As Long                    ' Port for the CME data subnode
    bConnectToCBOT As Boolean           ' Does the user want to connect to the CBOT?
    lCbotNode As Long                   ' Port for the CBOT data subnode
    bConnectToEurex As Boolean          ' Does the user want to connect to Eurex?
    lEurexNode As Long                  ' Port for the Eurex data subnode
    bConnectToLiffe As Boolean          ' Does the user want to connect to Liffe?
    lLiffeNode As Long                  ' Port for the Liffe data subnode
        
    bConnectedToMarket As Boolean       ' Are we connected to the Market server?
    bConnectedToCME As Boolean          ' Are we connected to the CME server?
    bConnectedToCBOT As Boolean         ' Are we connected to the CBT server?
    bConnectedToEurex As Boolean        ' Are we connected to the Eurex server?
    bConnectedToLiffe As Boolean        ' Are we connected to the Liffe server?
    bAskedPassword As Boolean           ' Have we asked the user for the password?
    bManualDisconnect As Boolean        ' Was this a manual disconnect attempt?
    bVerifyPositions As Boolean         ' Do we want to verify the positions with the user?
    dRefreshingStart As Double          ' Start of the refreshing time
    bRefreshingOrders As Boolean        ' Are we refreshing orders?
    
    dHeartBeat As Double                ' Heartbeat
    dLastStatusMsg As Double            ' Tick count of the last status message
    bAppLoaded As Boolean               ' Is the GenPhoton application loaded?
        
    ModifiedOrders As cGdTree           ' Collection of modified orders
    astrErrorCodes As cGdArray          ' Array of Photon error codes and descriptions
    astrBaseSyms As cGdArray            ' Array of base symbol translation information
    astrSubContract As cGdArray         ' Array of subscribed contracts
    
    BrokerInfo As cBrokerInfo           ' Object of broker information
End Type
Private m As mPrivate

Public Property Get UserID() As Long
    UserID = m.lUserID
End Property
Public Property Get Password() As String
    Password = m.strPassword
End Property
Public Property Get ServerIP() As String
    ServerIP = m.strServerIP
End Property
Public Property Get ServerPort() As Integer
    ServerPort = m.iServerPort
End Property
Public Property Get EuropeIP() As String
    EuropeIP = m.strEuropeIP
End Property
Public Property Get EuropePort() As Integer
    EuropePort = m.iEuropePort
End Property
Public Property Get MarketNode() As Long
    MarketNode = m.lMarketNode
End Property
Public Property Get CmeNode() As Long
    CmeNode = m.lCmeNode
End Property
Public Property Get CbotNode() As Long
    CbotNode = m.lCbotNode
End Property
Public Property Get EurexNode() As Long
    EurexNode = m.lEurexNode
End Property
Public Property Get LiffeNode() As Long
    LiffeNode = m.lLiffeNode
End Property

Public Property Get ConnectToMarket() As Boolean
    ConnectToMarket = m.bConnectToMarket
End Property
Public Property Get ConnectToCME() As Boolean
    ConnectToCME = m.bConnectToCME
End Property
Public Property Get ConnectToCBOT() As Boolean
    ConnectToCBOT = m.bConnectToCBOT
End Property
Public Property Get ConnectToEurex() As Boolean
    ConnectToEurex = m.bConnectToEurex
End Property
Public Property Get ConnectToLiffe() As Boolean
    ConnectToLiffe = m.bConnectToLiffe
End Property

Public Property Get ModifiedOrders() As cGdTree
    Set ModifiedOrders = m.ModifiedOrders
End Property
Public Property Let ModifiedOrders(ByVal ModOrders As cGdTree)
    Set m.ModifiedOrders = ModOrders
End Property

Public Property Get HeartBeat() As Double
    HeartBeat = m.dHeartBeat
End Property

Public Property Get LastStatusMessage() As Double
    LastStatusMessage = m.dLastStatusMsg
End Property

Public Property Get AppLoaded() As Boolean
    AppLoaded = m.bAppLoaded
End Property

Public Property Get ConnectionStatus() As eGDConnectionStatus
    ConnectionStatus = m.nStatus
End Property
Public Property Let ConnectionStatus(ByVal nNewStatus As eGDConnectionStatus)
    If nNewStatus <> m.nStatus Then
        m.nStatus = nNewStatus
        If nNewStatus = eGDConnectionStatus_Disconnected Then
            m.bConnectedToMarket = False
            m.bConnectedToCME = False
            m.bConnectedToCBOT = False
            m.bConnectedToEurex = False
            m.bConnectedToLiffe = False
        End If
        
        If Not g.Broker Is Nothing Then
            g.Broker.LastStatusChange = Now
        End If
        If nNewStatus = eGDConnectionStatus_Connected Then
            g.Broker.PositionsToVerify = True
        End If
    End If
End Property

Public Property Get DataStatus() As eGDConnectionStatus
    DataStatus = m.nDataStatus
End Property

Public Property Get CmeStatus() As eGDConnectionStatus
    CmeStatus = m.nCmeStatus
End Property

Public Property Get CbotStatus() As eGDConnectionStatus
    CbotStatus = m.nCbotStatus
End Property

Public Property Get EurexStatus() As eGDConnectionStatus
    EurexStatus = m.nEurexStatus
End Property

Public Property Get LiffeStatus() As eGDConnectionStatus
    LiffeStatus = m.nLiffeStatus
End Property

Public Property Get BrokerInfo() As cBrokerInfo
    Set BrokerInfo = m.BrokerInfo
End Property

Public Property Get Refreshing() As Boolean
    Refreshing = m.bRefreshingOrders
End Property

Public Property Get PositionVerify() As Boolean
    PositionVerify = m.bVerifyPositions
End Property
Public Property Let PositionVerify(ByVal bVerifyPositions As Boolean)
    m.bVerifyPositions = bVerifyPositions
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowConfig
'' Description: Determines whether we should show the config form on connect
'' Inputs:      None
'' Returns:     True if show configuration form, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowConfig() As Boolean
On Error GoTo ErrSection:

    If (m.lUserID = 0) Or (Len(m.strIniPassword) = 0) Or (Len(m.strServerIP) = 0) Or (m.iServerPort = 0) Then
        ShowConfig = True
    ElseIf (m.bConnectToCME = True) And (m.lCmeNode = 0) Then
        ShowConfig = True
    ElseIf (m.bConnectToCBOT = True) And (m.lCbotNode = 0) Then
        ShowConfig = True
    ElseIf (m.bConnectToEurex = True) And (m.lEurexNode = 0) Then
        ShowConfig = True
    ElseIf (m.bConnectToLiffe = True) And (m.lLiffeNode = 0) Then
        ShowConfig = True
    ElseIf (m.bConnectToMarket = True) And (m.lMarketNode = 0) Then
        ShowConfig = True
    Else
        ShowConfig = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.ShowConfig", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Connect
'' Description: Connect to the appropriate Photon servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Connect()
On Error GoTo ErrSection:

    Dim bConnect As Boolean             ' Should we attempt to connect?
    Dim strProcess As String            ' Name of the process to start
    Dim strPassword As String           ' Password given by the user

    If (Connected = False) And (g.bUnloading = False) Then
        bConnect = True
        GetConnectionProperties
        
        If ShowConfig = True Then
            If frmBrokerCfg.ShowMe(m.nBroker, False) Then
                GetConnectionProperties
            Else
                bConnect = False
            End If
        ElseIf Len(m.strPassword) = 0 Then
            strPassword = InfBox("Enter the password for Photon user " & Str(m.lUserID) & ":", , "+OK|-Cancel", "Photon Login", , , , , , "password", "", eGDAlign_Left)
            If Len(strPassword) = 0 Then
                bConnect = False
            Else
                m.strPassword = strPassword
            End If
        End If
        
        If (bConnect = True) And (g.bUnloading = False) Then
            If LiveTradingAllowed(m.nBroker) Then
                m.bManualDisconnect = False
                If (frmOnlineBroker.gdBroker.FindControlID("Photon") = 0) Or (KillProcess("GenPhoton", True) = 0) Then
                    DumpDebug "Trade Navigator version: " & FormatVersion & "." & Str(App.Revision) & " " & DateFormat(FileDate(App.Path & "\" & App.EXEName & ".EXE"), MM_DD_YYYY, HH_MM, AMPM_UPPER)
                    strProcess = AddSlash(App.Path) & "..\Brokers\Photon\GenPhoton.EXE"
                    DumpDebug "Starting application: " & strProcess
                    RunProcess strProcess, , , vbHide ' , vbMinimizedNoFocus
                Else
                    DumpDebug "Connecting as " & Str(m.lUserID)
                    SendPhotonMessage eGDPhotonMessageType_Connect, ConnectString
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.Connect"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disconnect
'' Description: Disconnect form the appropriate Photon servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Disconnect(Optional ByVal bForceReconnect As Boolean = False, Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    If Connected Then
        m.bManualDisconnect = Not bForceReconnect
        If Len(strReason) = 0 Then DumpDebug "Disconnecting" Else DumpDebug "Disconnecting: " & strReason
        SendPhotonMessage eGDPhotonMessageType_Disconnect, ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.Disconnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshConnections
'' Description: Refresh the connections in case ports have changed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshConnections(Optional ByVal bReconnect As Boolean = True)
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Time out counter

    If Connected Then
        Disconnect False, "RefreshConnections"
    End If
    
    lTimeOut = 0&
    Do While (Connected = True) And (lTimeOut < 20)
        Sleep 1
        lTimeOut = lTimeOut + 1&
    Loop
    
    If bReconnect And (Not Connected) Then
        If LiveTradingAllowed(m.nBroker) Then Connect
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.RefreshConnections", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResyncSubscriptionList
'' Description: Synchronize the subscription list with the GenPhoton program
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ResyncSubscriptionList()
On Error GoTo ErrSection:
    
    Dim astrSymbols As New cGdArray     ' List of symbols that should be subcribed to
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Translated symbol
    Dim strSymbols As String            ' String of symbols
    Dim strPhotonExchange As String     ' Photon exchange
    
    If (m.bConnectedToMarket = True) Then
        Set astrSymbols = g.RealTime.BrokerSubscriptionList(m.nBroker)
        If Not astrSymbols Is Nothing Then
            astrSymbols.Sort
            
            ' Walk through real-time list and subscribe to any symbols that aren't already
            ' subscribed to...
            strSymbols = ""
            For lIndex = 0 To astrSymbols.Size - 1
                strSymbol = PhotonSymbol(astrSymbols(lIndex), strPhotonExchange)
                If m.astrSubContract.BinarySearch(strSymbol & "|" & strPhotonExchange) = False Then
                    strSymbols = strSymbols & strSymbol & "|" & strPhotonExchange & ","
                End If
            Next lIndex
            If Len(strSymbols) > 0 Then
                If Right(strSymbols, 1) = "," Then strSymbols = Left(strSymbols, Len(strSymbols) - 1)
                SendPhotonMessage eGDPhotonMessageType_Subscribe, strSymbols
            End If
            
            ' Walk through subscribed list and unsubscribe to any symbols that are no longer
            ' in the real-time list...
            strSymbols = ""
            For lIndex = 0 To m.astrSubContract.Size - 1
                strSymbol = GenesisSymbol(Parse(m.astrSubContract(lIndex), "|", 1), Parse(m.astrSubContract(lIndex), "|", 2))
                If astrSymbols.BinarySearch(strSymbol) = False Then
                    strSymbols = strSymbols & m.astrSubContract(lIndex) & ","
                End If
            Next lIndex
            If Len(strSymbols) > 0 Then
                If Right(strSymbols, 1) = "," Then strSymbols = Left(strSymbols, Len(strSymbols) - 1)
                SendPhotonMessage eGDPhotonMessageType_Unsubscribe, strSymbols
            End If
        
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.ResyncSubscriptionList"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendOrder
'' Description: Attempt to add or amend the order as appropriate
'' Inputs:      Order to Send
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendOrder(Order As cPtOrder)
On Error GoTo ErrSection:

    If NotSent(Order.Status) Or (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        If AddOrder(Order) = False Then
            Order.Status = eTT_OrderStatus_Error
        Else
            Order.Status = eTT_OrderStatus_Sent
        End If
    Else
        If AmendOrder(Order) = False Then
            Order.Status = eTT_OrderStatus_Error
        Else
            Order.Status = eTT_OrderStatus_AmendPending
        End If
    End If
                            
    Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
    Order.Save
    
    If Order.Status <> eTT_OrderStatus_Error Then
        m.BrokerInfo.AddOrder Order, False
    End If
    
    OrderCallback Order
    g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.SendOrder"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOrder
'' Description: Add an order to the Photon trading interface
'' Inputs:      Order to add
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AddOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information to pass to Photon
    Dim strSymbol As String             ' Symbol back from a convert
    Dim strExchange As String           ' Exchange from convert symbol
    Dim dMult As Double                 ' Multiplier from convert symbol
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    strSymbol = PhotonSymbol(Order.Symbol, strExchange, dMult)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    
    If (Len(strSymbol) = 0) Or (Len(strExchange) = 0) Then
        DumpDebug "AddOrder: " & Order.Symbol & " cannot be traded through the Photon servers"
        InfBox Order.Symbol & " cannot be traded through the Photon servers", "!", , "Photon Order Error"
        Order.Message = Order.Symbol & " cannot be traded through the Photon servers"
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        DumpDebug "AddOrder: Not currently connected to Photon account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to Photon account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , "Photon Order Error"
        Order.Message = "Not currently connected to Photon account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "AddOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the Photon servers", "!", , "Photon Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "AddOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the Photon servers", "!", , "Photon Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
    Else
        astrOrder(0) = strExchange
        astrOrder(1) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(2) = Order.GenesisOrderID
        astrOrder(3) = Str(ePhotonCommand_Order)
        astrOrder(4) = strSymbol
        If Order.OrderDate = 0 Then
            astrOrder(5) = Format(CurrentTime, "YYYYMMDDHHMMSS")
        Else
            astrOrder(5) = Format(Order.OrderDate, "YYYYMMDDHHMMSS")
        End If
        astrOrder(6) = OrderType(Order.OrderType)
        astrOrder(7) = Str(Order.LimitPrice * dMult)
        astrOrder(8) = Str(Order.StopPrice * dMult)
        astrOrder(9) = Str(Order.Quantity)
        astrOrder(10) = ""
        If Order.Buy Then astrOrder(11) = "0" Else astrOrder(11) = "1"
        astrOrder(12) = ""
        
        'GenesisOrderToBrokerInfo Order, eTT_OrderStatus_Sent
        DumpDebug "AddOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
        SendPhotonMessage eGDPhotonMessageType_SendOrder, astrOrder.JoinFields(vbTab)
        bReturn = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
    AddOrder = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.AddOrder", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AmendOrder
'' Description: Amend an order to the Photon trading interface
'' Inputs:      Order to amend
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AmendOrder(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:
    
    Dim astrOrder As New cGdArray       ' Array of order information to pass to Photon
    Dim strSymbol As String             ' Symbol back from a convert
    Dim strExchange As String           ' Exchange from convert symbol
    Dim dMult As Double                 ' Multiplier from convert symbol
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    strSymbol = PhotonSymbol(Order.Symbol, strExchange, dMult)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    
    If (Len(strSymbol) = 0) Or (Len(strExchange) = 0) Then
        DumpDebug "AmendOrder: " & Order.Symbol & " cannot be traded through the Photon servers"
        InfBox Order.Symbol & " cannot be traded through the Photon servers", "!", , "Photon Order Error"
        Order.Message = Order.Symbol & " cannot be traded through the Photon servers"
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        DumpDebug "AmendOrder: Not currently connected to Photon account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to Photon account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , "Photon Order Error"
        Order.Message = "Not currently connected to Photon account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "AmendOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the Photon servers", "!", , "Photon Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "AmendOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the Photon servers", "!", , "Photon Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
    Else
        astrOrder(0) = strExchange
        astrOrder(1) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(2) = Order.GenesisOrderID
        astrOrder(3) = Str(ePhotonCommand_Modify)
        astrOrder(4) = strSymbol
        If Order.OrderDate = 0 Then
            astrOrder(5) = Format(CurrentTime, "YYYYMMDDHHMMSS")
        Else
            astrOrder(5) = Format(Order.OrderDate, "YYYYMMDDHHMMSS")
        End If
        astrOrder(6) = OrderType(Order.OrderType)
        astrOrder(7) = Str(Order.LimitPrice * dMult)
        astrOrder(8) = Str(Order.StopPrice * dMult)
        astrOrder(9) = Str(Order.RemainingQuantity)
        astrOrder(10) = Order.ExchangeID
        If Order.Buy Then astrOrder(11) = "0" Else astrOrder(11) = "1"
        astrOrder(12) = Str(Order.PrevQty)
        
        DumpDebug "ModifyOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & ")"
        SendPhotonMessage eGDPhotonMessageType_SendOrder, astrOrder.JoinFields(vbTab)
        bReturn = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
    AmendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.AmendOrder", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Cancel an order on the Photon trading interface
'' Inputs:      Order to cancel
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information to pass to Photon
    Dim strSymbol As String             ' Symbol back from a convert
    Dim strExchange As String           ' Exchange from convert symbol
    Dim dMult As Double                 ' Multiplier from convert symbol
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    strSymbol = PhotonSymbol(Order.Symbol, strExchange, dMult)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    
    If (Len(strSymbol) = 0) Or (Len(strExchange) = 0) Then
        DumpDebug "CancelOrder: " & Order.Symbol & " cannot be traded through the Photon servers"
        InfBox Order.Symbol & " cannot be traded through the Photon servers", "!", , "Photon Order Error"
        Order.Message = Order.Symbol & " cannot be traded through the Photon servers"
    ElseIf g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True Then
        DumpDebug "CancelOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
        
        Order.OrderType = nPrevOrderType
        Order.Status = eTT_OrderStatus_Cancelled
        Order.StatusDate = ConvertToBrokerDate(CurrentTime, g.Broker.AccountTypeForID(Order.AccountID), Order.Symbol, True)
        Order.Save
        
        g.TradingItems.OrderCallback Order.AutoTradeItemID, Order
        g.Broker.AddOrder Order
        OrderCallback Order
        
        g.Broker.SyntheticOrders.Remove Str(Order.OrderID)
        
        DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        DumpDebug "CancelOrder: Not currently connected to Photon account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to Photon account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , "Photon Order Error"
        Order.Message = "Not currently connected to Photon account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "CancelOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the Photon servers", "!", , "Photon Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "CancelOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the Photon servers", "!", , "Photon Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
    Else
        If Not g.TsoGroups Is Nothing Then
            g.TsoGroups.CancelRequested Order
        End If
        
        astrOrder(0) = strExchange
        astrOrder(1) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(2) = Order.GenesisOrderID
        astrOrder(3) = Str(ePhotonCommand_Cancel)
        astrOrder(4) = strSymbol
        If Order.OrderDate = 0 Then
            astrOrder(5) = Format(CurrentTime, "YYYYMMDDHHMMSS")
        Else
            astrOrder(5) = Format(Order.OrderDate, "YYYYMMDDHHMMSS")
        End If
        astrOrder(6) = OrderType(Order.OrderType)
        astrOrder(7) = Str(Order.LimitPrice * dMult)
        astrOrder(8) = Str(Order.StopPrice * dMult)
        astrOrder(9) = Str(Order.Quantity)
        astrOrder(10) = Order.ExchangeID
        If Order.Buy Then astrOrder(11) = "0" Else astrOrder(11) = "1"
        astrOrder(12) = Str(Order.PrevQty)
        
        DumpDebug "CancelOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & ")"
        SendPhotonMessage eGDPhotonMessageType_SendOrder, astrOrder.JoinFields(vbTab)
        
        Order.Status = eTT_OrderStatus_CancelPending
        Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
        Order.Save
        m.BrokerInfo.AddOrder Order, False
        bReturn = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
    CancelOrder = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.CancelOrder", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Park an order on the Photon trading interface
'' Inputs:      Order to Park
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ParkOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information to pass to Photon
    Dim strSymbol As String             ' Symbol back from a convert
    Dim strExchange As String           ' Exchange from convert symbol
    Dim dMult As Double                 ' Multiplier from convert symbol
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    strSymbol = PhotonSymbol(Order.Symbol, strExchange, dMult)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    
    If (Len(strSymbol) = 0) Or (Len(strExchange) = 0) Then
        DumpDebug "ParkOrder: " & Order.Symbol & " cannot be traded through the Photon servers"
        InfBox Order.Symbol & " cannot be traded through the Photon servers", "!", , "Photon Order Error"
        Order.Message = Order.Symbol & " cannot be traded through the Photon servers"
    ElseIf g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True Then
        DumpDebug "ParkOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
        
        Order.OrderType = nPrevOrderType
        Order.Status = eTT_OrderStatus_Parked
        Order.StatusDate = ConvertToBrokerDate(CurrentTime, g.Broker.AccountTypeForID(Order.AccountID), Order.Symbol, True)
        Order.Save
        
        g.TradingItems.OrderCallback Order.AutoTradeItemID, Order
        g.Broker.AddOrder Order
        OrderCallback Order
        
        g.Broker.SyntheticOrders.Remove Str(Order.OrderID)
        
        DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        DumpDebug "ParkOrder: Not currently connected to Photon account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to Photon account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , "Photon Order Error"
        Order.Message = "Not currently connected to Photon account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "ParkOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the Photon servers", "!", , "Photon Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "ParkOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the Photon servers", "!", , "Photon Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the Photon servers"
    Else
        astrOrder(0) = strExchange
        astrOrder(1) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(2) = Order.GenesisOrderID
        astrOrder(3) = Str(ePhotonCommand_Cancel)
        astrOrder(4) = strSymbol
        If Order.OrderDate = 0 Then
            astrOrder(5) = Format(CurrentTime, "YYYYMMDDHHMMSS")
        Else
            astrOrder(5) = Format(Order.OrderDate, "YYYYMMDDHHMMSS")
        End If
        astrOrder(6) = OrderType(Order.OrderType)
        astrOrder(7) = Str(Order.LimitPrice * dMult)
        astrOrder(8) = Str(Order.StopPrice * dMult)
        astrOrder(9) = Str(Order.Quantity)
        astrOrder(10) = Order.ExchangeID
        If Order.Buy Then astrOrder(11) = "0" Else astrOrder(11) = "1"
        astrOrder(12) = Str(Order.PrevQty)
        
        DumpDebug "ParkOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & ")"
        SendPhotonMessage eGDPhotonMessageType_SendOrder, astrOrder.JoinFields(vbTab)
        
        Order.Status = eTT_OrderStatus_ParkPending
        Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
        Order.Save
        m.BrokerInfo.AddOrder Order, False
        bReturn = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
        
    ParkOrder = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.ParkOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrderFromBrokerID
'' Description: Attempt to cancel the order
'' Inputs:      Broker ID, Genesis ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelOrderFromBrokerID(ByVal strBrokerID As String, Optional ByVal strGenesisID As String = "")
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order retrieved from broker info
    
    Set Order = m.BrokerInfo.Order(strBrokerID, strGenesisID)
    If Not Order Is Nothing Then
        CancelOrder Order
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.CancelOrderFromBrokerID"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelWorkingOrders
'' Description: Cancel all working orders for the given account and symbol
'' Inputs:      Account, Symbol, Auto Trade Item ID, Skip Order ID
'' Returns:     Number of orders cancelled
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelWorkingOrders(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, Optional ByVal lSkipOrderID As Long = -1&) As Long
On Error GoTo ErrSection:
    
    Dim Orders As cPtOrders             ' Collection of working orders
    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value for the function
    
    DumpDebug "Cancelling Working Orders for: " & strAccount & ", " & strSymbol & ", " & Str(lAtID)
    
    lReturn = 0&
    Set Orders = m.BrokerInfo.OrdersForSymbol(strAccount, strSymbol, lAtID, True)
    If Not Orders Is Nothing Then
        For lIndex = Orders.Count To 1 Step -1
            If (OrderIsPending(Orders(lIndex)) = False) And (Orders(lIndex).OrderID <> lSkipOrderID) Then
                CancelOrder Orders(lIndex)
                lReturn = lReturn + 1&
            End If
        Next lIndex
    End If
    
    CancelWorkingOrders = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.CancelWorkingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenPosition
'' Description: Flattent the position for the given account and symbol
'' Inputs:      Account, Symbol, Auto Trade Item ID, Confirm?, Genesis ID
'' Returns:     Did user Confirm?
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FlattenPosition(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, Optional ByVal bConfirm As Boolean = True, Optional strGenesisOrderID As String) As Boolean
On Error GoTo ErrSection:

    Dim lPosition As Long               ' Current position
    Dim strMessage As String            ' Message to display to the user
    Dim strReturn As String             ' Return from user question
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    DumpDebug "Flattening Position for: " & strAccount & ", " & strSymbol & ", " & Str(lAtID)
    lPosition = m.BrokerInfo.CurrentPosition(strAccount, strSymbol, lAtID)
    If lPosition <> 0& Then
        If bConfirm Then
            If lPosition < 0 Then
                strMessage = "BUY " & Abs(lPosition) & " " & strSymbol & " at the MARKET"
            Else
                strMessage = "SELL " & Abs(lPosition) & " " & strSymbol & " at the MARKET"
            End If
            strReturn = InfBox("Are you sure you want to|" & strMessage & "|in account '" & strAccount & "'|to flatten your position?", "?", "+Yes|-No", "Flatten for " & strSymbol & " in " & strAccount)
        Else
            strReturn = "Y"
        End If
        
        If strReturn = "Y" Then
            bReturn = True
            DumpDebug "User Confirmed: " & strMessage & " in account " & strAccount
            ExitPositionForSymbol strAccount, strSymbol, lAtID, lPosition, strGenesisOrderID
        End If
    End If
    
    FlattenPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.FlattenPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Ask Photon for the current orders in their database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetOrders()
On Error GoTo ErrSection:

    If Not Refreshing Then
        DumpDebug "Getting Orders"
        m.bRefreshingOrders = True
        SendPhotonMessage eGDPhotonMessageType_GetOrders, ""
    Else
        DumpDebug "Refresh requested but already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.GetOrders", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectionStatusForAccount
'' Description: Determine the connection status for the given account
'' Inputs:      Account Number
'' Returns:     Connection Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConnectionStatusForAccount(ByVal strAccountNumber As String) As eGDConnectionStatus
On Error GoTo ErrSection:

    ConnectionStatusForAccount = ConnectionStatus

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.ConnectionStatusForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextGenesisID
'' Description: Figure out the next unique identifier to use for an order
'' Inputs:      Account Number
'' Returns:     Next ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextGenesisID(ByVal strAccountNumber As String) As String
On Error GoTo ErrSection:

    Dim lDate As Long                   ' Date from the ini file
    Dim lCounter As Long                ' Counter from the ini file
    Dim lCustomerID As Long             ' Customer ID from the registry
    
    lDate = GetIniFileProperty("Date", 0&, "ID", m.strIniFile)
    lCounter = GetIniFileProperty("Counter", 0&, "ID", m.strIniFile)
    lCustomerID = (RI_GetLastDataServiceID \ 1000) Mod 1000000
    
    If lDate <> Date Then
        lDate = Date
        lCounter = 0&
    Else
        lCounter = lCounter + 1
    End If

    SetIniFileProperty "Date", lDate, "ID", m.strIniFile
    SetIniFileProperty "Counter", lCounter, "ID", m.strIniFile
    
    NextGenesisID = Format(lCustomerID, "000000") & Format(lDate, "00000") & Format(lCounter, "00000")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.NextGenesisID", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PhotonExchangeForSymbol
'' Description: Return what Photon exchange the given Genesis symbol is on
'' Inputs:      Genesis Symbol
'' Returns:     Photon Exchange
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PhotonExchangeForSymbol(ByVal strGenesisSymbol As String) As ePhotonExchange
On Error GoTo ErrSection:

    Dim strFeedSymbol As String         ' Feed symbol coming back from conversion
    Dim strExchange As String           ' Exchange coming back from conversion
    Dim dMult As Double                 ' Multiplier coming back from conversion

    strFeedSymbol = PhotonSymbol(strGenesisSymbol, strExchange, dMult)
    If Len(strFeedSymbol) > 0 Then
        PhotonExchangeForSymbol = Val(strExchange)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.PhotonExchangeForSymbol", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTradeableSymbol
'' Description: Is the given symbol in the conversion table?
'' Inputs:      Symbol
'' Returns:     True if tradeable, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsTradeableSymbol(ByVal strGenesisSymbol As String) As Boolean
On Error GoTo ErrSection:

    IsTradeableSymbol = (Len(PhotonSymbol(strGenesisSymbol)) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.IsTradeableSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTypeAllowed
'' Description: Is the given order type allowed for the given symbol?
'' Inputs:      Order Type, Symbol
'' Returns:     True if allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderTypeAllowed(ByVal nOrderType As eTT_OrderType, ByVal strGenesisSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim strBaseSymbol As String         ' Genesis base symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim strMask As String               ' Order type mask from the file
    
    bReturn = False
    
    If nOrderType <> eTT_OrderType_Adjustment Then
        strBaseSymbol = Parse(strGenesisSymbol, "-", 1)
        For lIndex = 0 To m.astrBaseSyms.Size - 1
            If Parse(m.astrBaseSyms(lIndex), vbTab, 3) = strBaseSymbol Then
                strMask = Parse(m.astrBaseSyms(lIndex), vbTab, 5)
                If Len(strMask) = 0 Then strMask = "11110000"
                If Len(strMask) = 8 Then strMask = strMask & "0"
                If (nOrderType >= 0) And (nOrderType < Len(strMask)) Then
                    bReturn = (Mid(strMask, nOrderType + 1, 1) = "1")
                End If
            End If
        Next lIndex
    End If
    
    OrderTypeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.OrderTypeAllowed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimeInForceAllowed
'' Description: Is the given time in force allowed for the given symbol?
'' Inputs:      Time In Force, Symbol
'' Returns:     True if allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TimeInForceAllowed(ByVal nTimeInForce As eTT_TimeInForce, ByVal strGenesisSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim strBaseSymbol As String         ' Genesis base symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim strMask As String               ' Order type mask from the file
    
    bReturn = False
    
    strBaseSymbol = Parse(strGenesisSymbol, "-", 1)
    For lIndex = 0 To m.astrBaseSyms.Size - 1
        If Parse(m.astrBaseSyms(lIndex), vbTab, 3) = strBaseSymbol Then
            strMask = Parse(m.astrBaseSyms(lIndex), vbTab, 6)
            If Len(strMask) = 0 Then strMask = "111"
            If (nTimeInForce >= 0) And (nTimeInForce < Len(strMask)) Then
                bReturn = (Mid(strMask, nTimeInForce + 1, 1) = "1")
            End If
        End If
    Next lIndex
    
    TimeInForceAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.TimeInForceAllowed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Send the given string to the test form and the debug log
'' Inputs:      String to Send
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DumpDebug(ByVal strDebug As String)
On Error GoTo ErrSection:

    GenesisLog strDebug

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.DumpDebug", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderStatus
'' Description: Convert a Photon Order status to a Genesis Order Status
'' Inputs:      Photon Order Status
'' Returns:     Genesis Order Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderStatus(ByVal strOrderStatus As String) As eTT_OrderStatus
On Error GoTo ErrSection:

    Select Case UCase(strOrderStatus)
        Case "E"                        ' Awaiting a response
            GenesisOrderStatus = eTT_OrderStatus_Sent
        Case "O"                        ' Acknowledged
            GenesisOrderStatus = eTT_OrderStatus_Working
        Case "OX"                       ' Acknowledged and partially executed
            GenesisOrderStatus = eTT_OrderStatus_Partial
        Case "OG"                       ' Frozen
            GenesisOrderStatus = eTT_OrderStatus_Frozen
        Case "TX"                       ' Totally executed
            GenesisOrderStatus = eTT_OrderStatus_Filled
        Case "A"                        ' Cancelled
            GenesisOrderStatus = eTT_OrderStatus_Cancelled
        Case "AX"                       ' Partially executed and cancelled
            GenesisOrderStatus = eTT_OrderStatus_BalCancelled
        Case "AG"                       ' Frozen and cancelled
            GenesisOrderStatus = eTT_OrderStatus_Cancelled
        Case "R"                        ' Rejected
            GenesisOrderStatus = eTT_OrderStatus_Rejected
        Case "RX"                       ' Partially executed and rejected
            GenesisOrderStatus = eTT_OrderStatus_Rejected
        Case " "                        ' Awaiting cancel acknowledgement
            GenesisOrderStatus = eTT_OrderStatus_CancelPending
        Case " X"                       ' Partially executed and awaiting cancel
            GenesisOrderStatus = eTT_OrderStatus_CancelPending
        Case "N"                        ' Awaiting modification acknowledgement
            GenesisOrderStatus = eTT_OrderStatus_AmendPending
        Case "NX"                       ' Partially executed and awaiting modify
            GenesisOrderStatus = eTT_OrderStatus_AmendPending
        Case "AA"                       ' Awaiting broker cancellation confirmation
            GenesisOrderStatus = eTT_OrderStatus_CancelPending
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.GenesisOrderStatus", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderType
'' Description: Convert Photon order type to Genesis order type
'' Inputs:      Photon Order Type
'' Returns:     Genesis Order Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderType(ByVal strOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Select Case strOrderType
        Case "M", "B"
            GenesisOrderType = eTT_OrderType_Market
        Case "L", "T", "O"
            GenesisOrderType = eTT_OrderType_Limit
        Case "S"
            GenesisOrderType = eTT_OrderType_Stop
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.GenesisOrderType", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateToDouble
'' Description: Convert a YYYYMMDDHHMMSS to a double
'' Inputs:      Date from the GetData call
'' Returns:     Date as double
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DateToDouble(ByVal strDate As String) As Double
On Error GoTo ErrSection:

    Dim dTime As Double                 ' Time portion of the date
    
    dTime = Val(Mid(strDate, 9, 2)) * 3600 + Val(Mid(strDate, 11, 2)) * 60 + Val(Right(strDate, 2))
    dTime = dTime / 86400#
    
    DateToDouble = JulFromLong(Val(Left(strDate, 8))) + dTime

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.DateToDouble", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderAnswer
'' Description: Handle the OrderAnswer callback from Photon servers
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrderAnswer(ByVal strMessage As String, Optional ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Order broken out into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim strDebug As String              ' String to print
    Dim strGenesisOrderID As String     ' Genesis Order ID
    Dim strBrokerID As String           ' Broker ID
    Dim strStatus As String             ' Photon Order Status
    Dim strPrevOrderText As String      ' Order text for previous order
    Dim strOrderText As String          ' Order text for current order
    Dim Order As New cPtOrder           ' Order object
    Dim lOrderID As Long                ' Order ID back from the order
    Dim Fill As New cPtFill             ' Fill object
    Dim strBrokerFillID As String       ' Broker Fill ID
    Dim strOldBrokerFillID As String    ' Old version of the broker fill ID
    Dim lFillIndex As Long              ' Index of the fill in the orders
    Dim dMult As Double                 ' Multiplier to get to our price
    Dim dPrice As Double                ' Fill Price
    Dim lQuantity As Long               ' Fill Quantity
    Dim nStatus As eTT_OrderStatus      ' Genesis order status
    Dim lPos As Long                    ' Position in array
    Dim strErrorMsg As String            ' Error message
    Dim bModify As Boolean              ' Are we coming off of a modify?
    Dim strAccountNumber As String      ' Account number
    Dim lAccountID As Long              ' Account ID for the given account number
    Dim bNewError As Boolean            ' New error message?
    Dim Bars As New cGdBars             ' Bars object
    Dim nOrderType As eTT_OrderType     ' Order Type
        
    bNewError = False
    astrOrder.SplitFields strMessage, vbTab
    
    strDebug = "OrderAnswer: "
    For lIndex = 0 To 77
        strDebug = strDebug & Str(lIndex) & "=" & astrOrder(lIndex) & ";"
    Next lIndex
    DumpDebug vbTab & Left(strDebug, Len(strDebug) - 1)
    
    strAccountNumber = astrOrder(10)
    If Len(strAccountNumber) > 0 Then
        lAccountID = CreateAccountFromNumber(strAccountNumber, m.nBroker)
        strGenesisOrderID = astrOrder(11)
        strBrokerID = astrOrder(12)
        strStatus = astrOrder(21)
    
        Set Order = LoadOrder(strAccountNumber, strGenesisOrderID, strBrokerID)
        If Order Is Nothing Then Set Order = New cPtOrder
        
        bModify = False
        If (m.ModifiedOrders.Exists(strGenesisOrderID) = True) And (GenesisOrderStatus(strStatus) = eTT_OrderStatus_Rejected) Then
            bModify = True
            strPrevOrderText = m.ModifiedOrders(strGenesisOrderID).OrderText
            strOrderText = Order.OrderText
            m.ModifiedOrders(strGenesisOrderID).Save
            
            m.BrokerInfo.AddOrder m.ModifiedOrders(strGenesisOrderID), bRefresh
            
            OrderCallback m.ModifiedOrders(strGenesisOrderID)
            bNewError = True
        ElseIf Val(astrOrder(52)) > Val(Order.BrokerID) Then
            With Order
                If .OrderID = 0 Then
                    .AccountID = lAccountID
                    .GenesisOrderID = strGenesisOrderID
                    ''.BrokerID = strBrokerID
                    .Buy = (astrOrder(7) = "0")
                    .DateSent = DateToDouble(astrOrder(22))
                    Select Case astrOrder(9)
                        Case "D"
                            .Expiration = JulFromLong(Val(astrOrder(55)))
                        Case "J"
                            .Expiration = Date * -1&
                        Case Else
                            .Expiration = 0
                    End Select
                    .SymbolOrSymbolID = GenesisSymbol(astrOrder(5), astrOrder(0), dMult)
                    .OrderDate = DateToDouble(astrOrder(25))
                    nOrderType = GenesisOrderType(astrOrder(8))
                    If Not ((nOrderType = eTT_OrderType_Market) And (.OrderType = eTT_OrderType_MIT)) Then
                        .OrderType = nOrderType
                        .LimitPrice = Val(astrOrder(54)) / dMult
                        .StopPrice = Val(astrOrder(58)) / dMult
                    End If
                    .Quantity = CLng(Val(astrOrder(59))) + .FillQuantity - CLng(Val(astrOrder(64)))
                Else
                    GenesisSymbol astrOrder(5), astrOrder(0), dMult
                End If
                SetBarProperties Bars, .Symbol
                
                .BrokerID = astrOrder(52)
                
                nStatus = GenesisOrderStatus(strStatus)
                If (nStatus = eTT_OrderStatus_Cancelled) And (.Status = eTT_OrderStatus_ParkPending) Then
                    .Status = eTT_OrderStatus_Parked
                ElseIf (nStatus = eTT_OrderStatus_Cancelled) And (.Status = eTT_OrderStatus_Parked) Then
                    .Status = eTT_OrderStatus_Parked
                Else
                    .Status = nStatus
                End If
                .DateExchAckn = DateToDouble(astrOrder(42))
                If IsDigit(strBrokerID, 1) Then .ExchangeID = strBrokerID
                .Message = astrOrder(49)
                
                If nStatus = eTT_OrderStatus_Rejected Then
                    strOrderText = Order.OrderText
                    If m.astrErrorCodes.BinarySearch(Replace(astrOrder(12), " ", "") & " - ", lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                        If m.astrErrorCodes(lPos) <> .Message Then
                            .Message = m.astrErrorCodes(lPos)
                            bNewError = True
                        End If
                    Else
                        If Replace(astrOrder(12), " ", "") <> .Message Then
                            .Message = Replace(astrOrder(12), " ", "")
                            bNewError = True
                        End If
                    End If
                End If
                
                .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.OrderDate, m.nBroker, .Symbol, False))
                .IsSnapshot = True
                
                .StatusDate = DateToDouble(astrOrder(22))
            End With
            
            dPrice = Val(astrOrder(63))
            If dPrice = 0# Then dPrice = Val(astrOrder(54))
            lQuantity = CLng(Val(astrOrder(64)))
            If lQuantity = 0& Then lQuantity = CLng(Val(astrOrder(53)))
            
            If (Right(strStatus, 1) = "X") And (dPrice <> 0# And lQuantity <> 0&) Then
                strOldBrokerFillID = strBrokerID & " " & astrOrder(52)
                strBrokerFillID = strBrokerID & " F" & astrOrder(60)
                lFillIndex = -1&
                For lIndex = 1 To Order.Fills.Count
                    If (Order.Fills(lIndex).BrokerID = strBrokerFillID) Or (Order.Fills(lIndex).BrokerID = strOldBrokerFillID) Then
                        lFillIndex = lIndex
                        Exit For
                    End If
                Next lIndex
                
                If lFillIndex = -1& Then
                    Set Fill = New cPtFill
                
                    With Fill
                        .AccountID = Order.AccountID
                        .AutoTradingItemID = Order.AutoTradeItemID
                        .BrokerID = strBrokerFillID
                        .BrokerOrderID = Order.BrokerID
                        .Buy = Order.Buy
                        .FillDate = DateToDouble(astrOrder(36))
                        .IsManual = False
                        .IsSnapshot = True
                        .OrderID = Order.OrderID
                        .Price = dPrice / dMult
                        .Quantity = lQuantity
                        .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.FillDate, m.nBroker, .Symbol, False))
                        .Symbol = Order.Symbol
                        .SymbolID = Order.SymbolID
                        .Save
                    End With
                    
                    Set Order.Fills(Str(Fill.FillID)) = Fill
                    Order.Quantity = CLng(Val(astrOrder(59))) + Order.FillQuantity
                    Order.Save
                    
                    m.BrokerInfo.AddOrder Order, bRefresh
                    m.BrokerInfo.AddFill Fill, bRefresh
                    
                    FillCallback Order, Fill
                    g.ActivityLogs.AddFillToActivityLog m.nBroker, Fill, False
                    DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                    DumpDebug Fill.FillText
                ElseIf (Order.Fills(lFillIndex).BrokerID = strOldBrokerFillID) Then
                    Order.Fills(lFillIndex).BrokerID = strBrokerFillID
                    Order.Save
                    
                    m.BrokerInfo.AddOrder Order, bRefresh
                    
                    OrderCallback Order
                    DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                Else
                    Order.Save
                    
                    m.BrokerInfo.AddOrder Order, bRefresh
                    
                    OrderCallback Order
                    DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                End If
            Else
                Order.Save
                
                m.BrokerInfo.AddOrder Order, bRefresh
                
                OrderCallback Order
                g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
            End If
        End If
        
        If m.ModifiedOrders.Exists(strGenesisOrderID) Then
            m.ModifiedOrders.Remove strGenesisOrderID
        End If
        
        If (GenesisOrderStatus(strStatus) = eTT_OrderStatus_Rejected) And (bNewError = True) Then
            If m.astrErrorCodes.BinarySearch(Replace(astrOrder(12), " ", "") & " - ", lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                strErrorMsg = m.astrErrorCodes(lPos)
            Else
                strErrorMsg = Replace(astrOrder(12), " ", "")
            End If
            
            If bModify = True Then
                If Len(strOrderText) > 0 Then
                    InfBox "Your attempt to modify|" & strPrevOrderText & "|to|" & strOrderText & "|was rejected for the following reason:||" & strErrorMsg & "||The original order has been restored.|", , , "Order Modification Rejected"
                    DumpDebug "Your attempt to modify '" & strPrevOrderText & "' to '" & strOrderText & "' was rejected for the following reason: '" & strErrorMsg & "'.  The original order has been restored."
                Else
                    InfBox "Your attempt to modify order to|" & strPrevOrderText & "|was rejected for the following reason:||" & strErrorMsg & "||The original order has been restored.|", , , "Order Modification Rejected"
                    DumpDebug "Your attempt to modify order to '" & strPrevOrderText & "' was rejected for the following reason: '" & strErrorMsg & "'.  The original order has been restored."
                End If
            Else
                InfBox "The order|" & strOrderText & "|was rejected for the following reason:||" & strErrorMsg & "|", , , "Order Rejected"
                DumpDebug "The order '" & strOrderText & "' was rejected for the following reason: '" & strErrorMsg & "'."
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.OrderAnswer", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming App Mail message from the Photon program
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDPhotonMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strTemp As String               ' Temporary string
    Dim lPos As Long                    ' Position of an item in an array
    Dim astrSymbols As New cGdArray     ' Array of symbols
    Dim lIndex As Long                  ' Index into a for loop
    Dim strGenesisSymbol As String      ' Genesis version of symbol

    'If (nType <> eGDPhotonMessageType_DisconnectedSLE) And (nType <> eGDPhotonMessageType_DisconnectedSLC) Then
    '    m.dHeartBeat = gdTickCount
    'End If
    
    If (nType <> eGDPhotonMessageType_DepthChanged) And (nType <> eGDPhotonMessageType_PriceChanged) And (nType <> eGDPhotonMessageType_CNXSLC) And (nType <> eGDPhotonMessageType_CNXSLE) And (nType <> eGDPhotonMessageType_ConnectionInfo) Then
        DumpDebug vbTab & "Photon Message Received (" & MessageTypeToString(nType) & "): " & strMessage
    End If

    Select Case nType
        Case eGDPhotonMessageType_CNXSLE
            m.dHeartBeat = gdTickCount
                
        Case eGDPhotonMessageType_OrderAnswer
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingOrders = True
                m.BrokerInfo.ClearOrders
            ElseIf UCase(strMessage) = "END" Then
                m.bRefreshingOrders = False
                m.BrokerInfo.RefreshingOrders = False
            Else
                'UpdateBrokerInfoOrder strMessage, False
                OrderAnswer strMessage, m.bRefreshingOrders
            End If
        
        Case eGDPhotonMessageType_AppLoaded
            m.bAppLoaded = True
            SendPhotonMessage eGDPhotonMessageType_Connect, ConnectString
        
        Case eGDPhotonMessageType_AppUnloaded
            m.bAppLoaded = False
        
        Case eGDPhotonMessageType_Heartbeat
            m.dHeartBeat = gdTickCount
        
        Case eGDPhotonMessageType_CNXSLC
            m.dHeartBeat = gdTickCount
        
        Case eGDPhotonMessageType_ConnectionInfo
            HandleConnectionInfo strMessage
            If (Not TotallyConnected) And (m.bManualDisconnect = True) Then m.dHeartBeat = 0#
            
        Case eGDPhotonMessageType_DepthChanged
            m.dHeartBeat = gdTickCount
            HandleDepthChange strMessage
            
        Case eGDPhotonMessageType_PriceChanged
            m.dHeartBeat = gdTickCount
            HandlePriceChange strMessage
            
        Case eGDPhotonMessageType_Subscribed
            m.dHeartBeat = gdTickCount
            astrSymbols.SplitFields strMessage, ","
            For lIndex = 0 To astrSymbols.Size - 1
                If m.astrSubContract.BinarySearch(astrSymbols(lIndex), lPos) = False Then
                    m.astrSubContract.Add astrSymbols(lIndex), lPos
                End If
                
                strGenesisSymbol = GenesisSymbol(Parse(astrSymbols(lIndex), "|", 1), Parse(astrSymbols(lIndex), "|", 2))
                If Len(strGenesisSymbol) > 0 Then
                    g.RealTime.AddBrokerRtSymbol Parse(strGenesisSymbol, "-", 1) & "-", "Photon/D"
                End If
            Next lIndex
        
        Case eGDPhotonMessageType_Unsubscribed
            m.dHeartBeat = gdTickCount
            astrSymbols.SplitFields strMessage, ","
            For lIndex = 0 To astrSymbols.Size - 1
                If m.astrSubContract.BinarySearch(astrSymbols(lIndex), lPos) = True Then
                    m.astrSubContract.Remove lPos
                End If
            Next lIndex
    
    End Select
    
    If Refreshing And (m.dRefreshingStart = 0#) Then
        m.dRefreshingStart = gdTickCount
    ElseIf Refreshing And (gdTickCount > m.dRefreshingStart + 30000#) Then
        m.bRefreshingOrders = False
    ElseIf (Not Refreshing) And (m.dRefreshingStart <> 0#) Then
        m.dRefreshingStart = 0#
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.HandleMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendPhotonMessage
'' Description: Send an App Mail message to the Photon program
'' Inputs:      Type of Message, Message, Send Now?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendPhotonMessage(ByVal nType As eGDPhotonMessageTypes, ByVal strMessage As String, Optional ByVal bSendNow As Boolean = True)
On Error GoTo ErrSection:

    Dim astrMessage As New cGdArray     ' Message split out into an array

    If nType = eGDPhotonMessageType_Connect Then
        astrMessage.SplitFields strMessage, vbTab
        astrMessage.Remove 1
        DumpDebug vbTab & "Sending Photon Message (" & MessageTypeToString(nType) & "): " & astrMessage.JoinFields(vbTab)
    Else
        DumpDebug vbTab & "Sending Photon Message (" & MessageTypeToString(nType) & "): " & strMessage
    End If
    
    frmOnlineBroker.gdBroker.CreateMessage "Photon", nType, strMessage, , bSendNow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.SendPhotonMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Connected
'' Description: Determine if any exchanges that are supposed to be connected are
'' Inputs:      None
'' Returns:     True if Connected, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Connected() As Boolean
On Error GoTo ErrSection:

    Dim bConnected As Boolean           ' Are we connected?
    
    If m.bConnectToCME And m.bConnectedToCME Then
        bConnected = True
    ElseIf m.bConnectToCBOT And m.bConnectedToCBOT Then
        bConnected = True
    ElseIf m.bConnectToEurex And m.bConnectedToEurex Then
        bConnected = True
    ElseIf m.bConnectToLiffe And m.bConnectedToLiffe Then
        bConnected = True
    ElseIf m.bConnectToMarket And m.bConnectedToMarket Then
        bConnected = True
    Else
        bConnected = False
    End If
    
    Connected = bConnected

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.Connected"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TotallyConnected
'' Description: Determine if all exchanges that are supposed to be connected are
'' Inputs:      None
'' Returns:     True if Totally Connected, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TotallyConnected() As Boolean
On Error GoTo ErrSection:

    Dim bConnected As Boolean           ' Are we connected?
    
    bConnected = True
    
    If m.bConnectToCME And (Not m.bConnectedToCME) Then bConnected = False
    If m.bConnectToCBOT And (Not m.bConnectedToCBOT) Then bConnected = False
    If m.bConnectToEurex And (Not m.bConnectedToEurex) Then bConnected = False
    If m.bConnectToLiffe And (Not m.bConnectedToLiffe) Then bConnected = False
    If m.bConnectToMarket And (Not m.bConnectedToMarket) Then bConnected = False
    
    TotallyConnected = bConnected

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.TotallyConnected"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyPositions
'' Description: Verify positions with the user
'' Inputs:      Account, Symbol, Force?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub VerifyPositions(Optional ByVal strAccount As String = "", Optional ByVal strSymbol As String = "", Optional ByVal bForce As Boolean = False)
On Error GoTo ErrSection:
    
    Dim bReturn As Boolean              ' Return value from the position confirm

    If m.bVerifyPositions Or bForce Then
        ' Only do this if there is not a modal dialog up...
        If frmMain.Enabled Then
            m.bVerifyPositions = False
        
            bReturn = frmPositionConfirm.ShowMe(m.nBroker, strAccount, strSymbol, True, True)
            If (Len(strAccount) = 0) And (Len(strSymbol) = 0) Then
                m.bVerifyPositions = Not bReturn
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.VerifyPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectString
'' Description: Build the connect string to pass to a connect message
'' Inputs:      None
'' Returns:     Connection string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ConnectString() As String
On Error GoTo ErrSection:

    Dim astrMessage As New cGdArray     ' Message to send to the app mail client
    
    astrMessage.Create eGDARRAY_Strings
    astrMessage(0) = Str(m.lUserID)
    astrMessage(1) = m.strPassword
    astrMessage(2) = m.strServerIP
    astrMessage(3) = Str(m.iServerPort)
    If m.bConnectToMarket Then astrMessage(4) = "1" Else astrMessage(4) = "0"
    astrMessage(5) = Str(m.lMarketNode)
    If m.bConnectToCME Then astrMessage(6) = "1" Else astrMessage(6) = "0"
    astrMessage(7) = Str(m.lCmeNode)
    If m.bConnectToCBOT Then astrMessage(8) = "1" Else astrMessage(8) = "0"
    astrMessage(9) = Str(m.lCbotNode)
    If m.bConnectToEurex Then astrMessage(10) = "1" Else astrMessage(10) = "0"
    astrMessage(11) = Str(m.lEurexNode)
    If m.bConnectToLiffe Then astrMessage(12) = "1" Else astrMessage(12) = "0"
    astrMessage(13) = Str(m.lLiffeNode)
    astrMessage(14) = m.strEuropeIP
    astrMessage(15) = Str(m.iEuropePort)
    
    ConnectString = astrMessage.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.ConnectString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetConnectionProperties
'' Description: Get the connection properties from the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetConnectionProperties() As Boolean
On Error GoTo ErrSection:

    Dim strConnect As String            ' Original connect string
    Dim lTempID As Long                 ' Temporary user id

    strConnect = ConnectString
    
    lTempID = GetIniFileProperty("UserID", 0&, "User", m.strIniFile)
    If lTempID <> m.lUserID Then m.strPassword = ""
    m.lUserID = lTempID
    m.strIniPassword = GetIniFileProperty("Password", "", "User", m.strIniFile)
    If Len(m.strIniPassword) > 0 Then
        SetIniFileProperty "Password", "", "User", m.strIniFile
        SetIniFileProperty "Password2", EncryptToHex(m.strIniPassword), "User", m.strIniFile
    Else
        m.strIniPassword = DecryptFromHex(GetIniFileProperty("Password2", "", "User", m.strIniFile))
    End If
    
    m.strServerIP = GetIniFileProperty("ServerIP", "198.63.193.40", "Server", m.strIniFile)
    m.iServerPort = GetIniFileProperty("ServerPort", 9030, "Server", m.strIniFile)
    m.strEuropeIP = GetIniFileProperty("ServerIPEurope", "209.120.215.41", "Server", m.strIniFile)
    m.iEuropePort = GetIniFileProperty("ServerPortEurope", 9030, "Server", m.strIniFile)
    
    m.lMarketNode = GetIniFileProperty("Market", 7247, "SubNodes", m.strIniFile)
    m.lCmeNode = GetIniFileProperty("CME", 18249, "SubNodes", m.strIniFile)
    m.lCbotNode = GetIniFileProperty("eCBOT", 18217, "SubNodes", m.strIniFile)
    m.lEurexNode = GetIniFileProperty("Eurex", 18263, "SubNodes", m.strIniFile)
    m.lLiffeNode = GetIniFileProperty("Liffe", 12874, "SubNodes", m.strIniFile)
    
    m.bConnectToMarket = (GetIniFileProperty("Market", vbChecked, "Connect", m.strIniFile) = vbChecked)
    m.bConnectToCME = (GetIniFileProperty("CME", vbChecked, "Connect", m.strIniFile) = vbChecked)
    m.bConnectToCBOT = (GetIniFileProperty("eCBOT", vbChecked, "Connect", m.strIniFile) = vbChecked)
    m.bConnectToEurex = (GetIniFileProperty("Eurex", vbUnchecked, "Connect", m.strIniFile) = vbChecked)
    m.bConnectToLiffe = (GetIniFileProperty("Liffe", vbUnchecked, "Connect", m.strIniFile) = vbChecked)
    
    GetConnectionProperties = (ConnectString <> strConnect)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.GetConnectionProperties"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PhotonSymbol
'' Description: Convert a Genesis Symbol to a Photon symbol
'' Inputs:      Genesis Symbol
'' Returns:     Photon Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PhotonSymbol(ByVal strGenesisSymbol As String, Optional strPhotonExchange As String, Optional dMult As Double) As String
On Error GoTo ErrSection:

    Dim strBase As String               ' Base symbol
    Dim strContract As String           ' Contract
    Dim strReturn As String             ' Return from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim lMonth As Long                  ' Month from the contract
    Dim lYear As Long                   ' Year from the contract
    
    strBase = Parse(strGenesisSymbol, "-", 1)
    strContract = Parse(strGenesisSymbol, "-", 2)
    lMonth = Val(strContract) Mod 100
    lYear = Val(strContract) \ 100
    strReturn = ""
    
    For lIndex = 0 To m.astrBaseSyms.Size - 1
        If Parse(m.astrBaseSyms(lIndex), vbTab, 3) = strBase Then
            strReturn = Parse(m.astrBaseSyms(lIndex), vbTab, 1)
            strPhotonExchange = Parse(m.astrBaseSyms(lIndex), vbTab, 2)
            dMult = Val(Parse(m.astrBaseSyms(lIndex), vbTab, 4))
            
            Select Case strPhotonExchange
                Case "37"           ' CME
                    strReturn = strReturn & MonthToCode(lMonth) & Right(Str(lYear), 1)
                    
                Case "96", "41"     ' Cbot, Liffe
                    strReturn = Pad(strReturn, 5, "L") & Mid(strContract, 3, 2) & Right(strContract, 2) & "00000F"
                
                Case "30"           ' Eurex
                    strReturn = Pad(strReturn, 4, "L") & Right(strContract, 2) & Mid(strContract, 3, 2)
            
            End Select
            
            Exit For
        End If
    Next lIndex

    PhotonSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.PhotonSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Convert a Photon Symbol to a Genesis symbol
'' Inputs:      Photon Symbol, Photon Exchange
'' Returns:     Genesis Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisSymbol(ByVal strPhotonSymbol As String, ByVal strPhotonExchange As String, Optional dMult As Double) As String
On Error GoTo ErrSection:

    Dim strBase As String               ' Base portion of the symbol
    Dim strContract As String           ' Contract portion of the symbol
    Dim strReturn As String             ' Return value from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim lYear As Long                   ' Year from the contract
    Dim lMonth As Long                  ' Month from the contract
    
    Select Case strPhotonExchange
        Case "37"           ' Cme
            strBase = Left(strPhotonSymbol, Len(strPhotonSymbol) - 2)
            strContract = Right(strPhotonSymbol, 2)
            
        Case "96", "41"     ' Cbot, Liffe
            strBase = Trim(Left(strPhotonSymbol, 5))
            strContract = MonthToCode(Val(Mid(strPhotonSymbol, 8, 2))) & Mid(strPhotonSymbol, 7, 1)
            
        Case "30"           ' Eurex
            strBase = Trim(Left(strPhotonSymbol, 4))
            strContract = MonthToCode(Val(Mid(strPhotonSymbol, 5, 2))) & Right(strPhotonSymbol, 1)
    
    End Select
    
    For lIndex = 0 To m.astrBaseSyms.Size - 1
        If Parse(m.astrBaseSyms(lIndex), vbTab, 1) = strBase Then
            strReturn = Parse(m.astrBaseSyms(lIndex), vbTab, 3)
            dMult = Val(Parse(m.astrBaseSyms(lIndex), vbTab, 4))
            
            lMonth = CodeToMonth(Left(strContract, 1))
            lYear = CLng(Val(Right(strContract, 1))) + 2000
            Do While (lYear * 100 + lMonth) < (Year(Date) * 100 + Month(Date))
                lYear = lYear + 10
            Loop
            
            strReturn = strReturn & "-" & Str(lYear) & Format(lMonth, "00")
            Exit For
        End If
    Next lIndex
    
    GenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.GenesisSymbol"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrder
'' Description: Load the order from the database for Genesis ID or Broker ID
'' Inputs:      Account Number, Genesis Order ID, Broker ID
'' Returns:     Order (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadOrder(ByVal strAccountNumber As String, ByVal strGenesisOrderID As String, ByVal strBrokerID As String) As cPtOrder
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lAccountID As Long              ' Account ID for the given account number
    Dim Order As New cPtOrder           ' Order to return
    Dim bFound As Boolean               ' Have we found the order?

    bFound = False
    lAccountID = g.Broker.AccountIDForNumber(strAccountNumber)
    If Len(strGenesisOrderID) > 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                    "WHERE [GenesisOrderID]='" & strGenesisOrderID & "' AND [AccountID]=" & lAccountID & ";", dbOpenDynaset)
        If Not (rs.EOF And rs.BOF) Then
            Set Order = New cPtOrder
            If Order.Load(rs!OrderID) Then
                Set LoadOrder = Order
                bFound = True
            End If
        End If
    End If
    If (bFound = False) And (Len(strBrokerID) > 0) Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                    "WHERE [BrokerOrderID]='" & strBrokerID & "' AND [AccountID]=" & lAccountID & ";", dbOpenDynaset)
        If Not (rs.EOF And rs.BOF) Then
            Set Order = New cPtOrder
            If Order.Load(rs!OrderID) Then
                Set LoadOrder = Order
                bFound = True
            End If
        End If
    End If
    
    If bFound Then
        Set LoadOrder = Order
    Else
        Set LoadOrder = Nothing
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.LoadOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisLog
'' Description: Send a string to the log file for the day
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GenesisLog(ByVal strMessage As String)
On Error Resume Next

    Dim fh As Integer                   ' File handle to open file with

    fh = FreeFile
    Open AddSlash(App.Path) & "Photon\TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append Shared As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " - " & strMessage
        Close #fh
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleConnectionInfo
'' Description: Do the appropriate thing with the given connection information
'' Inputs:      Connection Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleConnectionInfo(ByVal strConnectionInfo As String)
On Error GoTo ErrSection:

    Dim dTickCount As Double            ' Current tick count from message
    Dim astrStatus As New cGdArray      ' Connection status' for each exchange
    Dim astrError As New cGdArray       ' Error messages for each exchange
    Dim nStatus As eGDConnectionStatus  ' Connection status for the exchange
    Dim strError As String              ' Error string
    Dim bInProgress As Boolean          ' Is at least one exchange still in progress?
    Dim bConnected As Boolean           ' Is at least one exchange connected?
    Dim nNewStatus As eGDConnectionStatus ' New connection status
    Static strLastError As String       ' Last error message displayed

    ' Get information from the connection message...
    dTickCount = Val(Parse(strConnectionInfo, vbTab, 1))
    astrStatus.SplitFields Parse(strConnectionInfo, vbTab, 2), ","
    astrError.SplitFields Parse(strConnectionInfo, vbTab, 3), ","
    bInProgress = False
    bConnected = False
    strError = ""
    
    ' Market data connection status...
    nStatus = CLng(Val(astrStatus(0)))
    Select Case nStatus
        Case eGDConnectionStatus_Disconnected
            m.bConnectedToMarket = False
            m.nDataStatus = eGDConnectionStatus_Disconnected
            If Len(astrError(0)) > 0 Then strError = strError & "Data: " & astrError(0) & "|"
            
            If g.RealTime.UseBrokerFeed = m.nBroker Then
                g.RealTime.UseBrokerFeed = 0
            End If
            m.astrSubContract.Size = 0
            
        Case eGDConnectionStatus_Disconnecting
            m.bConnectedToMarket = False
            m.nDataStatus = eGDConnectionStatus_Disconnecting
            bInProgress = True
            
        Case eGDConnectionStatus_Connecting
            m.bConnectedToMarket = False
            m.nDataStatus = eGDConnectionStatus_Connecting
            bInProgress = True
            
        Case eGDConnectionStatus_Connected
            m.bConnectedToMarket = True
            m.nDataStatus = eGDConnectionStatus_Connected
            
            If g.RealTime.UseBrokerFeed = 0 Then
                g.RealTime.UseBrokerFeed = m.nBroker
            End If
            
            bConnected = True
    End Select
    
    ' CME Connection status...
    nStatus = CLng(Val(astrStatus(1)))
    Select Case nStatus
        Case eGDConnectionStatus_Disconnected
            m.bConnectedToCME = False
            m.nCmeStatus = eGDConnectionStatus_Disconnected
            If Len(astrError(1)) > 0 Then strError = strError & "CME: " & astrError(1) & "|"
            
        Case eGDConnectionStatus_Disconnecting
            m.bConnectedToCME = False
            m.nCmeStatus = eGDConnectionStatus_Disconnecting
            bInProgress = True
            
        Case eGDConnectionStatus_Connecting
            m.bConnectedToCME = False
            m.nCmeStatus = eGDConnectionStatus_Connecting
            bInProgress = True
            
        Case eGDConnectionStatus_Connected
            m.bConnectedToCME = True
            m.nCmeStatus = eGDConnectionStatus_Connected
            bConnected = True
    End Select

    ' Cbot Connection status...
    nStatus = CLng(Val(astrStatus(2)))
    Select Case nStatus
        Case eGDConnectionStatus_Disconnected
            m.bConnectedToCBOT = False
            m.nCbotStatus = eGDConnectionStatus_Disconnected
            If Len(astrError(2)) > 0 Then strError = strError & "Cbot: " & astrError(2) & "|"
            
        Case eGDConnectionStatus_Disconnecting
            m.bConnectedToCBOT = False
            m.nCbotStatus = eGDConnectionStatus_Disconnecting
            bInProgress = True
            
        Case eGDConnectionStatus_Connecting
            m.bConnectedToCBOT = False
            m.nCbotStatus = eGDConnectionStatus_Connecting
            bInProgress = True
            
        Case eGDConnectionStatus_Connected
            m.bConnectedToCBOT = True
            m.nCbotStatus = eGDConnectionStatus_Connected
            bConnected = True
    End Select

    ' Eurex Connection status...
    nStatus = CLng(Val(astrStatus(3)))
    Select Case nStatus
        Case eGDConnectionStatus_Disconnected
            m.bConnectedToEurex = False
            m.nEurexStatus = eGDConnectionStatus_Disconnected
            If Len(astrError(3)) > 0 Then strError = strError & "Eurex: " & astrError(3) & "|"
            
        Case eGDConnectionStatus_Disconnecting
            m.bConnectedToEurex = False
            m.nEurexStatus = eGDConnectionStatus_Disconnecting
            bInProgress = True
            
        Case eGDConnectionStatus_Connecting
            m.bConnectedToEurex = False
            m.nEurexStatus = eGDConnectionStatus_Connecting
            bInProgress = True
            
        Case eGDConnectionStatus_Connected
            m.bConnectedToEurex = True
            m.nEurexStatus = eGDConnectionStatus_Connected
            bConnected = True
    End Select

    ' Liffe Connection status...
    nStatus = CLng(Val(astrStatus(4)))
    Select Case nStatus
        Case eGDConnectionStatus_Disconnected
            m.bConnectedToLiffe = False
            m.nLiffeStatus = eGDConnectionStatus_Disconnected
            If Len(astrError(4)) > 0 Then strError = strError & "Liffe: " & astrError(4) & "|"
            
        Case eGDConnectionStatus_Disconnecting
            m.bConnectedToLiffe = False
            m.nLiffeStatus = eGDConnectionStatus_Disconnecting
            bInProgress = True
            
        Case eGDConnectionStatus_Connecting
            m.bConnectedToLiffe = False
            m.nLiffeStatus = eGDConnectionStatus_Connecting
            bInProgress = True
            
        Case eGDConnectionStatus_Connected
            m.bConnectedToLiffe = True
            m.nLiffeStatus = eGDConnectionStatus_Connected
            bConnected = True
    End Select
    
    If bInProgress Then
        m.dLastStatusMsg = dTickCount
        If ConnectionStatus = eGDConnectionStatus_Connected Then
            nNewStatus = eGDConnectionStatus_Disconnecting
        ElseIf ConnectionStatus = eGDConnectionStatus_Disconnected Then
            nNewStatus = eGDConnectionStatus_Connecting
        End If
    ElseIf bConnected Then
        m.dLastStatusMsg = dTickCount
        If TotallyConnected Then
            If ConnectionStatus <> eGDConnectionStatus_Connected Then
                SetIniFileProperty "Last", EncryptToHex(Str(Date)), "Connect", m.strIniFile
                DumpDebug "Connected to Photon"
                GetOrders
            End If
            nNewStatus = eGDConnectionStatus_Connected
        ElseIf ConnectionStatus = eGDConnectionStatus_Connected Then
            nNewStatus = eGDConnectionStatus_Disconnecting
        ElseIf ConnectionStatus = eGDConnectionStatus_Disconnected Then
            nNewStatus = eGDConnectionStatus_Connecting
        End If
    Else
        If ConnectionStatus <> eGDConnectionStatus_Disconnected Then
            DumpDebug "Disconnected from Photon"
            If m.bManualDisconnect Then
                m.dLastStatusMsg = 0#
                m.bVerifyPositions = True
            Else
                m.dLastStatusMsg = dTickCount - 120000
            End If
            
            m.bRefreshingOrders = False
        End If
        nNewStatus = eGDConnectionStatus_Disconnected
    End If
    
    If nNewStatus <> ConnectionStatus Then
        ConnectionStatus = nNewStatus
    End If
    
    If (Len(strError) > 0) And (strError <> strLastError) Then
        InfBox "Error connecting to Photon:|" & strError & "|", "!", "+-OK", "Photon Connection Error", True, , , , , , , eGDAlign_Left
    End If
    strLastError = strError
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.HandleConnectionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandlePriceChange
'' Description: Handle a price change message from the Photon servers
'' Inputs:      Price Information
'' Returns:     None
''
'' Message Format: Symbol<t>Exch<t>Price<t>Size<t>Time<t>TickCountRecd<t>TickCountSent
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandlePriceChange(ByVal strPrice As String)
On Error GoTo ErrSection:

    Dim strGenesisSymbol As String      ' Genesis symbol for symbol passed in
    Dim strPhotonExchange As String     ' Photon exchange
    Dim dMult As Double                 ' Price multiplier
    Dim dExchTime As Double             ' Genesis exchange time
    Dim dGenesisPrice As Double         ' Genesis price
    Dim dFeedTime As Double             ' Genesis Feed time

    ' Convert symbol...
    strPhotonExchange = Parse(strPrice, vbTab, 2)
    strGenesisSymbol = GenesisSymbol(Parse(strPrice, vbTab, 1), strPhotonExchange, dMult)
    If Len(strGenesisSymbol) > 0 Then
        dFeedTime = g.RealTime.FeedTime(strGenesisSymbol)
    
        ' Convert time to NY for domestic exchanges...
        If strPhotonExchange = "37" Or strPhotonExchange = "96" Then
            dExchTime = ConvertTimeZone(Int(dFeedTime) + Val(Parse(strPrice, vbTab, 5)), "CHI", "NY")
        Else
            dExchTime = Int(dFeedTime) + Val(Parse(strPrice, vbTab, 5))
        End If
        If dExchTime > dFeedTime + 0.5 Then
            dExchTime = dExchTime - 1
        ElseIf dExchTime < dFeedTime - 0.5 Then
            dExchTime = dExchTime + 1
        End If
        
        ' Convert price to Genesis price...
        dGenesisPrice = Val(Parse(strPrice, vbTab, 3)) / dMult
    
        ' Send to realtime object (Genesis Symbol<t>Exchange Time<t>Price<t>Size<t>)...
        g.RealTime.StreamedMessage 20, strGenesisSymbol & vbTab & Str(dExchTime) & vbTab & Str(dGenesisPrice) & vbTab & Parse(strPrice, vbTab, 4) & vbTab & Parse(strPrice, vbTab, 6) & vbTab & Parse(strPrice, vbTab, 7), m.nBroker
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.HandlePriceChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleDepthChange
'' Description: Handle a depth of market change message from the Photon servers
'' Inputs:      Depth Information
'' Returns:     None
''
'' Message Format: Symbol<t>Exch<t>Level<t>Bid<t>BidSize<t>Ask<t>AskSize
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleDepthChange(ByVal strDepth As String)
On Error GoTo ErrSection:

    Dim strGenesisSymbol As String      ' Genesis symbol for symbol passed in
    Dim strPhotonExchange As String     ' Photon exchange
    Dim dMult As Double                 ' Price multiplier
    Dim dExchTime As Double             ' Genesis exchange time
    Dim dGenesisBid As Double           ' Genesis price
    Dim dGenesisAsk As Double           ' Genesis price
    Dim strExchTimeZone As String       ' Exchange time zone from the Bars
    Dim Bars As cGdBars                 ' Temporary bars object
    Dim lLevel As Long                  ' Level for the depth of market

    ' Convert symbol...
    strPhotonExchange = Parse(strDepth, vbTab, 2)
    strGenesisSymbol = GenesisSymbol(Parse(strDepth, vbTab, 1), strPhotonExchange, dMult)
    If Len(strGenesisSymbol) > 0 Then
        ' Determine the exchange time to pass...
        dExchTime = g.RealTime.FeedTime
        Set Bars = New cGdBars
        SetBarProperties Bars, strGenesisSymbol
        strExchTimeZone = Bars.Prop(eBARS_ExchangeTimeZoneInf)
        If Len(strExchTimeZone) > 0 And UCase(strExchTimeZone) <> "NY" And dExchTime > 0 Then
            dExchTime = ConvertTimeZone(dExchTime, "NY", strExchTimeZone)
        End If
        Set Bars = Nothing
        
        lLevel = CLng(Val(Parse(strDepth, vbTab, 3)))
        dGenesisBid = Val(Parse(strDepth, vbTab, 4)) / dMult
        dGenesisAsk = Val(Parse(strDepth, vbTab, 6)) / dMult
        
        ' If level 1 changed, send best bid/ask (Genesis Symbol<t>Exchange Time<t>Price<t>Size<t>)...
        If lLevel = 1& Then
            g.RealTime.StreamedMessage 25, strGenesisSymbol & vbTab & Str(dExchTime) & vbTab & Str(dGenesisBid) & vbTab & Parse(strDepth, vbTab, 5), m.nBroker
            g.RealTime.StreamedMessage 26, strGenesisSymbol & vbTab & Str(dExchTime) & vbTab & Str(dGenesisAsk) & vbTab & Parse(strDepth, vbTab, 7), m.nBroker
        End If
        
        ' Send depth of market message (Genesis Symbol<t>Level<t>Exchange Time<t>Bid<t>Bid Size<t>Ask<t>Ask Size<t>Bid Tick<t>Quote Condition<t>)...
        g.RealTime.StreamedMessage 52, strGenesisSymbol & vbTab & Str(lLevel) & vbTab & Str(dExchTime) & vbTab & Str(dGenesisBid) & vbTab & Parse(strDepth, vbTab, 5) & vbTab & Str(dGenesisAsk) & vbTab & Parse(strDepth, vbTab, 7) & vbTab & "" & vbTab & "Y", m.nBroker
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.HandleDepthChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MessageTypeToString
'' Description: Convert a message type enumeration to a string
'' Inputs:      Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MessageTypeToString(ByVal nType As eGDPhotonMessageTypes) As String
On Error GoTo ErrSection:

    Select Case nType
        Case eGDPhotonMessageType_Connect
            MessageTypeToString = "eGDPhotonMessageType_Connect"
        Case eGDPhotonMessageType_Disconnect
            MessageTypeToString = "eGDPhotonMessageType_Disconnect"
        Case eGDPhotonMessageType_SendOrder
            MessageTypeToString = "eGDPhotonMessageType_SendOrder"
        Case eGDPhotonMessageType_UnloadApp
            MessageTypeToString = "eGDPhotonMessageType_UnloadApp"
        Case eGDPhotonMessageType_GetOrders
            MessageTypeToString = "eGDPhotonMessageType_GetOrders"
        Case eGDPhotonMessageType_Subscribe
            MessageTypeToString = "eGDPhotonMessageType_Subscribe"
        Case eGDPhotonMessageType_Unsubscribe
            MessageTypeToString = "eGDPhotonMessageType_Unsubscribe"
        Case eGDPhotonMessageType_ConnectExchange
            MessageTypeToString = "eGDPhotonMessageType_ConnectExchange"
        Case eGDPhotonMessageType_DisconnectExchange
            MessageTypeToString = "eGDPhotonMessageType_DisconnectExchange"
        
        Case eGDPhotonMessageType_CNXSLE
            MessageTypeToString = "eGDPhotonMessageType_CNXSLE"
        Case eGDPhotonMessageType_ConnectedSLE
            MessageTypeToString = "eGDPhotonMessageType_ConnectedSLE"
        Case eGDPhotonMessageType_DisconnectedSLE
            MessageTypeToString = "eGDPhotonMessageType_DisconnectedSLE"
        Case eGDPhotonMessageType_OrderAnswer
            MessageTypeToString = "eGDPhotonMessageType_OrderAnswer"
        Case eGDPhotonMessageType_AppLoaded
            MessageTypeToString = "eGDPhotonMessageType_AppLoaded"
        Case eGDPhotonMessageType_AppUnloaded
            MessageTypeToString = "eGDPhotonMessageType_AppUnloaded"
        Case eGDPhotonMessageType_Heartbeat
            MessageTypeToString = "eGDPhotonMessageType_Heartbeat"
        Case eGDPhotonMessageType_CNXSLC
            MessageTypeToString = "eGDPhotonMessageType_CNXSLC"
        Case eGDPhotonMessageType_ConnectedSLC
            MessageTypeToString = "eGDPhotonMessageType_ConnectedSLC"
        Case eGDPhotonMessageType_DisconnectedSLC
            MessageTypeToString = "eGDPhotonMessageType_DisconnectedSLC"
        Case eGDPhotonMessageType_DepthChanged
            MessageTypeToString = "eGDPhotonMessageType_DepthChanged"
        Case eGDPhotonMessageType_PriceChanged
            MessageTypeToString = "eGDPhotonMessageType_PriceChanged"
        Case eGDPhotonMessageType_ConnectionInfo
            MessageTypeToString = "eGDPhotonMessageType_ConnectionInfo"
        Case eGDPhotonMessageType_Subscribed
            MessageTypeToString = "eGDPhotonMessageType_Subscribed"
        Case eGDPhotonMessageType_Unsubscribed
            MessageTypeToString = "eGDPhotonMessageType_Unsubscribed"
            
        Case Else
            MessageTypeToString = Str(nType)
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPhoton.MessageTypeToString"

End Function

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBrokerInfoOrder
'' Description: Add the order string to the broker info class
'' Inputs:      Order String, Refresh?
'' Returns:     None
''
'' RecordFrom:  NumRequest, Chaining, OrdClass, AnswerType, Symbol, Data,
''              B/S, PriceType, Val, AccountRef1, ClientOrderRef, ExchOrderRef,
''              PriceControl, AccountType1, AccountType2, MemberID, ClearingPos,
''              CTI, Fee, RoutingRef, Status, ServDate, Memo, Fex, Order Date,
''              Trigger Type, Secured Mod, Strategy, Trade Date, User Command,
''              Num Decimal, Session ID, Party Code, ExecRef, ExecDate, ExecType,
''              TraderID1, TraderID2, ExchOrderRef, AckType, AckDate, CommandAck,
''              RejectCode, RejectDate, CommandRej, ExchMsg, MessageCode,
''              MessageText, Bitmap, UserID, Index, Qty, Price, ValDate,
''              MinQty, DispQty, Stop, Remaining Qty, Num Trades, Avg Price,
''              Total Exec Qty, Old Qty, Old Price, Ton1, Exchange, Market,
''              Settle Date, Settle Num, Settle Price, Trade Val, Ton2,
''              Fix Price, Num Legs, Delta
'' RecordTo:    Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateBrokerInfoOrder(ByVal strOrderString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Order information split out from string
    Dim astrToAdd As cGdArray           ' Order information to add to broker info
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim dMult As Double                 ' Price Multiplier
    Dim dPrice As Double                ' Fill Price
    Dim lQuantity As Long               ' Fill Quantity
    
    If Len(strOrderString) > 0 Then
        Set astrOrder = New cGdArray
        astrOrder.Create eGDARRAY_Strings
        Set astrToAdd = New cGdArray
        astrToAdd.Create eGDARRAY_Strings
        
        astrOrder.SplitFields strOrderString, vbTab
        
        strGenesisSymbol = GenesisSymbol(astrOrder(5), astrOrder(0), dMult)
        If Len(strGenesisSymbol) > 0 Then
            astrToAdd(0) = astrOrder(12)
            astrToAdd(1) = astrOrder(11)
            astrToAdd(2) = ""
            astrToAdd(3) = astrOrder(10)
            astrToAdd(4) = strGenesisSymbol
            astrToAdd(5) = Str(GenesisOrderType(astrOrder(8)))
            If (astrOrder(7) = "0") Then astrToAdd(6) = Str(CLng(True)) Else astrToAdd(6) = Str(CLng(False))
            astrOrder(7) = astrOrder(53)
            astrOrder(8) = Str(Val(astrOrder(54)) / dMult)
            astrOrder(9) = Str(Val(astrOrder(58)) / dMult)
            Select Case astrOrder(9)
                Case "D"
                    astrOrder(10) = Str(JulFromLong(Val(astrOrder(55))))
                Case "J"
                    astrOrder(10) = "-1"
                Case Else
                    astrOrder(10) = "0"
            End Select
            astrOrder(11) = Str(GenesisOrderStatus(astrOrder(21)))
            astrOrder(12) = astrOrder(59)
            astrOrder(13) = Str(DateToDouble(astrOrder(25)))
        
            m.BrokerInfo.AddOrder astrToAdd.JoinFields(vbTab), bRefresh
        
            dPrice = Val(astrOrder(63))
            If dPrice = 0# Then dPrice = Val(astrOrder(54))
            lQuantity = CLng(Val(astrOrder(64)))
            If lQuantity = 0& Then lQuantity = CLng(Val(astrOrder(53)))
            
            If (Right(astrOrder(21), 1) = "X") And ((dPrice <> 0#) And (lQuantity <> 0&)) Then
                AddFillToBrokerInfo strOrderString, bRefresh
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.UpdateBrokerInfoOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderToBrokerInfo
'' Description: Add the Genesis order string to the broker info class
'' Inputs:      Order String, Refresh?
'' Returns:     None, Status
''
'' RecordTo:    Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GenesisOrderToBrokerInfo(ByVal Order As cPtOrder, ByVal nStatus As eTT_OrderStatus)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.GenesisOrderToBrokerInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelBrokerInfoOrder
'' Description: Cancel the order with the given Broker ID
'' Inputs:      Order
'' Returns:     None
''
'' RecordFrom:  Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date
'' RecordTo:    Exchange, Account, Genesis ID, Command, Symbol, Order Date,
''              Order Type, Limit Price, Stop Price, Quantity, Exchange ID,
''              B/S, Prev Quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelBrokerInfoOrder(ByVal strOrder As String)
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Order broken out into an array
    Dim astrToCancel As cGdArray        ' Order to send to the cancel message
    Dim strPhotonSymbol As String       ' Photon Symbol
    Dim strExchange As String           ' Photon Exchange
    Dim dMult As Double                 ' Multiplier for converting prices

    If Len(strOrder) > 0 Then
        Set astrOrder = New cGdArray
        astrOrder.Create eGDARRAY_Strings
        Set astrToCancel = New cGdArray
        astrToCancel.Create eGDARRAY_Strings
    
        astrOrder.SplitFields strOrder, vbTab
        strPhotonSymbol = PhotonSymbol(astrOrder(4), strExchange, dMult)
        If Len(strPhotonSymbol) > 0 Then
            astrToCancel(0) = strExchange
            astrToCancel(1) = astrOrder(3)
            astrToCancel(2) = astrOrder(1)
            astrToCancel(3) = Str(ePhotonCommand_Cancel)
            astrToCancel(4) = strPhotonSymbol
            astrToCancel(5) = Format(Val(astrOrder(13)), "YYYYMMDDHHMMSS")
            astrToCancel(6) = OrderType(CLng(Val(astrOrder(5))))
            astrToCancel(7) = Str(Val(astrOrder(8)) * dMult)
            astrToCancel(8) = Str(Val(astrOrder(9)) * dMult)
            astrToCancel(9) = astrOrder(7)
            astrToCancel(10) = ""
            If astrOrder(6) = "0" Then astrToCancel(11) = "0" Else astrToCancel(11) = "1"
            astrToCancel(12) = ""
            
            SendPhotonMessage eGDPhotonMessageType_SendOrder, astrToCancel.JoinFields(vbTab)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.CancelBrokerInfoOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFillToBrokerInfo
'' Description: Add the fill string to the broker info class
'' Inputs:      Fill String, Refresh?
'' Returns:     None
''
'' RecordFrom:  NumRequest, Chaining, OrdClass, AnswerType, Symbol, Data,
''              B/S, PriceType, Val, AccountRef1, ClientOrderRef, ExchOrderRef,
''              PriceControl, AccountType1, AccountType2, MemberID, ClearingPos,
''              CTI, Fee, RoutingRef, Status, ServDate, Memo, Fex, Order Date,
''              Trigger Type, Secured Mod, Strategy, Trade Date, User Command,
''              Num Decimal, Session ID, Party Code, ExecRef, ExecDate, ExecType,
''              TraderID1, TraderID2, ExchOrderRef, AckType, AckDate, CommandAck,
''              RejectCode, RejectDate, CommandRej, ExchMsg, MessageCode,
''              MessageText, Bitmap, UserID, Index, Qty, Price, ValDate,
''              MinQty, DispQty, Stop, Remaining Qty, Num Trades, Avg Price,
''              Total Exec Qty, Old Qty, Old Price, Ton1, Exchange, Market,
''              Settle Date, Settle Num, Settle Price, Trade Val, Ton2,
''              Fix Price, Num Legs, Delta
'' RecordTo:    Broker ID, Genesis ID, Fill ID, Account, Symbol, Fill Date,
''              B/S, Fill Quantity, Fill Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddFillToBrokerInfo(ByVal strFillString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim astrFill As cGdArray            ' Fill information split from the string
    Dim astrToAdd As cGdArray           ' Array to add to the broker information
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim dMult As Double                 ' Price Multiplier
    Dim dPrice As Double                ' Fill Price
    Dim lQuantity As Long               ' Fill Quantity
    
    If Len(strFillString) > 0 Then
        Set astrFill = New cGdArray
        astrFill.Create eGDARRAY_Strings
        Set astrToAdd = New cGdArray
        astrToAdd.Create eGDARRAY_Strings
        
        astrFill.SplitFields strFillString, vbTab
        
        strGenesisSymbol = GenesisSymbol(astrFill(5), astrFill(0), dMult)
        If Len(strGenesisSymbol) > 0 Then
            dPrice = Val(astrFill(63))
            If dPrice = 0# Then dPrice = Val(astrFill(54))
            lQuantity = CLng(Val(astrFill(64)))
            If lQuantity = 0& Then lQuantity = CLng(Val(astrFill(53)))
            
            astrToAdd(0) = astrFill(12)
            astrToAdd(1) = astrFill(11)
            astrToAdd(2) = astrFill(12) & " F" & astrFill(60)
            astrToAdd(3) = astrFill(10)
            astrToAdd(4) = strGenesisSymbol
            astrToAdd(5) = Str(DateToDouble(astrFill(36)))
            If (astrFill(7) = "0") Then astrToAdd(6) = Str(CLng(True)) Else astrToAdd(6) = Str(CLng(False))
            astrToAdd(7) = Str(lQuantity)
            astrToAdd(8) = Str(dPrice / dMult)
            
            m.BrokerInfo.AddFill astrToAdd.JoinFields(vbTab), bRefresh
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.AddFillToBrokerInfo"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddPositionToBrokerInfo
'' Description: Add the position string to the broker info class
'' Inputs:      Position String, Refresh?
'' Returns:     None
''
'' RecordFrom:  Error, Account, Contract, MaxLots, Overnight Pos, PNL, PNL (ticks),
''              Position, Position Limit, Previous Close, Previous Settle
'' RecordTo:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddPositionToBrokerInfo(ByVal strPositionString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.AddPositionToBrokerInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Do any initialization that needs to be done for the class
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return code from the registration
    Dim strDrive As String              ' Drive for the Photon OCX and DLL's

    m.bAppLoaded = False
    
    m.nBroker = eTT_AccountType_Photon
    
    If Not DirExist(AddSlash(App.Path) & "Photon") Then MkDir AddSlash(App.Path) & "Photon"
    KillFile AddSlash(App.Path) & "Photon\*.LOG /o=-14"
    
    Set m.BrokerInfo = New cBrokerInfo
    m.BrokerInfo.Init m.nBroker, AddSlash(App.Path) & "Photon"

    ' Attempt to register the necessary files...
    strDrive = Parse(App.Path, ":", 1) & ":"
    ChangePath AddSlash(CommonFilesPath) & "PhotonAPI"
    lReturn = gdRegisterFiles(AddSlash(CommonFilesPath) & "PhotonAPI\ActivePhoton.OCX", ByVal 0)
    ChangePath App.Path
    
    m.strIniFile = AddSlash(App.Path) & "Photon.INI"

    Set m.ModifiedOrders = New cGdTree
    Set m.astrErrorCodes = New cGdArray
    
    Set m.astrBaseSyms = New cGdArray
    m.astrBaseSyms.Create eGDARRAY_Strings
    m.astrBaseSyms.FromFile AddSlash(App.Path) & "Provided\PhoToGen.TXT"
    m.astrBaseSyms.Sort
    
    m.astrErrorCodes.FromFile AddSlash(App.Path) & "Provided\PhotonError.TXT"
    m.astrErrorCodes.Sort
    
    Set m.astrSubContract = New cGdArray
    m.astrSubContract.Create eGDARRAY_Strings
    
    GetConnectionProperties
    
    m.bVerifyPositions = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.Class.Initialize", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Do any clean up that needs to be done for the class
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.BrokerInfo = Nothing
    Set m.ModifiedOrders = Nothing
    Set m.astrBaseSyms = Nothing
    Set m.astrErrorCodes = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPhoton.Class.Terminate", eGDRaiseError_Raise

End Sub

