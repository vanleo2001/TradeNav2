VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFXCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cFXCM.cls
'' Description: Wrapper class for interacting with the FXCM server
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 06/09/2009   DAJ         Set the BrokerInfo connection status
'' 09/01/2009   DAJ         Use new Parked order status
'' 03/11/2010   DAJ         Use global collections
'' 03/07/2011   DAJ         Broker Disconnect Alerts
'' 06/21/2011   DAJ         Last connection status change
'' 01/31/2012   DAJ         Handle unsolicted cancels in TradeSense order groups
'' 03/21/2012   DAJ         Fix for amending a partially filled order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strIniFile As String                ' Ini file for the FXCM information
    nBroker As eTT_AccountType          ' Broker type for the FXCM broker
    
    strUserName As String               ' User name to logon to the server with
    strPassword As String               ' Password to logon to the server with
    strUrl As String                    ' Host IP for the FXCM server
    strConnection As String             ' Type of account
    
    lLotSize As Long                    ' Lot size for the account
    dPipCost As Double                  ' Pip cost in dollars????
    
    bHedging As Boolean                 ' Account has hedging turned on
        
    nStatus As eGDConnectionStatus      ' Connection status to the FXCM servers
    bWasConnected As Boolean            ' This broker was has been connected
    dHeartBeat As Double                ' Tick count as of last message receieved
    bReconnecting As Boolean            ' Are we currently reconnecting?
    lReconnectAttempts As Long          ' Reconnect attempts that have been made
    bAppLoaded As Boolean               ' Is the FXCM application loaded?
    bManualDisconnect As Boolean        ' Is the user manually disconnecting?
    bAskedPassword As Boolean           ' Have we asked for the password yet?
    dLastFillTime As Double             ' Tick count of when we received the last fill
    strLastMessage As String            ' Last message from the connection status message
    bVerifyPositions As Boolean         ' Do we want to verify the positions with the user?
    
    dRefreshingStart As Double          ' Start of the refreshing time
    bRefreshingAccounts As Boolean      ' Are we refreshing accounts?
    bRefreshingOrders As Boolean        ' Are we refreshing orders?
    bRefreshingFills As Boolean         ' Are we refreshing fills?
    bRefreshingPositions As Boolean     ' Are we refreshing positions?
    
    astrAccounts As cGdArray            ' List of accounts currently connected to
    astrSymbols As cGdArray             ' Array of symbol conversion information
    astrPositions As cGdArray           ' Array for Open Positions with PositionID's
    astrConsolidated As cGdArray        ' Array for Consolidated Positions
    
    BrokerInfo As cBrokerInfo           ' Object of broker information
End Type
Private m As mPrivate

Public Property Get UserName() As String
    UserName = m.strUserName
End Property
Public Property Let UserName(ByVal strUserName As String)
    m.strUserName = strUserName
End Property

Public Property Get Password() As String
    Password = m.strPassword
End Property
Public Property Let Password(ByVal strPassword As String)
    m.strPassword = strPassword
    m.bAskedPassword = True
End Property

Public Property Get Url() As String
    Url = m.strUrl
End Property
Public Property Let Url(ByVal strUrl As String)
    m.strUrl = strUrl
End Property

Public Property Get Connection() As String
    Connection = m.strConnection
End Property
Public Property Let Connection(ByVal strConnection As String)
    m.strConnection = strConnection
End Property

Public Property Get ConnectionStatus() As eGDConnectionStatus
    ConnectionStatus = m.nStatus
End Property
Private Property Let ConnectionStatus(ByVal nNewStatus As eGDConnectionStatus)
    If nNewStatus <> m.nStatus Then
        m.nStatus = nNewStatus
        
        If Not g.Broker Is Nothing Then
            g.Broker.LastStatusChange = Now
        End If
        If nNewStatus = eGDConnectionStatus_Connected Then
            g.Broker.PositionsToVerify = True
        End If
    End If
End Property

Public Property Get AppLoaded() As Boolean
    AppLoaded = m.bAppLoaded
End Property

Public Property Get BrokerInfo() As cBrokerInfo
    Set BrokerInfo = m.BrokerInfo
End Property

Public Property Get Refreshing() As Boolean
    Refreshing = m.bRefreshingAccounts Or m.bRefreshingPositions Or m.bRefreshingOrders Or m.bRefreshingFills
End Property

Public Property Get PositionVerify() As Boolean
    PositionVerify = m.bVerifyPositions
End Property
Public Property Let PositionVerify(ByVal bVerifyPositions As Boolean)
    m.bVerifyPositions = bVerifyPositions
End Property
Public Property Get PipCost() As Double
    PipCost = m.dPipCost
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

    If (Len(m.strUserName) = 0) Or (Len(m.strUrl) = 0) Or (Len(m.strConnection) = 0) Then
        ShowConfig = True
    Else
        ShowConfig = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.ShowConfig"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Connect
'' Description: Attempt a connection to the FXCM servers if not already connected
'' Inputs:      New Account, Are we switching?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Connect(Optional ByVal strNewAccount As String = "", Optional ByVal bSwitching As Boolean = False)
On Error GoTo ErrSection:

    Dim bConnect As Boolean             ' Should we still try to connect?
    Dim strProcess As String            ' Name of the process to start
    Dim strPath As String               ' Path to FXOrder2Go API files
    Dim lTimeOut As Long                ' Timeout variable
    Dim bUpdateAPI As Boolean           ' Should we check for API updates?
    
    bUpdateAPI = Not m.bAskedPassword
        
    If g.bUnloading = False Then
        If (ConnectionStatus = eGDConnectionStatus_Disconnected) Then
            If (m.bAskedPassword = True) And ((strNewAccount = m.strUserName) Or (Len(strNewAccount) = 0)) Then
                bConnect = True
            Else
                bConnect = frmFxcmLogin.ShowMe(strNewAccount)
                m.bAskedPassword = bConnect
            End If
        Else
            If ((strNewAccount <> m.strUserName) And (Len(strNewAccount) > 0)) Or (bSwitching = True) Then
                bConnect = frmFxcmLogin.ShowMe(strNewAccount, False, bSwitching)
                If bConnect = True Then
                    m.bAskedPassword = True
                    Disconnect False, "Changing accounts to " & strNewAccount
                    
                    lTimeOut = 0&
                    Do While (ConnectionStatus <> eGDConnectionStatus_Disconnected) And (lTimeOut < 30&)
                        Sleep 1
                        lTimeOut = lTimeOut + 1&
                    Loop
                    
                    bConnect = (ConnectionStatus = eGDConnectionStatus_Disconnected)
                End If
            End If
        End If
        
        If (bConnect = True) And (g.bUnloading = False) Then
            If LiveTradingAllowed(m.nBroker) Then
                ConnectionStatus = eGDConnectionStatus_Connecting
                
                If ((frmOnlineBroker.gdBroker.FindControlID("FXCM") = 0) Or (KillProcess("GenFXCM", True) = 0)) Then
                    If KillProcess("GenFXCM") > 0 Then
                        Sleep 3
                    End If
                    
                    'Check for FXCM Api files installed
                    strPath = GetRegistryValue(rkLocalMachine, "Software\CandleWorks\FxOrder2Go", "Path", "")
                    If strPath = "" Or (Not FileExist(strPath & "RunOrder2GoRE.bat")) Then
                        ' If not installed, launch browser with URL for install
                        DumpDebug "FXCM API not installed, running web install"
                        RunProcess InternetBrowser, "http://www.fxcorporate.com/Order2Go/FXOrder2GoRE.EXE", True
                        DumpDebug "FXCM API Install Complete"
                        bUpdateAPI = True
                    End If
                        
                    ' launch batch file for FXCM API updater
                    If bUpdateAPI Then
                        RunProcess strPath & "RunOrder2GoRE.bat", , True, vbHide
                    End If
                    
                    DumpDebug "Trade Navigator version: " & FormatVersion & "." & Str(App.Revision) & " " & DateFormat(FileDate(App.Path & "\" & App.EXEName & ".EXE"), MM_DD_YYYY, HH_MM, AMPM_UPPER)
                    strProcess = AddSlash(App.Path) & "..\Brokers\FXCM\GenFXCM.EXE"
                    DumpDebug "Loading application " & strProcess
                    RunProcess strProcess, , , vbHide ' , vbMinimizedNoFocus
                Else
                    DumpDebug "Connecting to FXCM as " & m.strUserName & " on " & m.strUrl & " (" & m.strConnection & ")"
                    SendFxcmMessage eGDFxcmMessageType_Connect, m.strUserName & vbTab & m.strPassword & vbTab & m.strUrl & vbTab & m.strConnection
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.Connect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disconnect
'' Description: Disconnect form the FXCM servers if connected
'' Inputs:      Force Reconnect?, Reason for Disconnect
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Disconnect(Optional ByVal bForceReconnect As Boolean = False, Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    If (ConnectionStatus = eGDConnectionStatus_Connected) Then
        m.bManualDisconnect = Not bForceReconnect
        If m.bManualDisconnect Then m.dHeartBeat = 0#
        If Len(strReason) = 0 Then DumpDebug "Disconnecting from FXCM" Else DumpDebug "Disconnecting from FXCM: " & strReason
        SendFxcmMessage eGDFxcmMessageType_Disconnect, ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.Disconnect"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reconnect
'' Description: Attempt to reconnect to the FXCM servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Reconnect()
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Time out variable

    If m.bReconnecting = False Then
        m.bReconnecting = True
        
        If ConnectionStatus = eGDConnectionStatus_Connected Then
            ConnectionStatus = eGDConnectionStatus_Disconnecting
            Disconnect True, "Reconnecting"
        End If
        
        lTimeOut = 0&
        Do While (m.bAppLoaded = True) And (lTimeOut < 30&)
            Sleep 1#
            lTimeOut = lTimeOut + 1&
        Loop
                        
        If (m.bAppLoaded = False) Or (lTimeOut >= 30&) Then
            m.lReconnectAttempts = m.lReconnectAttempts + 1&
            ConnectionStatus = eGDConnectionStatus_Disconnected
            Connect
        End If
        
        lTimeOut = 0&
        Do While (ConnectionStatus <> eGDConnectionStatus_Connected) And (lTimeOut < 30&)
            Sleep 1#
            lTimeOut = lTimeOut + 1&
        Loop
            
        m.bReconnecting = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.Reconnect"
    
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
                            
    Order.StatusDate = ConvertToBrokerDate(Now, m.nBroker, Order.Symbol)
    Order.Save
    
    If Order.Status <> eTT_OrderStatus_Error Then
        m.BrokerInfo.AddOrder Order, False
    End If
    
    RefreshOrder Order
    g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.SendOrder"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOrder
'' Description: Attempt to add the order
'' Inputs:      Order to add
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AddOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information
    Dim strFxcmSymbol As String         ' Fxcm Symbol
    Dim strFxcmOrderType As String      ' Fxcm Order Type
    Dim dMult As Double                 ' Multiplier for converting prices
    Dim lQuantity As Integer
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    strFxcmSymbol = FxcmSymbol(Order.Symbol, dMult)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strFxcmOrderType = FxcmOrderType(Order.OrderType)
    lQuantity = Order.Quantity
    
    If Len(strFxcmSymbol) = 0 Then
        DumpDebug "AddOrder: " & Order.Symbol & " cannot be converted to a FXCM symbol"
        InfBox Order.Symbol & " cannot be traded through the FXCM servers", "!", , "FXCM Order Error"
    ElseIf Len(strFxcmOrderType) = 0 Then
        DumpDebug "AddOrder: " & OrderType(Order.OrderType) & " cannot be converted to a FXCM order type"
        InfBox OrderType(Order.OrderType) & " is not a valid order type for FXCM", "!", , "FXCM Order Error"
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        DumpDebug "AddOrder: Not currently connected to FXCM account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to FXCM account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , "FXCM Order Error"
        Order.Message = "Not currently connected to FXCM account " & g.Broker.AccountNameForID(Order.AccountID)
     ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "AddOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the FXCM servers", "!", , "FXCM Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "AddOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the FXCM servers", "!", , "FXCM Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
    Else
        ' add a new order
        If lQuantity <> 0 Then
            ' Send new entry order to Broker
            DumpDebug "AddOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
            astrOrder(0) = Order.GenesisOrderID
            astrOrder(1) = ""
            astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
            astrOrder(3) = strFxcmSymbol
            astrOrder(4) = ""
            astrOrder(5) = strFxcmOrderType
            If Order.Buy Then astrOrder(6) = "B" Else astrOrder(6) = "S"
            astrOrder(7) = Str(Order.Quantity * m.lLotSize)
            
            If Order.OrderType = eTT_OrderType_Limit Then
                astrOrder(8) = FxcmPrice(Order.LimitPrice, dMult)
            Else
                astrOrder(8) = ""
            End If
            
            If Order.OrderType = eTT_OrderType_Stop Then
                astrOrder(9) = FxcmPrice(Order.StopPrice, dMult)
            Else
                astrOrder(9) = ""
            End If

            astrOrder(10) = ""
            astrOrder(11) = ""
            astrOrder(12) = ""
            astrOrder(13) = ""
            astrOrder(14) = ""
            astrOrder(15) = ""
            astrOrder(16) = ""
            
            'GenesisOrderToBrokerInfo Order, eTT_OrderStatus_Sent
            SendFxcmMessage eGDFxcmMessageType_AddOrder, astrOrder.JoinFields(vbTab)
        
            AddOrder = True
        End If
    End If

    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.AddOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AmendOrder
'' Description: Attempt to Amend the order
'' Inputs:      Order to Amend
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AmendOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information
    Dim strFxcmSymbol As String         ' Fxcm Symbol
    Dim strFxcmOrderType As String      ' Fxcm Order Type
    Dim dMult As Double                 ' Multiplier for converting prices
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    strFxcmSymbol = FxcmSymbol(Order.Symbol, dMult)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strFxcmOrderType = FxcmOrderType(Order.OrderType)
    
    If Len(strFxcmSymbol) = 0 Then
        DumpDebug "AmendOrder: " & Order.Symbol & " cannot be converted to a FXCM symbol"
        InfBox Order.Symbol & " cannot be traded through the FXCM servers", "!", , "FXCM Order Error"
    ElseIf Len(strFxcmOrderType) = 0 Then
        DumpDebug "AmendOrder: " & OrderType(Order.OrderType) & " cannot be converted to a FXCM order type"
        InfBox OrderType(Order.OrderType) & " is not a valid order type for FXCM", "!", , "FXCM Order Error"
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        DumpDebug "AmendOrder: Not currently connected to FXCM account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to FXCM account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , "FXCM Order Error"
        Order.Message = "Not currently connected to FXCM account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "AmendOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the FXCM servers", "!", , "FXCM Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "AmendOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the FXCM servers", "!", , "FXCM Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
    Else
        DumpDebug "AmendOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"

        ' Need to check if ammend quantity do Cancel/Replace???
        
        ' if it's an entry order send ammend message
        astrOrder(0) = Order.GenesisOrderID
        astrOrder(1) = Order.BrokerID
        astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(3) = strFxcmSymbol
        astrOrder(4) = ""
        astrOrder(5) = strFxcmOrderType
        If Order.Buy Then astrOrder(6) = "B" Else astrOrder(6) = "S"
        astrOrder(7) = Str(Order.Quantity * m.lLotSize)
        
        If Order.OrderType = eTT_OrderType_Limit Then
            astrOrder(8) = FxcmPrice(Order.LimitPrice, dMult)
        Else
            astrOrder(8) = ""
        End If
            
        If Order.OrderType = eTT_OrderType_Stop Then
            astrOrder(9) = FxcmPrice(Order.StopPrice, dMult)
        Else
            astrOrder(9) = ""
        End If
        
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = ""
        astrOrder(13) = ""
        astrOrder(14) = ""
        astrOrder(15) = ""
        astrOrder(16) = ""
        
        SendFxcmMessage eGDFxcmMessageType_AmendOrder, astrOrder.JoinFields(vbTab)
        
        AmendOrder = True
    End If

    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.AmendOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Attempt to Cancel the order
'' Inputs:      Order to Cancel
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information
'    Dim astrPosition As New cGdArray    ' Array to hold Position info
    Dim strFxcmSymbol As String         ' Fxcm Symbol
    Dim strFxcmOrderType As String      ' Fxcm Order Type
    Dim dMult As Double                 ' Multiplier for converting prices
'    Dim lIndex As Integer               ' Index into array
'    Dim bFound As Boolean               ' Order found flag
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
'    bFound = False
    strFxcmSymbol = FxcmSymbol(Order.Symbol, dMult)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strFxcmOrderType = FxcmOrderType(Order.OrderType)
    
    If Len(strFxcmSymbol) = 0 Then
        DumpDebug "CancelOrder: " & Order.Symbol & " cannot be converted to a FXCM symbol"
        InfBox Order.Symbol & " cannot be traded through the FXCM servers", "!", , "FXCM Order Error"
    ElseIf Len(strFxcmOrderType) = 0 Then
        DumpDebug "CancelOrder: " & OrderType(Order.OrderType) & " cannot be converted to a FXCM order type"
        InfBox OrderType(Order.OrderType) & " is not a valid order type for FXCM", "!", , "FXCM Order Error"
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
        DumpDebug "CancelOrder: Not currently connected to FXCM account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to FXCM account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , "FXCM Order Error"
        Order.Message = "Not currently connected to FXCM account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "CancelOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the FXCM servers", "!", , "FXCM Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "CancelOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the FXCM servers", "!", , "FXCM Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
    Else
        If Not g.TsoGroups Is Nothing Then
            g.TsoGroups.CancelRequested Order
        End If
        
        DumpDebug "CancelOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
        
        'If it's an entry order, send cancel
        astrOrder(0) = Order.GenesisOrderID
        astrOrder(1) = Order.BrokerID
        astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(3) = strFxcmSymbol
        astrOrder(4) = ""
        astrOrder(5) = strFxcmOrderType
        If Order.Buy Then astrOrder(6) = "B" Else astrOrder(6) = "S"
        astrOrder(7) = Str(Order.Quantity * m.lLotSize)
        If Order.OrderType = eTT_OrderType_Limit Then
            astrOrder(8) = FxcmPrice(Order.LimitPrice, dMult)
        Else
            astrOrder(8) = ""
        End If
        astrOrder(9) = ""
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = ""
        astrOrder(13) = ""
        astrOrder(14) = ""
        astrOrder(15) = ""
        astrOrder(16) = ""
        
        SendFxcmMessage eGDFxcmMessageType_CancelOrder, astrOrder.JoinFields(vbTab)
                
        Order.Status = eTT_OrderStatus_CancelPending
        Order.StatusDate = ConvertToBrokerDate(Now, m.nBroker, Order.Symbol)
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
    RaiseError "cFXCM.CancelOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Attempt to Park a working order (Cancel order at broker, keep order in UI)
'' Inputs:      Order to Cancel
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ParkOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information
'    Dim astrPosition As New cGdArray    ' Array to hold Position info
    Dim strFxcmSymbol As String         ' Fxcm Symbol
    Dim strFxcmOrderType As String      ' Fxcm Order Type
    Dim dMult As Double                 ' Multiplier for converting prices
'    Dim lIndex As Integer               ' Index into array
'    Dim bFound As Boolean               ' Order found flag
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
'    bFound = False
    strFxcmSymbol = FxcmSymbol(Order.Symbol, dMult)
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strFxcmOrderType = FxcmOrderType(Order.OrderType)
    
    If Len(strFxcmSymbol) = 0 Then
        DumpDebug "ParkOrder: " & Order.Symbol & " cannot be converted to a FXCM symbol"
        InfBox Order.Symbol & " cannot be traded through the FXCM servers", "!", , "FXCM Order Error"
    ElseIf Len(strFxcmOrderType) = 0 Then
        DumpDebug "ParkOrder: " & OrderType(Order.OrderType) & " cannot be converted to a FXCM order type"
        InfBox OrderType(Order.OrderType) & " is not a valid order type for FXCM", "!", , "FXCM Order Error"
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
        DumpDebug "ParkOrder: Not currently connected to FXCM account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to FXCM account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , "FXCM Order Error"
        Order.Message = "Not currently connected to FXCM account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "ParkOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the FXCM servers", "!", , "FXCM Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "ParkOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the FXCM servers", "!", , "FXCM Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the FXCM servers"
    Else
        DumpDebug "ParkOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
                
        'If it's an entry order, send cancel
        astrOrder(0) = Order.GenesisOrderID
        astrOrder(1) = Order.BrokerID
        astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(3) = strFxcmSymbol
        astrOrder(4) = ""
        astrOrder(5) = strFxcmOrderType
        If Order.Buy Then astrOrder(6) = "B" Else astrOrder(6) = "S"
        astrOrder(7) = Str(Order.Quantity * m.lLotSize)
        If Order.OrderType = eTT_OrderType_Limit Then
            astrOrder(8) = FxcmPrice(Order.LimitPrice, dMult)
        Else
            astrOrder(8) = ""
        End If
        astrOrder(9) = ""
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = ""
        astrOrder(13) = ""
        astrOrder(14) = ""
        astrOrder(15) = ""
        astrOrder(16) = ""
        
        SendFxcmMessage eGDFxcmMessageType_CancelOrder, astrOrder.JoinFields(vbTab)
                
        Order.Status = eTT_OrderStatus_ParkPending
        Order.StatusDate = ConvertToBrokerDate(Now, m.nBroker, Order.Symbol)
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
    RaiseError "cFXCM.ParkOrder"
    
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
    RaiseError "cFXCM.CancelOrderFromBrokerID"
    
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
    RaiseError "cFXCM.CancelWorkingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenPosition
'' Description: Flattent the position for the given account and symbol
'' Inputs:      Account, Symbol, Auto Trade Item ID, Confirm?
'' Returns:     Did user Confirm?
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FlattenPosition(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, Optional ByVal bConfirm As Boolean = True) As Boolean
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
            ExitPositionForSymbol strAccount, strSymbol, lAtID, lPosition
        End If
    End If
    
    FlattenPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.FlattenPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Ask the server for the available accounts for this user
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAccounts()
On Error GoTo ErrSection:

    If Not Refreshing Then
        DumpDebug "Getting Accounts"
        m.bRefreshingAccounts = True
        
        SendFxcmMessage eGDFxcmMessageType_GetAccounts, ""
    Else
        DumpDebug "Account refresh called for but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.GetAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Ask the server for the available Orders for this user
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetOrders(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long

    If Not Refreshing Then
        m.bRefreshingOrders = True
        If Len(strAccount) = 0 Then
            DumpDebug "Getting Orders"
            For lIndex = 0 To m.astrAccounts.Size - 1
                SendFxcmMessage eGDFxcmMessageType_GetOrders, m.astrAccounts(lIndex)
            Next lIndex
        Else
            DumpDebug "Getting Orders for " & strAccount
            SendFxcmMessage eGDFxcmMessageType_GetOrders, strAccount
        End If
    Else
        DumpDebug "Order refresh called for but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.GetOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFills
'' Description: Ask the server for the available Fills for this user
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetFills(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long

    If Not Refreshing Then
        m.bRefreshingFills = True
        If Len(strAccount) = 0 Then
            DumpDebug "Getting Fills"
            For lIndex = 0 To m.astrAccounts.Size - 1
                SendFxcmMessage eGDFxcmMessageType_GetFills, m.astrAccounts(lIndex)
            Next lIndex
        Else
            DumpDebug "Getting Fills for " & strAccount
            SendFxcmMessage eGDFxcmMessageType_GetFills, strAccount
        End If
    Else
        DumpDebug "Fill refresh called for but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.GetFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositions
'' Description: Ask the server for the available Positions for this user
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPositions(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long
    
    If Not Refreshing Then
        m.bRefreshingPositions = True
        m.astrPositions.Clear
        
        If Len(strAccount) = 0 Then
            DumpDebug "Getting Positions"
            For lIndex = 0 To m.astrAccounts.Size - 1
                SendFxcmMessage eGDFxcmMessageType_GetPositions, m.astrAccounts(lIndex)
            Next lIndex
        Else
            DumpDebug "Getting Positions for " & strAccount
            SendFxcmMessage eGDFxcmMessageType_GetPositions, strAccount
        End If
        
    Else
        DumpDebug "Position refresh called for but a refresh is already in progress"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Refresh
'' Description: Ask the server for accounts, orders, fills, and positions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Refresh()
On Error GoTo ErrSection:

    If (ConnectionStatus = eGDConnectionStatus_Connected) Then
        If Not Refreshing Then
            DumpDebug "Refreshing"
            GetAccounts
        Else
            DumpDebug "Refresh called for but a refresh is already in progress"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.Refresh"
    
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

    Dim nReturn As eGDConnectionStatus  ' Return value from the function

    Select Case ConnectionStatus
        Case eGDConnectionStatus_Connecting
            nReturn = eGDConnectionStatus_Connecting
            
        Case eGDConnectionStatus_Disconnecting
            nReturn = eGDConnectionStatus_Disconnecting
            
        Case eGDConnectionStatus_Disconnected
            nReturn = eGDConnectionStatus_Disconnected
            
        Case eGDConnectionStatus_Connected
            If m.astrAccounts.BinarySearch(strAccountNumber) Then
                nReturn = eGDConnectionStatus_Connected
            Else
                nReturn = eGDConnectionStatus_Disconnected
            End If
            
    End Select
    
    ConnectionStatusForAccount = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.ConnectionStatusForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextGenesisID
'' Description: Determine the next unique Genesis ID for the given account
'' Inputs:      Account Number
'' Returns:     Next Unique Genesis ID
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
    RaiseError "cFXCM.NextGenesisID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckHeartbeat
'' Description: Check the heartbeat to make sure that we are still connected
''              to the FXCM servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckHeartbeat()
On Error GoTo ErrSection:

    If (m.dHeartBeat <> 0) And (g.bUnloading = False) Then
        If (m.dHeartBeat < (gdTickCount - 120000)) Then
            If m.lReconnectAttempts = 1& Then
                InfBox "Communication has been lost with FXCM.  Attempting to reconnect...", , "+-OK", "FXCM Connection", True
            End If
            Reconnect
        ElseIf (ConnectionStatus = eGDConnectionStatus_Connected) Then
            If m.lReconnectAttempts > 1& Then
                InfBox "TradeNavigator has reconnected to FXCM.", , "+-OK", "FXCM Connection", True
            End If
            m.lReconnectAttempts = 0&
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.CheckHeartbeat"
    
End Sub

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

    Dim strDate As String               ' Date portion of the string
    Dim strTime As String               ' Time portion of the string
    Dim strAmPm As String               ' PM portion of the date string
    Dim lHour As Long                   ' Hour from the time
    Dim lMinute As Long                 ' Minute from the time
    Dim dSecond As Double               ' Second from the time
    Dim dTime As Double                 ' Fractional part of the day
    Dim dGMTOffset As Double            ' Gmt offset if time in NY
    Dim dReturn As Double               ' Return value of the function
    Dim lYear As Long                   ' Year
    Dim lMonth As Long                  ' Month
    Dim lDay As Long                    ' Day
    
    dReturn = 0#
    dGMTOffset = 0#
    
    If Len(strDateTime) > 0 Then
        If InStr(1, strDateTime, "-") Then
            strDate = Parse(strDateTime, "-", 1)
            strTime = Parse(strDateTime, "-", 2)
            strAmPm = ""
            lMonth = CLng(Val(Mid$(strDate, 5, 2)))
            lDay = CLng(Val(Right$(strDate, 2)))
            lYear = CLng(Val(Left$(strDate, 4)))
            dGMTOffset = 0 - (4# / 24#)
        Else
            strDate = Parse(strDateTime, " ", 1)
            strTime = Parse(strDateTime, " ", 2)
            strAmPm = Parse(strDateTime, " ", 3)
            lMonth = CLng(Val(Parse(strDate, "/", 1)))
            lDay = CLng(Val(Parse(strDate, "/", 2)))
            lYear = CLng(Val(Parse(strDate, "/", 3)))
        End If
        
        lHour = CLng(Val(Parse(strTime, ":", 1)))
        lMinute = CLng(Val(Parse(strTime, ":", 2)))
        dSecond = Val(Parse(strTime, ":", 3))
        
        If InStr(1, strAmPm, "PM") And lHour < 12 Then lHour = lHour + 12
        
        dTime = (lHour / 24#) + (lMinute / 1440#) + (dSecond / 86400#)
        dReturn = dGMTOffset + JulFromLong(CLng((lYear * 10000) + (lMonth * 100) + lDay)) + dTime
    End If
    
    DateToDouble = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.DateToDouble"
    
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
    RaiseError "cFXCM.DumpDebug"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendFxcmMessage
'' Description: Send an App Mail message to the FXCM program
'' Inputs:      Type of Message, Message, Send Now?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendFxcmMessage(ByVal nType As eGDFxcmMessageTypes, ByVal strMessage As String, Optional ByVal bSendNow As Boolean = True)
On Error GoTo ErrSection:

    Dim astrMessage As New cGdArray     ' Message split out into an array

    If nType = eGDFxcmMessageType_Connect Then
        astrMessage.SplitFields strMessage, vbTab
        astrMessage(1) = SetStrToChar(astrMessage(1), "*")
        DumpDebug vbTab & "Sending FXCM Message (" & MessageTypeToString(nType) & "): " & astrMessage.JoinFields(vbTab)
    Else
        DumpDebug vbTab & "Sending FXCM Message (" & MessageTypeToString(nType) & "): " & strMessage
    End If
    
    frmOnlineBroker.gdBroker.CreateMessage "FXCM", nType, strMessage, , bSendNow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.SendFxcmMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming App Mail message from the FXCM program
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDFxcmMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    If (nType <> eGDFxcmMessageType_ConnectionInfo) And (nType <> eGDFxcmMessageType_Heartbeat) And (nType <> eGDFxcmMessageType_QuoteR) Then
        DumpDebug vbTab & "FXCM Message Received (" & MessageTypeToString(nType) & "): " & strMessage
    End If
    
    Select Case nType
        Case eGDFxcmMessageType_ConnectionInfo
            HandleConnectionInfo strMessage
        
        Case eGDFxcmMessageType_AppLoaded
            m.bAppLoaded = True
            If LiveTradingAllowed(m.nBroker) Then
                DumpDebug "Connecting to FXCM as " & m.strUserName & " on " & m.strUrl & " (" & m.strConnection & ")"
                SendFxcmMessage eGDFxcmMessageType_Connect, m.strUserName & vbTab & m.strPassword & vbTab & m.strUrl & vbTab & m.strConnection
            Else
                DumpDebug "Unloading the application since live trading is not allowed"
                SendFxcmMessage eGDFxcmMessageType_UnloadApp, ""
            End If
        
        Case eGDFxcmMessageType_AppUnloaded
            m.bAppLoaded = False
        
        Case eGDFxcmMessageType_Heartbeat
            m.dHeartBeat = gdTickCount
            If m.dLastFillTime > 0 Then
                If gdTickCount > m.dLastFillTime + 5000 Then
                    GetPositions
                    m.dLastFillTime = 0#
                End If
            End If
        
        Case eGDFxcmMessageType_AccountR
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingAccounts = True
                m.BrokerInfo.ClearAccounts
                m.astrAccounts.Clear
            ElseIf UCase(strMessage) = "END" Then
                m.BrokerInfo.RefreshingAccounts = False
                m.bRefreshingAccounts = False
                GetOrders
            Else
                HandleAccountInfo strMessage
                m.BrokerInfo.AddAccount strMessage
                m.BrokerInfo.SetConnectionStatusForAccount eGDConnectionStatus_Connected, Parse(strMessage, vbTab, 1)
            End If
        
        Case eGDFxcmMessageType_FillR
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingFills = True
'                m.BrokerInfo.ClearFills
            ElseIf UCase(strMessage) = "END" Then
                m.BrokerInfo.RefreshingFills = False
                m.bRefreshingFills = False
                GetPositions
            Else
                'AddFillToBrokerInfo strMessage, True
                HandleOrderInfo strMessage, True
            End If
        
        Case eGDFxcmMessageType_Order
            m.dHeartBeat = gdTickCount
            'UpdateBrokerInfoOrder strMessage, False
            HandleOrderInfo strMessage, False
            GetPositions
        
        Case eGDFxcmMessageType_OrderR
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingOrders = True
                m.BrokerInfo.ClearOrders
            ElseIf UCase(strMessage) = "END" Then
                m.BrokerInfo.RefreshingOrders = False
                m.bRefreshingOrders = False
                GetFills
            Else
                'UpdateBrokerInfoOrder strMessage, True
                HandleOrderInfo strMessage, True
            End If
        
        Case eGDFxcmMessageType_PositionR
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingPositions = True
                BeginPosition
            ElseIf UCase(strMessage) = "END" Then
                m.BrokerInfo.RefreshingPositions = False
                m.bRefreshingPositions = False
                EndPosition
            Else
                AddPosition strMessage
'                AddPositionToBrokerInfo strMessage, True
                HandlePositionInfo strMessage
            End If
    
    End Select
    
    If Refreshing And (m.dRefreshingStart = 0#) Then
        m.dRefreshingStart = gdTickCount
    ElseIf Refreshing And (gdTickCount > m.dRefreshingStart + 30000#) Then
        m.bRefreshingAccounts = False
        m.bRefreshingOrders = False
        m.bRefreshingFills = False
        m.bRefreshingPositions = False
    ElseIf (Not Refreshing) And (m.dRefreshingStart <> 0#) Then
        m.dRefreshingStart = 0#
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.HandleMessage"

End Sub

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

    If (ConnectionStatus = eGDConnectionStatus_Connected) And (m.astrAccounts.Size > 0) Then
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
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.VerifyPositions"
    
End Sub

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

    IsTradeableSymbol = (Len(FxcmSymbol(strGenesisSymbol)) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.IsTradeableSymbol"
    
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
    Dim lIndex As Long                  ' Index into a for loop
    Dim strMask As String               ' Order type mask from the file
    
    bReturn = False
    
    If nOrderType <> eTT_OrderType_Adjustment Then
        For lIndex = 0 To m.astrSymbols.Size - 1
            If Parse(m.astrSymbols(lIndex), vbTab, 2) = strGenesisSymbol Then
                strMask = Parse(m.astrSymbols(lIndex), vbTab, 4)
                If Len(strMask) = 0 Then strMask = "11110000"
                If Len(strMask) = 8 Then strMask = strMask & "0"
                If (nOrderType >= 0) And (nOrderType < Len(strMask)) Then
                    bReturn = (Mid(strMask, nOrderType + 1, 1) = "1")
                End If
                
                Exit For
            End If
        Next lIndex
    End If
    
    OrderTypeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.OrderTypeAllowed"
    
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
    Dim lIndex As Long                  ' Index into a for loop
    Dim strMask As String               ' Order type mask from the file
    
    bReturn = False
    
    For lIndex = 0 To m.astrSymbols.Size - 1
        If Parse(m.astrSymbols(lIndex), vbTab, 2) = strGenesisSymbol Then
            strMask = Parse(m.astrSymbols(lIndex), vbTab, 5)
            If Len(strMask) = 0 Then strMask = "111"
            If (nTimeInForce >= 0) And (nTimeInForce < Len(strMask)) Then
                bReturn = (Mid(strMask, nTimeInForce + 1, 1) = "1")
            End If
            
            Exit For
        End If
    Next lIndex
    
    TimeInForceAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.TimeInForceAllowed"
    
End Function

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
    Dim nStatus As eGDConnectionStatus  ' Connection status
    Dim strError As String              ' Error message back
    Dim frm As frmAlertPopup            ' Alert pop-up message
    Static strLastUser As String        ' Last user name
    
    dTickCount = Val(Parse(strConnectionInfo, vbTab, 1))
    
    If Len(Parse(strConnectionInfo, vbTab, 2)) > 0 Then
        nStatus = CLng(Val(Parse(strConnectionInfo, vbTab, 2)))
        strError = Parse(strConnectionInfo, vbTab, 3)
        
        Select Case nStatus
            Case eGDConnectionStatus_Disconnected
                m.BrokerInfo.SetConnectionStatus nStatus
                
                If (ConnectionStatus <> eGDConnectionStatus_Disconnected) Then
                    If (Len(strError) > 0) Or (m.bManualDisconnect = True) Then
                        m.dHeartBeat = 0#
                        m.bVerifyPositions = True
                    End If
                
                    If m.bWasConnected = True Then
                        m.bWasConnected = False
                        g.Alerts.CheckBrokerStatusAlert m.nBroker
                    End If
                    
                    If Len(strError) > 0 Then
                        InfBox strError, "!", "+-OK", "FXCM Connection Error", True
                        If InStr(strError, "Incorrect user name or password") <> 0 Then
                            m.bAskedPassword = False
                            m.strPassword = ""
                        End If
                    End If
                    
                    m.astrAccounts.Clear
                    
                    m.bRefreshingAccounts = False
                    m.bRefreshingOrders = False
                    m.bRefreshingFills = False
                    m.bRefreshingPositions = False
                    
                    DumpDebug "Unloading Application since we have been disconnected"
                    SendFxcmMessage eGDFxcmMessageType_UnloadApp, ""
                End If
            
            Case eGDConnectionStatus_Disconnecting
                m.BrokerInfo.SetConnectionStatusForUser nStatus, strLastUser
            
            Case eGDConnectionStatus_Connecting
                m.BrokerInfo.SetConnectionStatusForUser nStatus, m.strUserName
            
            Case eGDConnectionStatus_Connected
                If (ConnectionStatus <> eGDConnectionStatus_Connected) Then
                    GetAccounts
                End If
                m.bWasConnected = True
        
        End Select
        
        If m.strUserName <> strLastUser Then
            strLastUser = m.strUserName
        End If
        
        If (nStatus <> ConnectionStatus) Or (Len(strError) > 0) Then
            ConnectionStatus = nStatus
            
            g.ActivityLogs.AddStatusToActivityLog m.nBroker, nStatus
            
            DumpDebug vbTab & "FXCM Message Received (" & MessageTypeToString(eGDFxcmMessageType_ConnectionInfo) & "): " & strConnectionInfo
        End If
        
        If (Len(strError) > 0) And (strError <> m.strLastMessage) Then
            Set frm = New frmAlertPopup
            frm.ShowMessageBox strError, "FXCM Message", vbCenter
        End If
        m.strLastMessage = strError
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.HandleConnectionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleAccountInfo
'' Description: Do the appropriate thing with the given account information
'' Inputs:      Account Information
'' Returns:     None
''
'' Fields:      Account#, Balance, Equity, Margin Used, Margin Available
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleAccountInfo(ByVal strAccountInfo As String)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strAccountNumber As String      ' Account Number from the callback
    Dim lAccountID As Long              ' Account ID
    Dim lPos As Long                    ' Position of an item in an array

    If Len(strAccountInfo) > 0 Then
        ' Get Account number from Server String
        strAccountNumber = Parse(strAccountInfo, vbTab, 1)
        
        ' Get LotSize for trading if available
        m.lLotSize = Int(Val(Parse(strAccountInfo, vbTab, 7)))
        If m.lLotSize = 0 Then m.lLotSize = 1
      
        If (m.lLotSize >= 10000) Then
            m.dPipCost = m.lLotSize / 10000#
        Else
            m.dPipCost = 1#
        End If
        
        ' Get Hedging mode
        If Left$(Parse(strAccountInfo, vbTab, 8), 1) = "Y" Then
            m.bHedging = True
        Else
            m.bHedging = False
        End If
        
        
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                "WHERE [AccountNumber]='" & strAccountNumber & "';", dbOpenDynaset)
        If (rs.BOF And rs.EOF) Then
            DumpDebug "Creating account " & strAccountNumber
            rs.AddNew
            rs!AccountNumber = strAccountNumber
            rs!Name = strAccountNumber
            rs!StartingBalance = 0#
            rs!CurrentBalance = 0#
            rs!StartingDate = Date
            rs!UserName = m.strUserName
            EncryptField rs!Password, m.strPassword
            rs!AccountType = m.nBroker
            rs!Broker = ""
            rs!Comms = 0#
            rs!SecTypeMask = 4
            rs!FillRT = False
            rs!ClosedProfit = 0#
            lAccountID = rs!AccountID
            rs.Update
        Else
            lAccountID = rs!AccountID
        End If
        
        If m.astrAccounts.BinarySearch(strAccountNumber, lPos) = False Then
            m.astrAccounts.Add strAccountNumber, lPos
        End If
        
        RefreshAccount lAccountID
        
        SendFxcmMessage eGDFxcmMessageType_GetOrders, strAccountNumber
        SendFxcmMessage eGDFxcmMessageType_GetFills, strAccountNumber
        SendFxcmMessage eGDFxcmMessageType_GetPositions, strAccountNumber
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.HandleAccountInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleOrderInfo
'' Description: Do the appropriate thing with the given order information
'' Inputs:      Order Information, Coming from a Refresh?
'' Returns:     None
''
'' Fields:      GenesisID, BrokerID, FillID, Account#, Symbol, SymbolLong, Type,
''              Side, Qty, Limit, Stop, Strike, Expiration, Exchange, SecType,
''              TimeInForce, Maturity, Reserved, QtyExec, QtyRemain, QtyCum,
''              PriceExec, TimeExec, Event, Status, Reason
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleOrderInfo(ByVal strOrderInfo As String, Optional ByVal bRefresh As Boolean = False)
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim strAccountNumber As String      ' Account number
    Dim dMult As Double                 ' Price multiplier
    Dim Order As cPtOrder               ' Order to update
    Dim Fill As cPtFill                 ' Fill to update
    Dim bFillReceived As Boolean        ' Did we receive a fill?
    Dim bFillExisted As Boolean         ' Did the fill already exist?
    Dim nPrevStatus As eTT_OrderStatus  ' Previous order status
    Dim lIndex As Long                  ' Index into a for loop
    Dim lOldOrderID As Long             ' Order ID for the previous version of this order
    Dim OldOrder As cPtOrder            ' Old Order
    Dim frm As frmAlertPopup            ' Alert popup form
    Dim Bars As New cGdBars             ' Bars object
    Dim nOrderType As eTT_OrderType     ' Order Type
    
    If Len(strOrderInfo) > 0 Then
        astrOrder.SplitFields strOrderInfo, vbTab
        strGenesisSymbol = GenesisSymbol(astrOrder(4), dMult)
        strAccountNumber = astrOrder(3)
        
        If Len(strGenesisSymbol) > 0 Then
            SetBarProperties Bars, strGenesisSymbol
            
            'Load order if it exists, otherwise create a new order
            Set Order = LoadOrder(strAccountNumber, astrOrder(0), astrOrder(1), lOldOrderID)
            If Order Is Nothing Then Set Order = New cPtOrder
            
            With Order
                'Update order information
                nPrevStatus = .Status
                
                .GenesisOrderID = astrOrder(0)
'                If astrOrder(0) = "" Then .GenesisOrderID = astrOrder(1)
                .BrokerID = astrOrder(1)
                .AccountID = g.Broker.AccountIDForNumber(strAccountNumber)
                
                .Buy = (astrOrder(7) = "1") Or (UCase(astrOrder(7)) = "B")
                .Quantity = CLng(Abs(Val(astrOrder(8)))) / m.lLotSize
                .SymbolOrSymbolID = strGenesisSymbol
                nOrderType = GenesisOrderType(astrOrder(6))
                If Not ((nOrderType = eTT_OrderType_Market) And (.OrderType = eTT_OrderType_MIT)) Then
                    .OrderType = nOrderType
                    .LimitPrice = GenesisPrice(astrOrder(9), dMult)
                    .StopPrice = GenesisPrice(astrOrder(10), dMult)
                End If
                If .OrderDate = 0# Then .OrderDate = DateToDouble(astrOrder(22))
                .Message = astrOrder(25)
                
                .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.OrderDate, m.nBroker, .Symbol, False))
                .IsSnapshot = True
                
                If .OrderID = 0& Then
                    .Save
                    
                    If lOldOrderID <> 0& Then
                        Set OldOrder = New cPtOrder
                        If OldOrder.Load(lOldOrderID) Then
                            OldOrder.GenesisOrderID = ""
                            OldOrder.Save
                            RefreshOrder OldOrder
                        End If
                        
                        .PreviousBrokerID = OldOrder.BrokerID
                        .PreviousOrder = OldOrder
                        
                        OrderIDChanged lOldOrderID, .OrderID
                    End If
                End If
                
                'If the order is a Fill (has a FillID)
'                If Len(astrOrder(2)) > 0 Then
                 If astrOrder(24) = "Filled" Then
                    bFillReceived = True
                    bFillExisted = False
                    
                    For lIndex = 1 To .Fills.Count
'                        If .Fills(lIndex).BrokerID = astrOrder(2) Then
                        If .Fills(lIndex).BrokerID = astrOrder(1) Then
                            Set Fill = .Fills(lIndex)
                            bFillExisted = True
                            Exit For
                        End If
                    Next lIndex
                    
                    If Fill Is Nothing Then Set Fill = New cPtFill
                    With Fill
                        .AccountID = Order.AccountID
                        .AutoTradingItemID = Order.AutoTradeItemID
'                        .BrokerID = astrOrder(2)
                        .BrokerID = astrOrder(1)
                        .BrokerOrderID = Order.BrokerID
                        .Buy = Order.Buy
                        .FillDate = DateToDouble(astrOrder(22))
                        .IsManual = False
                        .IsSnapshot = True
                        .OrderID = Order.OrderID
                        .Price = GenesisPrice(astrOrder(21), dMult)
                        .Quantity = CLng(Abs(Val(astrOrder(18)))) / m.lLotSize
                        .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.FillDate, m.nBroker, .Symbol, False))
                        .Symbol = Order.Symbol
                        .SymbolID = Order.SymbolID
                        
                        ' If we don't have a Fill ID yet (this is a new fill), we need to save
                        ' the fill to get a Fill ID for the key into the Fills collection of
                        ' the order...
                        If .FillID = 0& Then .Save
                    End With
                    
                    Set .Fills(Str(Fill.FillID)) = Fill
                    DumpDebug Fill.FillText
                Else
                    bFillReceived = False
                    bFillExisted = False
                End If
                
                .Status = GenesisOrderStatus(astrOrder(24), astrOrder(23))
                .StatusDate = DateToDouble(astrOrder(22))
                If (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_ParkPending) Then
                    .Status = eTT_OrderStatus_Parked
                ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_Parked) Then
                    .Status = eTT_OrderStatus_Parked
                ElseIf (Order.FillQuantity > Order.Quantity) Then
                    .Status = eTT_OrderStatus_OverFilled
                ElseIf (.Status = eTT_OrderStatus_Filled) And (Order.FillQuantity < Order.Quantity) Then
                    .Status = eTT_OrderStatus_Partial
                ElseIf (.Status = eTT_OrderStatus_Working) And (Order.FillQuantity > 0) Then
                    .Status = eTT_OrderStatus_Partial
                ElseIf (.Status = eTT_OrderStatus_Cancelled) And (Order.FillQuantity = Order.Quantity) Then
                    .Status = eTT_OrderStatus_Filled
                ElseIf (.Status = eTT_OrderStatus_Cancelled) And (Order.FillQuantity > 0) Then
                    .Status = eTT_OrderStatus_BalCancelled
                ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_Amended) Then
                    .Status = eTT_OrderStatus_Amended
                End If
                
                .Save
                
                If bFillReceived Then
                    m.BrokerInfo.AddOrder Order, bRefresh
                    m.BrokerInfo.AddFill Fill, bRefresh
                    
                    FillCallback Order, Fill, bFillExisted
                    If UCase(astrOrder(23)) <> "REWIND" Then
                        g.ActivityLogs.AddFillToActivityLog m.nBroker, Fill, False
                        m.dLastFillTime = gdTickCount
                    End If
                Else
                    m.BrokerInfo.AddOrder Order, bRefresh
                
                    OrderCallback Order
                    If UCase(astrOrder(23)) <> "REWIND" Then
                        g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                    End If
                End If
                
                ' If there was an Error message
                If (Len(astrOrder(25)) > 0) And (UCase(astrOrder(23)) <> "REWIND") Then
                    Set frm = New frmAlertPopup
                    frm.ShowMessageBox Order.OrderText & " REJECTED||" & astrOrder(25), "FXCM Order " & Order.BrokerID & " Rejected", vbLeftJustify
                End If
                
                If (Order.Status = eTT_OrderStatus_OverFilled) And (UCase(astrOrder(23)) = "REWIND") Then
                    Set frm = New frmAlertPopup
                    frm.ShowMessageBox "Trade Navigator has received duplicate fill|information for order " & Order.BrokerID & ".||This will cause this order to appear over-filled and will likely result in a position mismatch.||PLEASE CALL YOUR BROKER AND VERIFY YOUR POSITIONS.", "Over Filled Order on FXCM", vbCenter
                End If
    
                DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
            End With 'With Order
        End If 'If Len(strGenesisSymbol) > 0 Then
    End If  'If Len(strOrderInfo) > 0 Then

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.HandleOrderInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandlePositionInfo
'' Description: Do the appropriate thing with the given position information
'' Inputs:      Position Information
'' Returns:     None
''
'' Fields:      Account#, Symbol, Position, Realized Profit, Unrealized Profit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandlePositionInfo(ByVal strPositionInfo As String)
On Error GoTo ErrSection:

    Dim astrPosition As New cGdArray    ' Array of position information
'   Dim strGenesisSymbol As String      ' Genesis symbol
'   Dim dMult As Double                 ' Multiplier for the symbol
    Dim lIndex As Integer
    Dim bFound As Boolean
    
    bFound = False
    
    If Len(strPositionInfo) > 0 Then
        astrPosition.SplitFields strPositionInfo, vbTab
        For lIndex = 0 To m.astrPositions.Size - 1
            If Parse(m.astrPositions(lIndex), vbTab, 13) = astrPosition(12) Then
                m.astrPositions(lIndex) = strPositionInfo
                bFound = True
                Exit For
            End If
        Next lIndex
    
        If Not bFound Then
            m.astrPositions.Add strPositionInfo
        End If
    
'        strGenesisSymbol = GenesisSymbol(astrPosition(1), dMult)
    End If
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.HandlePositionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FxcmSymbol
'' Description: Convert a Genesis symbol to a FXCM symbol
'' Inputs:      Genesis symbol, Multiplier
'' Returns:     FXCM Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FxcmSymbol(ByVal strGenesisSymbol As String, Optional dMult As Double) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return from the function
    Dim lIndex As Long                  ' Index into a for loop
    
    strReturn = ""
    
    For lIndex = 0 To m.astrSymbols.Size - 1
        If Parse(m.astrSymbols(lIndex), vbTab, 2) = strGenesisSymbol Then
            strReturn = Parse(m.astrSymbols(lIndex), vbTab, 1)
            dMult = Val(Parse(m.astrSymbols(lIndex), vbTab, 3))
            Exit For
        End If
    Next lIndex

    FxcmSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.FxcmSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Convert a FXCM symbol to a Genesis symbol
'' Inputs:      FXCM symbol, Multiplier
'' Returns:     Genesis Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisSymbol(ByVal strFxcmSymbol As String, dMult As Double) As String
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position of base symbol in the array
    Dim strReturn As String             ' Return from the function
    
    strReturn = ""
    
    If m.astrSymbols.BinarySearch(strFxcmSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        strReturn = Parse(m.astrSymbols(lPos), vbTab, 2)
        dMult = Val(Parse(m.astrSymbols(lPos), vbTab, 3))
    End If

    GenesisSymbol = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.GenesisSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FxcmOrderType
'' Description: Convert a Genesis order type to a FXCM order type
'' Inputs:      Genesis order type
'' Returns:     FXCM Order Type (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FxcmOrderType(ByVal nOrderType As eTT_OrderType) As String
On Error GoTo ErrSection:

    Select Case nOrderType
        Case eTT_OrderType_Market
            FxcmOrderType = "1"
        Case eTT_OrderType_Limit
            FxcmOrderType = "2"
        Case eTT_OrderType_Stop
            FxcmOrderType = "3"
        Case Else
            FxcmOrderType = ""
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.FxcmOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderType
'' Description: Convert a FXCM order type to a Genesis order type
'' Inputs:      FXCM order type
'' Returns:     Genesis Order Type (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderType(ByVal strFxcmOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Select Case strFxcmOrderType
        Case "1"
            GenesisOrderType = eTT_OrderType_Market
        Case "2"
            GenesisOrderType = eTT_OrderType_Limit
        Case "3"
            GenesisOrderType = eTT_OrderType_Stop
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.GenesisOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FxcmPrice
'' Description: Convert a Genesis price to a FXCM price
'' Inputs:      Genesis Price, Multiplier
'' Returns:     FXCM Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FxcmPrice(ByVal dGenesisPrice As Double, ByVal dMult As Double) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from the function
    
    If dMult <> 0 Then
        Select Case dMult
            Case 32
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 32, "00")
            
            Case 32.5
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 320, "000")
            
            Case Else
                strReturn = Str(dGenesisPrice / dMult)

        End Select
    End If
    
    FxcmPrice = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.FxcmPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisPrice
'' Description: Convert a FXCM price to a Genesis price
'' Inputs:      FXCM Price, Multiplier
'' Returns:     Genesis Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisPrice(ByVal strFxcmPrice As String, ByVal dMult As Double) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value from the function
    
    If dMult <> 0 Then
        Select Case dMult
            Case 32
                dReturn = Val(Left(strFxcmPrice, Len(strFxcmPrice) - 2)) + (Val(Right(strFxcmPrice, 2)) / 32)
            
            Case 32.5
                dReturn = Val(Left(strFxcmPrice, Len(strFxcmPrice) - 3)) + ((Val(Right(strFxcmPrice, 3)) / 10) / 32)
            
            Case Else
                dReturn = Val(strFxcmPrice) * dMult
                
        End Select
    End If
    
    GenesisPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.GenesisPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderStatus
'' Description: Convert a FXCM order status to a Genesis order status
'' Inputs:      FXCM order status
'' Returns:     Genesis order status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderStatus(ByVal strFxcmOrderStatus As String, ByVal strFxcmOrderEvent As String) As eTT_OrderStatus
On Error GoTo ErrSection:

    Select Case UCase(strFxcmOrderStatus)
        Case "PENDING"
            GenesisOrderStatus = eTT_OrderStatus_Working
        Case "FILLED"
            GenesisOrderStatus = eTT_OrderStatus_Filled
        Case "CANCELED"
            If UCase(strFxcmOrderEvent) = "REPLACE" Then
                GenesisOrderStatus = eTT_OrderStatus_Amended
            Else
                GenesisOrderStatus = eTT_OrderStatus_Cancelled
            End If
        Case "REJECTED"
            GenesisOrderStatus = eTT_OrderStatus_Rejected
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.GenesisOrderStatus"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrder
'' Description: Load the order from the database for Genesis ID or Broker ID
'' Inputs:      Account Number, Genesis Order ID, Broker ID
'' Returns:     Order (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadOrder(ByVal strAccountNumber As String, ByVal strGenesisOrderID As String, ByVal strBrokerID As String, Optional lOldOrderID As Long) As cPtOrder
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lAccountID As Long              ' Account ID for the given account number
    Dim Order As cPtOrder               ' Order to return
    Dim bFound As Boolean               ' Have we found the order?

    bFound = False
    lAccountID = g.Broker.AccountIDForNumber(strAccountNumber)
    lOldOrderID = 0&
        
    Set Order = m.BrokerInfo.Order(strBrokerID, strGenesisOrderID)
    If Not Order Is Nothing Then
        If (Len(strBrokerID) = 0) Or (Len(Order.BrokerID) = 0) Or (Order.BrokerID = strBrokerID) Then
            bFound = True
        Else
            bFound = False
            lOldOrderID = Order.OrderID
        End If
    End If
    
    If bFound = False Then
        If Len(strGenesisOrderID) > 0 Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                        "WHERE [GenesisOrderID]='" & strGenesisOrderID & "' AND [AccountID]=" & lAccountID & ";", dbOpenDynaset)
            If Not (rs.EOF And rs.BOF) Then
                Set Order = New cPtOrder
                If Order.Load(rs!OrderID) Then
                    If (Len(strBrokerID) = 0) Or (Len(Order.BrokerID) = 0) Or (Order.BrokerID = strBrokerID) Then
                        bFound = True
                    Else
                        bFound = False
                        lOldOrderID = Order.OrderID
                    End If
                End If
            End If
        End If
        If (bFound = False) And (Len(strBrokerID) > 0) Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                        "WHERE [BrokerOrderID]='" & strBrokerID & "' AND [AccountID]=" & lAccountID & ";", dbOpenDynaset)
            If Not (rs.EOF And rs.BOF) Then
                Set Order = New cPtOrder
                If Order.Load(rs!OrderID) Then
                    bFound = True
                End If
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
    RaiseError "cFXCM.LoadOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MessageTypeToString
'' Description: Convert a message type enumeration to a string
'' Inputs:      Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MessageTypeToString(ByVal nType As eGDFxcmMessageTypes) As String
On Error GoTo ErrSection:

    Select Case nType
        Case eGDFxcmMessageType_Connect
            MessageTypeToString = "eGDFxcmMessageType_Connect"
        Case eGDFxcmMessageType_Disconnect
            MessageTypeToString = "eGDFxcmMessageType_Disconnect"
        Case eGDFxcmMessageType_AddOrder
            MessageTypeToString = "eGDFxcmMessageType_AddOrder"
        Case eGDFxcmMessageType_AmendOrder
            MessageTypeToString = "eGDFxcmMessageType_AmendOrder"
        Case eGDFxcmMessageType_CancelOrder
            MessageTypeToString = "eGDFxcmMessageType_CancelOrder"
        Case eGDFxcmMessageType_UnloadApp
            MessageTypeToString = "eGDFxcmMessageType_UnloadApp"
        Case eGDFxcmMessageType_GetAccounts
            MessageTypeToString = "eGDFxcmMessageType_GetAccounts"
        Case eGDFxcmMessageType_GetOrders
            MessageTypeToString = "eGDFxcmMessageType_GetOrders"
        Case eGDFxcmMessageType_GetFills
            MessageTypeToString = "eGDFxcmMessageType_GetFills"
        Case eGDFxcmMessageType_GetPositions
            MessageTypeToString = "eGDFxcmMessageType_GetPositions"
        Case eGDFxcmMessageType_ClosePosition
            MessageTypeToString = "eGDFxcmMessageType_ClosePosition"
        Case eGDFxcmMessageType_PositionStopLimit
            MessageTypeToString = "eGDFxcmMessageType_PositionStopLimit"
        Case eGDFxcmMessageType_DeleteStopLimit
            MessageTypeToString = "eGDFxcmMessageType_DeleteStopLimit"
        
        Case eGDFxcmMessageType_ConnectionInfo
            MessageTypeToString = "eGDFxcmMessageType_ConnectionInfo"
        Case eGDFxcmMessageType_AppLoaded
            MessageTypeToString = "eGDFxcmMessageType_AppLoaded"
        Case eGDFxcmMessageType_AppUnloaded
            MessageTypeToString = "eGDFxcmMessageType_AppUnloaded"
        Case eGDFxcmMessageType_Heartbeat
            MessageTypeToString = "eGDFxcmMessageType_Heartbeat"
        Case eGDFxcmMessageType_Order
            MessageTypeToString = "eGDFxcmMessageType_Order"
        Case eGDFxcmMessageType_AccountR
            MessageTypeToString = "eGDFxcmMessageType_AccountR"
        Case eGDFxcmMessageType_OrderR
            MessageTypeToString = "eGDFxcmMessageType_OrderR"
        Case eGDFxcmMessageType_FillR
            MessageTypeToString = "eGDFxcmMessageType_FillR"
        Case eGDFxcmMessageType_PositionR
            MessageTypeToString = "eGDFxcmMessageType_PositionR"
        Case eGDFxcmMessageType_ContractR
            MessageTypeToString = "eGDFxcmMessageType_ContractR"
        Case eGDFxcmMessageType_QuoteR
            MessageTypeToString = "eGDFxcmMessageType_QuoteR"
            
        Case Else
            MessageTypeToString = Str(nType)
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFXCM.MessageTypeToString"
    
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
    Open AddSlash(App.Path) & "FXCM\TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append Shared As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " - " & strMessage
        Close #fh
    End If

End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBrokerInfoOrder
'' Description: Add the order string to the broker info class
'' Inputs:      Order String, Refresh?
'' Returns:     None
''
'' RecordFrom:  Genesis ID, Order ID, Fill ID, Account, Base Symbol, Symbol,
''              Order Type, B/S, Quantity, Limit, Stop, Strike, Expiration,
''              Exchange, Instrument, Time In Force, Maturity, Reserved,
''              Quantity Filled, Remaining Quantity, Quantity Cum, Fill Price,
''              Fill Time, Event, Status, Message
'' RecordTo:    Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateBrokerInfoOrder(ByVal strOrderString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Order information split out from string
    Dim astrToAdd As cGdArray           ' Order information to add to broker info
    Dim dMult As Double                 ' Price Multiplier
    
    Set astrOrder = New cGdArray
    astrOrder.Create eGDARRAY_Strings
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings
    
    astrOrder.SplitFields strOrderString, vbTab
    
    astrToAdd(0) = astrOrder(1)
    astrToAdd(1) = astrOrder(0)
    astrToAdd(2) = ""
    astrToAdd(3) = astrOrder(3)
    astrToAdd(4) = GenesisSymbol(astrOrder(4), dMult)
    astrToAdd(5) = Str(GenesisOrderType(astrOrder(6)))
    If UCase(astrOrder(7)) = "B" Then astrToAdd(6) = Str(CLng(True)) Else astrToAdd(6) = Str(CLng(False))
    astrToAdd(7) = Str(CLng(Abs(Val(astrOrder(8)))) / m.lLotSize)
    astrToAdd(8) = Str(GenesisPrice(astrOrder(9), dMult))
    astrToAdd(9) = Str(GenesisPrice(astrOrder(10), dMult))
    astrToAdd(10) = "0"
    astrToAdd(11) = Str(GenesisOrderStatus(astrOrder(24), astrOrder(23)))
    astrToAdd(12) = astrOrder(19)
    astrToAdd(13) = Str(DateToDouble(astrOrder(22)))
    
    m.BrokerInfo.AddOrder astrToAdd.JoinFields(vbTab), bRefresh
    
    If Len(astrOrder(2)) > 0 Then
       AddFillToBrokerInfo strOrderString, bRefresh
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.UpdateBrokerInfoOrder"
    
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

    Dim astrToAdd As cGdArray           ' Order information to add to broker info
    
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings
    
    With Order
        astrToAdd(0) = .BrokerID
        astrToAdd(1) = .GenesisOrderID
        astrToAdd(2) = ""
        astrToAdd(3) = g.Broker.AccountNumberForID(.AccountID)
        astrToAdd(4) = .Symbol
        astrToAdd(5) = Str(.OrderType)
        astrToAdd(6) = Str(CLng(.Buy))
        astrToAdd(7) = Str(.Quantity)
        astrToAdd(8) = Str(.LimitPrice)
        astrToAdd(9) = Str(.StopPrice)
        astrToAdd(10) = Str(.Expiration)
        astrToAdd(11) = Str(nStatus)
        astrToAdd(12) = "0"
        astrToAdd(13) = Str(.OrderDate)
    
        If .AutoTradeItemID <> 0& Then
            m.BrokerInfo.AddAtOrder .AutoTradeItemID, .BrokerID, .GenesisOrderID
        End If
    End With
    
    m.BrokerInfo.AddOrder astrToAdd.JoinFields(vbTab), False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.GenesisOrderToBrokerInfo"
    
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
'' RecordTo:    Genesis ID, Broker ID, Account, Base Symbol, Symbol, Type,
''              B/S, Qty, Limit, Stop, Strike, Expiration, Exchange, SecType,
''              Time In Force, Maturity, Reserved
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelBrokerInfoOrder(ByVal strOrder As String)
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Order broken out into an array
    Dim astrToCancel As cGdArray        ' Order to send to the cancel message
    Dim strFxcmSymbol As String         ' Fxcm Symbol
    Dim dMult As Double                 ' Multiplier for converting prices

    If Len(strOrder) > 0 Then
        Set astrOrder = New cGdArray
        astrOrder.Create eGDARRAY_Strings
        Set astrToCancel = New cGdArray
        astrToCancel.Create eGDARRAY_Strings
    
        astrOrder.SplitFields strOrder, vbTab
        strFxcmSymbol = FxcmSymbol(astrOrder(4), dMult)
        If Len(strFxcmSymbol) > 0 Then
            astrToCancel(0) = astrOrder(1)
            astrToCancel(1) = astrOrder(0)
            astrToCancel(2) = astrOrder(3)
            astrToCancel(3) = strFxcmSymbol
            astrToCancel(4) = ""
            astrToCancel(5) = FxcmOrderType(CLng(Val(astrOrder(5))))
            If astrOrder(6) = "0" Then astrToCancel(6) = "2" Else astrToCancel(6) = "1"
            astrToCancel(7) = Str(CLng(Val(astrOrder(7))) * m.lLotSize)
            astrToCancel(8) = FxcmPrice(Val(astrOrder(8)), dMult)
            astrToCancel(9) = ""
            astrToCancel(10) = ""
            astrToCancel(11) = ""
            astrToCancel(12) = ""
            astrToCancel(13) = ""
            astrToCancel(14) = ""
            astrToCancel(15) = ""
            astrToCancel(16) = ""
            
            SendFxcmMessage eGDFxcmMessageType_CancelOrder, astrToCancel.JoinFields(vbTab)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.CancelBrokerInfoOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFillToBrokerInfo
'' Description: Add the fill string to the broker info class
'' Inputs:      Fill String, Refresh?
'' Returns:     None
''
'' RecordFrom:  Genesis ID, Order ID, Fill ID, Account, Base Symbol, Symbol,
''              Order Type, B/S, Quantity, Limit, Stop, Strike, Expiration,
''              Exchange, Instrument, Time In Force, Maturity, Reserved,
''              Quantity Filled, Remaining Quantity, Quantity Cum, Fill Price,
''              Fill Time, Event, Status, Message
'' RecordTo:    Broker ID, Genesis ID, Fill ID, Account, Symbol, Fill Date,
''              B/S, Fill Quantity, Fill Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddFillToBrokerInfo(ByVal strFillString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim astrFill As cGdArray            ' Fill information split from the string
    Dim astrToAdd As cGdArray           ' Array to add to the broker information
    Dim dMult As Double                 ' Price Multiplier
    
    Set astrFill = New cGdArray
    astrFill.Create eGDARRAY_Strings
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings
    
    astrFill.SplitFields strFillString, vbTab
    
    astrToAdd(0) = astrFill(1)
    astrToAdd(1) = astrFill(0)
    astrToAdd(2) = astrFill(2)
    astrToAdd(3) = astrFill(3)
    astrToAdd(4) = GenesisSymbol(astrFill(4), dMult)
    astrToAdd(5) = Str(DateToDouble(astrFill(22)))
    If UCase(astrFill(7)) = "B" Then astrToAdd(6) = Str(CLng(True)) Else astrToAdd(6) = Str(CLng(False))
    astrToAdd(7) = Str(CLng(Abs(Val(astrFill(18)))) / m.lLotSize)
    astrToAdd(8) = Str(GenesisPrice(astrFill(21), dMult))
    
    m.BrokerInfo.AddFill astrToAdd.JoinFields(vbTab), bRefresh

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.AddFillToBrokerInfo"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddPositionToBrokerInfo
'' Description: Add the position string to the broker info class
'' Inputs:      Position String, Refresh?
'' Returns:     None
''
'' RecordFrom:  Account, Symbol, Position, Average Entry, Entry Sum, Entry Count
'' RecordTo:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddPositionToBrokerInfo(ByVal strPositionString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim astrPosition As cGdArray        ' Array of position information split out from string
    Dim astrToAdd As cGdArray           ' Array of information to pass to broker info class
    Dim dMult As Double                 ' Price Multiplier
    Dim FillSum As cAccountPosition     ' Collection to return Fill Summary
    
    Set astrPosition = New cGdArray
    astrPosition.Create eGDARRAY_Strings
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings
        
    astrPosition.SplitFields strPositionString, vbTab
        
    astrToAdd(0) = astrPosition(0)
    astrToAdd(1) = GenesisSymbol(astrPosition(1), dMult)
    astrToAdd(2) = Str(CLng(Val(astrPosition(2))) / m.lLotSize)
    astrToAdd(3) = GenesisPrice(astrPosition(3), dMult)
    
    'For overnight postition use CurrentPosition - Buys + Sells
    Set FillSum = m.BrokerInfo.FillSummaryForSymbol(astrToAdd(0), astrToAdd(1), -1)
    If Not FillSum Is Nothing Then
        astrToAdd(4) = Str(FillSum.CurrentPositionSnapshot - FillSum.NumBuysSnapshot + FillSum.NumSellsSnapshot)
    Else
        astrToAdd(4) = "0"
    End If
    
    m.BrokerInfo.AddPosition astrToAdd.JoinFields(vbTab), bRefresh

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.AddPositionToBrokerInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BeginPosition
'' Description: Clear the Consolidated Position Array
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BeginPosition()
On Error GoTo ErrSection:

    m.astrConsolidated.Clear

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.BeginPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EndPosition
'' Description: Add the Consolidated position string(s) to the broker info class
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EndPosition()
On Error GoTo ErrSection:

    Dim lIndex As Integer

    For lIndex = 0 To m.astrConsolidated.Size - 1
        AddPositionToBrokerInfo m.astrConsolidated(lIndex), True
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.BeginPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddPosition
'' Description: Add the position string to the Consolidated Position Array
'' Inputs:      Position String
'' Returns:     None
''
'' RecordFrom:  Account, Symbol, Position, Realized, Unrealized, Avg Entry
'' RecordTo:    Account, Symbol, Position, Average Entry, Entry Sum, Entry Count
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddPosition(ByVal strPositionString As String)
On Error GoTo ErrSection:

    Dim astrPosition As cGdArray        ' Array of position information split out from string
    Dim astrToAdd As cGdArray           ' Array of information to pass to broker info class\
    Dim astrOldPosition As cGdArray    ' Array for existing position
    Dim bFound As Boolean               ' Found position flag
    Dim lIndex As Integer               ' Index for Array
    
    bFound = False
    Set astrPosition = New cGdArray
    astrPosition.Create eGDARRAY_Strings
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings
    Set astrOldPosition = New cGdArray
    astrOldPosition.Create eGDARRAY_Strings
        
    astrPosition.SplitFields strPositionString, vbTab
        
    ' Check to see if this Account and Symbol are already in array, if found update record
    For lIndex = 0 To m.astrConsolidated.Size - 1
        If Parse(m.astrConsolidated(lIndex), vbTab, 2) = astrPosition(1) Then
            astrOldPosition.SplitFields m.astrConsolidated(lIndex), vbTab
            If astrOldPosition(0) = astrPosition(0) Then
                astrToAdd(0) = astrPosition(0)
                astrToAdd(1) = astrPosition(1)
                astrToAdd(2) = Str(CLng(Val(astrOldPosition(2))) + CLng(Val(astrPosition(3))))
                astrToAdd(4) = Str(Val(astrOldPosition(4)) + Abs((Val(astrPosition(3)) * Val(astrPosition(6)))))
                astrToAdd(5) = Str(Val(astrOldPosition(5)) + Abs(Val(astrPosition(3))))
                astrToAdd(3) = Str(Val(astrToAdd(4)) / Abs(Val(astrToAdd(5))))
                m.astrConsolidated(lIndex) = astrToAdd.JoinFields(vbTab)
                bFound = True
                Exit For
            End If
        End If
    Next lIndex
        
    'If Account and Symbol not in Array, Add new record
    If Not bFound Then
        astrToAdd(0) = astrPosition(0)
        astrToAdd(1) = astrPosition(1)
        astrToAdd(2) = astrPosition(3)
        astrToAdd(3) = astrPosition(6)
        astrToAdd(4) = Str(Val(astrToAdd(3)) * Val(astrToAdd(2)))
        astrToAdd(5) = Str(Abs(Val(astrToAdd(2))))
        m.astrConsolidated.Add astrToAdd.JoinFields(vbTab)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.AddPosition"
    
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

    m.strIniFile = AddSlash(App.Path) & "FXCM.INI"
    m.nBroker = eTT_AccountType_FXCM

    If Not DirExist(AddSlash(App.Path) & "FXCM") Then MkDir AddSlash(App.Path) & "FXCM"
    KillFile AddSlash(App.Path) & "FXCM\*.LOG /o=-14"
    
    Set m.BrokerInfo = New cBrokerInfo
    m.BrokerInfo.Init m.nBroker, AddSlash(App.Path) & "Fxcm"
    
    Set m.astrSymbols = New cGdArray
    m.astrSymbols.Create eGDARRAY_Strings
    m.astrSymbols.FromFile AddSlash(App.Path) & "Provided\FxcmToGen.TXT"
    m.astrSymbols.Sort
    
    Set m.astrAccounts = New cGdArray
    m.astrAccounts.Create eGDARRAY_Strings
    
    Set m.astrPositions = New cGdArray
    m.astrPositions.Create eGDARRAY_Strings
    
    Set m.astrConsolidated = New cGdArray
    m.astrPositions.Create eGDARRAY_Strings
            
    m.bWasConnected = False
    m.bVerifyPositions = True
    
    m.lLotSize = 10000
    m.dPipCost = 1#
    
    m.bHedging = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.Class_Initialize"
    
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
    Set m.astrSymbols = Nothing
    Set m.astrAccounts = Nothing
    Set m.astrPositions = Nothing
    Set m.astrConsolidated = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFXCM.Class_Terminate"
    
End Sub
