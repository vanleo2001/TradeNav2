VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPATS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPATS.cls
'' Description: Wrapper class for routines to access PATS API
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/24/2009   DAJ         Fixed GenesisPrice for non-decimal bases
'' 03/24/2009   DAJ         Added support for 64ths and 1/2 64ths
'' 06/09/2009   DAJ         Set the BrokerInfo connection status
'' 06/09/2090   DAJ         Fixed the Quarter 32nds
'' 09/01/2009   DAJ         Use new Parked order status
'' 03/11/2010   DAJ         Use global collections
'' 03/07/2011   DAJ         Broker Disconnected Alerts
'' 06/21/2011   DAJ         Last connection status change
'' 01/18/2012   DAJ         Enhanced logging for automated trading
'' 01/31/2012   DAJ         Handle unsolicted cancels in TradeSense order groups
'' 02/14/2012   DAJ         New status alerts for position mismatch / auto trade disabled
'' 03/21/2012   DAJ         Fix for amending a partially filled order
'' 05/03/2012   DAJ         Force month name abbreviations to be in English
'' 08/03/2012   DAJ         Remove Robbins
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kAppID = "GFTTradeNav2.0"
Private Const kAppLicense = "KAD104V001"
'Private Const kAppID = "APPID_2284"
'Private Const kAppLicense = "BETA"
Private Const kAppVersion = "3.05"

Private Type mPrivate
    strIniFile As String                ' Ini file for the PATS information
    nBroker As eTT_AccountType          ' Broker type for the TT broker
    strBrokerName As String             ' Broker name for the TT broker
    strProcessName As String            ' Name of the process to check
    strProcessPath As String            ' Full path of the process to start
    strDLLPath As String                ' Full path of the PatsAPI DLL
    strControlID As String              ' Control ID
    strLogPath As String                ' Path for the log files

    strUserName As String               ' Username to log onto the PATS servers
    strPassword As String               ' Password to logon to the server with
    strIniPassword As String            ' Password to log on with
    strHostIP As String                 ' Host IP for the PATS server
    strHostPort As String               ' Host port for the PATS server
    strPriceIP As String                ' Price IP for the PATS servers
    strPricePort As String              ' Price Port for the PATS servers
    strEnvironment As String            ' Environment of the PATS server
    strSuperTAS As String               ' Enable SuperTAS?
    
    nStatus As eGDConnectionStatus      ' Connection status to the PATS servers
    bWasConnected As Boolean            ' This broker was has been connected
    dHeartBeat As Double                ' Tick count as of last message receieved
    bReconnecting As Boolean            ' Are we currently reconnecting?
    lReconnectAttempts As Long          ' Reconnect attempts that have been made
    bAppLoaded As Boolean               ' Is the PATS application loaded?
    bManualDisconnect As Boolean        ' Is the user manually disconnecting?
    bAskedPassword As Boolean           ' Have we asked for the password yet?
    dLastFillTime As Double             ' Tick count of when we received the last fill
    strLastMessage As String            ' Last message from the connection status message
    bVerifyPositions As Boolean         ' Do we want to verify the positions with the user?
    bVerboseRefresh As Boolean          ' Should we tell the user what we are doing?
    bSyncInProgress As Boolean          ' Are we in the process of synchronizing positions?
    bFixCalledOnce As Boolean           ' Has the fix positions call been done at least once?
    
    astrAccounts As cGdArray            ' List of accounts currently connected to
    astrBrokerPos As cGdArray           ' List of current positions
    astrToFix As cGdArray               ' List of account/symbol pairs that need their position fixed
    
    dRefreshingStart As Double          ' Start of the refreshing time
    bRefreshingAccounts As Boolean      ' Are we refreshing accounts?
    bRefreshingOrders As Boolean        ' Are we refreshing orders?
    bRefreshingFills As Boolean         ' Are we refreshing fills?
    bRefreshingPositions As Boolean     ' Are we refreshing positions?
    
    astrSymbols As cGdArray             ' Array of symbol conversion information
    
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

Public Property Get HostIP() As String
    HostIP = m.strHostIP
End Property
Public Property Let HostIP(ByVal strHostIP As String)
    m.strHostIP = strHostIP
End Property

Public Property Get HostPort() As String
    HostPort = m.strHostPort
End Property
Public Property Let HostPort(ByVal strHostPort As String)
    m.strHostPort = strHostPort
End Property

Public Property Get PriceIP() As String
    PriceIP = m.strPriceIP
End Property
Public Property Let PriceIP(ByVal strPriceIP As String)
    m.strPriceIP = strPriceIP
End Property

Public Property Get PricePort() As String
    PricePort = m.strPricePort
End Property
Public Property Let PricePort(ByVal strPricePort As String)
    m.strPricePort = strPricePort
End Property

Public Property Get Environment() As String
    Environment = m.strEnvironment
End Property
Public Property Let Environment(ByVal strEnvironment As String)
    m.strEnvironment = strEnvironment
End Property

Public Property Get SuperTAS() As String
    SuperTAS = m.strSuperTAS
End Property
Public Property Let SuperTAS(ByVal strSuperTAS As String)
    m.strSuperTAS = strSuperTAS
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPosition
'' Description: Retrieve the broker reported current position
'' Inputs:      Symbol, Account
'' Returns:     Current Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get CurrentPosition(ByVal strAccountNumber As String, ByVal strSymbol As String) As Long
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = 0&
    If m.astrBrokerPos.BinarySearch(strAccountNumber & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        lReturn = CLng(Val(Parse(m.astrBrokerPos(lPos), vbTab, 3)))
    End If
    
    CurrentPosition = lReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPATS.CurrentPosition.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPosition
'' Description: Set the broker reported current position
'' Inputs:      Symbol, Account, Current Position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let CurrentPosition(ByVal strAccountNumber As String, ByVal strSymbol As String, ByVal lCurrentPosition As Long)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    
    If m.astrBrokerPos.BinarySearch(strAccountNumber & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        m.astrBrokerPos(lPos) = strAccountNumber & vbTab & strSymbol & vbTab & Str(lCurrentPosition) & vbTab & Parse(m.astrBrokerPos(lPos), vbTab, 4) & vbTab & Parse(m.astrBrokerPos(lPos), vbTab, 5) & vbTab & Parse(m.astrBrokerPos(lPos), vbTab, 6)
    Else
        m.astrBrokerPos.Add strAccountNumber & vbTab & strSymbol & vbTab & Str(lCurrentPosition) & vbTab & "0" & vbTab & "" & vbTab & "", lPos
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPATS.CurrentPosition.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CarriedPosition
'' Description: Retrieve the broker reported carried position
'' Inputs:      Symbol, Account
'' Returns:     Carried Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get CarriedPosition(ByVal strAccountNumber As String, ByVal strSymbol As String) As Long
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = 0&
    If m.astrBrokerPos.BinarySearch(strAccountNumber & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        lReturn = CLng(Val(Parse(m.astrBrokerPos(lPos), vbTab, 4)))
    End If
    
    CarriedPosition = lReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPATS.CarriedPosition.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CarriedPosition
'' Description: Set the broker reported carried position
'' Inputs:      Symbol, Account, Carried Position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let CarriedPosition(ByVal strAccountNumber As String, ByVal strSymbol As String, ByVal lCarriedPosition As Long)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    
    If m.astrBrokerPos.BinarySearch(strAccountNumber & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        m.astrBrokerPos(lPos) = strAccountNumber & vbTab & strSymbol & vbTab & Parse(m.astrBrokerPos(lPos), vbTab, 3) & vbTab & Str(lCarriedPosition) & vbTab & Parse(m.astrBrokerPos(lPos), vbTab, 5) & vbTab & Parse(m.astrBrokerPos(lPos), vbTab, 6)
    Else
        m.astrBrokerPos.Add strAccountNumber & vbTab & strSymbol & vbTab & "0" & vbTab & Str(lCarriedPosition) & vbTab & "" & vbTab & "", lPos
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPATS.CarriedPosition.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the object based on what broker it is
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:
    
    Dim strSymbolFile As String         ' Symbol conversion file
    Dim rs As Recordset                 ' Recordset into the database

    Select Case nBroker
        Case eTT_AccountType_PATS
            m.strIniFile = AddSlash(App.Path) & "Pats3.INI"
            m.nBroker = eTT_AccountType_PATS
            m.strBrokerName = "PATS"
            m.strProcessName = "GenPATS"
            m.strProcessPath = AddSlash(App.Path) & "..\Brokers\Pats\GenPats.EXE"
            m.strDLLPath = AddSlash(App.Path) & "..\Brokers\Pats\PatsAPI.DLL"
            m.strControlID = "Pats"
            m.strLogPath = AddSlash(App.Path) & "Pats"
            strSymbolFile = AddSlash(App.Path) & "Provided\PatsToGen.TXT"
        
'        Case eTT_AccountType_Robbins
'            m.strIniFile = AddSlash(App.Path) & "Robbins.INI"
'            m.nBroker = eTT_AccountType_Robbins
'            m.strBrokerName = "Robbins"
'            m.strProcessName = "GenRobbins"
'            m.strProcessPath = AddSlash(App.Path) & "..\Brokers\Robbins\GenRobbins.EXE"
'            m.strDLLPath = AddSlash(App.Path) & "..\Brokers\Robbins\PatsAPI.DLL"
'            m.strControlID = "Robbins"
'            m.strLogPath = AddSlash(App.Path) & "Robbins"
'            strSymbolFile = AddSlash(App.Path) & "Provided\RobToGen.TXT"
        
        Case eTT_AccountType_Rosenthal
            m.strIniFile = AddSlash(App.Path) & "Rosenthal.INI"
            m.nBroker = eTT_AccountType_Rosenthal
            m.strBrokerName = "Rosenthal"
            m.strProcessName = "GenRosenthal"
            m.strProcessPath = AddSlash(App.Path) & "..\Brokers\Rosenthal\GenRosenthal.EXE"
            m.strDLLPath = AddSlash(App.Path) & "..\Brokers\Rosenthal\PatsAPI.DLL"
            m.strControlID = "Rosenthal"
            m.strLogPath = AddSlash(App.Path) & "Rosenthal"
            strSymbolFile = AddSlash(App.Path) & "Provided\RoseToGen.TXT"
    
    End Select
            
    If Not DirExist(m.strLogPath) Then MkDir m.strLogPath
    KillFile m.strLogPath & "\*.LOG /o=-30"
    
    Set m.BrokerInfo = New cBrokerInfo
    m.BrokerInfo.Init m.nBroker, m.strLogPath
    
    Set m.astrAccounts = New cGdArray
    m.astrAccounts.Create eGDARRAY_Strings
    
    Set m.astrBrokerPos = New cGdArray
    m.astrBrokerPos.Create eGDARRAY_Strings
    
    Set m.astrSymbols = New cGdArray
    m.astrSymbols.Create eGDARRAY_Strings
    m.astrSymbols.FromFile strSymbolFile
    m.astrSymbols.Sort
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.Init"

End Sub

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

    If (Len(m.strUserName) = 0) Or (Len(m.strIniPassword) = 0) Or (Len(m.strHostIP) = 0) Or (Len(m.strHostPort) = 0) Or (Len(m.strPriceIP) = 0) Or (Len(m.strPricePort) = 0) Then
        ShowConfig = True
    Else
        ShowConfig = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.ShowConfig"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Connect
'' Description: Attempt a connection to the PATS servers
''              if not already connected
'' Inputs:      New Account, Are we switching?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Connect(Optional ByVal strNewAccount As String = "", Optional ByVal bSwitching As Boolean = False)
On Error GoTo ErrSection:

    Dim bConnect As Boolean             ' Should we still try to connect?
    Dim strProcess As String            ' Name of the process to start
    Dim lTimeOut As Long                ' Timeout variable
    
    If g.bUnloading = False Then
        If (ConnectionStatus = eGDConnectionStatus_Disconnected) Then
            If (m.bAskedPassword = True) And ((strNewAccount = m.strUserName) Or (Len(strNewAccount) = 0)) Then
                bConnect = True
            Else
                bConnect = frmBrokerLogin.ShowMe(m.nBroker, strNewAccount)
                m.bAskedPassword = bConnect
            End If
        Else
            If ((strNewAccount <> m.strUserName) And (Len(strNewAccount) > 0)) Or (bSwitching = True) Then
                bConnect = frmBrokerLogin.ShowMe(m.nBroker, strNewAccount, False, bSwitching)
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
                
                If ((frmOnlineBroker.gdBroker.FindControlID(m.strControlID) = 0) Or (KillProcess(m.strProcessName, True) = 0)) Then
                    If KillProcess(m.strProcessName) > 0 Then
                        Sleep 3
                    End If
                        
                    FileCopy AddSlash(App.Path) & "..\Brokers\GenPATS.EXE", m.strProcessPath, True
                    FileCopy AddSlash(App.Path) & "..\Brokers\PatsAPI.DLL", m.strDLLPath, True
                    
                    DumpDebug "Trade Navigator version: " & FormatVersion & "." & Str(App.Revision) & " " & DateFormat(FileDate(App.Path & "\" & App.EXEName & ".EXE"), MM_DD_YYYY, HH_MM, AMPM_UPPER)
                    DumpDebug "Loading application " & m.strProcessPath
                    RunProcess m.strProcessPath, , , vbHide
                Else
                    ' UserName, Password, HostIP, HostPort, PriceIP, PricePort, AppID, AppLicense,
                    ' NewPassword, Environment, Version, SuperTAS
                    ' Environment = (G) ptGateway, (C) ptClient, (T) ptTestClient, (g) ptTestGateway, (D) ptDemoClient
                    DumpDebug "Connecting to " & m.strBrokerName & " as " & m.strUserName & " (" & m.strHostIP & ":" & m.strHostPort & ", " & m.strPriceIP & ":" & m.strPricePort & ")"
                    SendPatsMessage eGDPatsMessageType_Connect, m.strUserName & vbTab & m.strPassword & vbTab & m.strHostIP & vbTab & m.strHostPort & vbTab & m.strPriceIP & vbTab & m.strPricePort & vbTab & kAppID & vbTab & kAppLicense & vbTab & "" & vbTab & EnvChar & vbTab & kAppVersion & vbTab & m.strSuperTAS
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.Connect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disconnect
'' Description: Disconnect form the PATS servers if connected
'' Inputs:      Force Reconnect?, Reason for Disconnect
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Disconnect(Optional ByVal bForceReconnect As Boolean = False, Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    If (ConnectionStatus = eGDConnectionStatus_Connected) Then
        m.bManualDisconnect = Not bForceReconnect
        If m.bManualDisconnect Then
            m.dHeartBeat = 0#
            g.OrderStrategies.DeactivateExitsForBroker m.nBroker, "Disconnecting from " & m.strBrokerName
        End If
        
        If Len(strReason) = 0 Then DumpDebug "Disconnecting from " & m.strBrokerName Else DumpDebug "Disconnecting from " & m.strBrokerName & ": " & strReason
        SendPatsMessage eGDPatsMessageType_Disconnect, ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.Disconnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reconnect
'' Description: Attempt to reconnect to the PATS servers
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
    RaiseError "cPats.Reconnect"
    
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

    Order.Enter = OrderIsEntry(Order, False)

    If (Order.Status = eTT_OrderStatus_Open) Or (Order.Status = eTT_OrderStatus_Parked) Or (Order.Status = eTT_OrderStatus_TriggerPending) Or (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
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
    RaiseError "cIntBrokers.SendOrder"

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
    Dim strPatsSymbol As String         ' Pats Symbol
    Dim strPatsOrderType As String      ' Pats Order Type
    Dim dMult As Double                 ' Multiplier for converting prices
    Dim strExchange As String           ' Exchange
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    strPatsSymbol = PatsSymbol(Order.Symbol, dMult, strExchange)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strPatsOrderType = PatsOrderType(Order.OrderType)
    
    If Len(strPatsSymbol) = 0 Then
        DumpDebug "AddOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
    ElseIf Len(strPatsOrderType) = 0 Then
        DumpDebug "AddOrder: " & OrderType(Order.OrderType) & " cannot be converted to a " & m.strBrokerName & " order type"
        InfBox OrderType(Order.OrderType) & " is not a valid order type for " & m.strBrokerName, "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " cannot be converted to a " & m.strBrokerName & " order type"
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        DumpDebug "AddOrder: Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to " & m.strBrokerName & " account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "AddOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "AddOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    Else
        DumpDebug "AddOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
    
        astrOrder(0) = Order.GenesisOrderID
        astrOrder(1) = ""
        astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(3) = Parse(strPatsSymbol, vbTab, 1)
        astrOrder(4) = ""
        astrOrder(5) = strPatsOrderType
        If Order.Buy Then astrOrder(6) = "B" Else astrOrder(6) = "S"
        astrOrder(7) = Str(Order.Quantity)
        Select Case Order.OrderType
            Case eTT_OrderType_Market
                astrOrder(8) = ""
                astrOrder(9) = ""
            
            Case eTT_OrderType_Stop
                astrOrder(8) = PatsPrice(Order.StopPrice, dMult)
                astrOrder(9) = ""
            
            Case eTT_OrderType_Limit
                astrOrder(8) = PatsPrice(Order.LimitPrice, dMult)
                astrOrder(9) = ""
            
            Case eTT_OrderType_StopWithLimit
                astrOrder(8) = PatsPrice(Order.LimitPrice, dMult)
                astrOrder(9) = PatsPrice(Order.StopPrice, dMult)
                
        End Select
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = strExchange
        astrOrder(13) = ""
        astrOrder(14) = ""
        astrOrder(15) = Parse(strPatsSymbol, vbTab, 2)
        astrOrder(16) = ""
        If Order.Enter Then
            astrOrder(17) = "O"
        Else
            astrOrder(17) = "C"
        End If
    
        SendPatsMessage eGDPatsMessageType_AddOrder, astrOrder.JoinFields(vbTab)
        bReturn = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
    AddOrder = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.AddOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AmendOrder
'' Description: Attempt to amend the order
'' Inputs:      Order to amend
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AmendOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information
    Dim strPatsSymbol As String         ' Pats Symbol
    Dim strPatsOrderType As String      ' Pats Order Type
    Dim dMult As Double                 ' Multiplier for converting prices
    Dim strExchange As String           ' Exchange
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    strPatsSymbol = PatsSymbol(Order.Symbol, dMult, strExchange)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strPatsOrderType = PatsOrderType(Order.OrderType)
    
    If Len(strPatsSymbol) = 0 Then
        DumpDebug "AmendOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
    ElseIf Len(strPatsOrderType) = 0 Then
        DumpDebug "AmendOrder: " & OrderType(Order.OrderType) & " cannot be converted to a " & m.strBrokerName & " order type"
        InfBox OrderType(Order.OrderType) & " is not a valid order type for " & m.strBrokerName, "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " cannot be converted to a " & m.strBrokerName & " order type"
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        DumpDebug "AmendOrder: Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to " & m.strBrokerName & " account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "AmendOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "AmendOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    Else
        DumpDebug "AmendOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
    
        astrOrder(0) = Order.GenesisOrderID
        astrOrder(1) = Order.BrokerID
        astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(3) = ""
        astrOrder(4) = ""
        astrOrder(5) = ""
        astrOrder(6) = ""
        astrOrder(7) = Str(Order.Quantity)
        Select Case Order.OrderType
            Case eTT_OrderType_Market
                astrOrder(8) = ""
                astrOrder(9) = ""
            
            Case eTT_OrderType_Stop
                astrOrder(8) = PatsPrice(Order.StopPrice, dMult)
                astrOrder(9) = ""
            
            Case eTT_OrderType_Limit
                astrOrder(8) = PatsPrice(Order.LimitPrice, dMult)
                astrOrder(9) = ""
            
            Case eTT_OrderType_StopWithLimit
                astrOrder(8) = PatsPrice(Order.LimitPrice, dMult)
                astrOrder(9) = PatsPrice(Order.StopPrice, dMult)
                
        End Select
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = ""
        astrOrder(13) = ""
        astrOrder(14) = ""
        astrOrder(15) = ""
        astrOrder(16) = ""
        astrOrder(17) = ""
    
        SendPatsMessage eGDPatsMessageType_AmendOrder, astrOrder.JoinFields(vbTab)
        
        bReturn = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
    AmendOrder = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.AmendOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Attempt to cancel the order
'' Inputs:      Order to cancel
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information
    Dim strPatsSymbol As String         ' Pats Symbol
    Dim strPatsOrderType As String      ' Pats Order Type
    Dim dMult As Double                 ' Multiplier for converting prices
    Dim strExchange As String           ' Exchange
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    strPatsSymbol = PatsSymbol(Order.Symbol, dMult, strExchange)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strPatsOrderType = PatsOrderType(Order.OrderType)
    
    If Len(strPatsSymbol) = 0 Then
        DumpDebug "CancelOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
    ElseIf Len(strPatsOrderType) = 0 Then
        DumpDebug "CancelOrder: " & OrderType(Order.OrderType) & " cannot be converted to a " & m.strBrokerName & " order type"
        InfBox OrderType(Order.OrderType) & " is not a valid order type for " & m.strBrokerName, "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " cannot be converted to a " & m.strBrokerName & " order type"
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
        DumpDebug "CancelOrder: Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to " & m.strBrokerName & " account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "CancelOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "CancelOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    Else
        If Not g.TsoGroups Is Nothing Then
            g.TsoGroups.CancelRequested Order
        End If
        
        DumpDebug "CancelOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
    
        astrOrder(0) = Order.GenesisOrderID
        astrOrder(1) = Order.BrokerID
        astrOrder(2) = ""
        astrOrder(3) = ""
        astrOrder(4) = ""
        astrOrder(5) = ""
        astrOrder(6) = ""
        astrOrder(7) = ""
        astrOrder(8) = ""
        astrOrder(9) = ""
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = ""
        astrOrder(13) = ""
        astrOrder(14) = ""
        astrOrder(15) = ""
        astrOrder(16) = ""
        astrOrder(17) = ""
    
        SendPatsMessage eGDPatsMessageType_CancelOrder, astrOrder.JoinFields(vbTab)
        
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
    RaiseError "cPats.CancelOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Attempt to Park the order
'' Inputs:      Order to Park
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ParkOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information
    Dim strPatsSymbol As String         ' Pats Symbol
    Dim strPatsOrderType As String      ' Pats Order Type
    Dim dMult As Double                 ' Multiplier for converting prices
    Dim strExchange As String           ' Exchange
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    strPatsSymbol = PatsSymbol(Order.Symbol, dMult, strExchange)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strPatsOrderType = PatsOrderType(Order.OrderType)
    
    If Len(strPatsSymbol) = 0 Then
        DumpDebug "ParkOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
    ElseIf Len(strPatsOrderType) = 0 Then
        DumpDebug "ParkOrder: " & OrderType(Order.OrderType) & " cannot be converted to a " & m.strBrokerName & " order type"
        InfBox OrderType(Order.OrderType) & " is not a valid order type for " & m.strBrokerName, "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " cannot be converted to a " & m.strBrokerName & " order type"
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
        DumpDebug "ParkOrder: Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to " & m.strBrokerName & " account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "ParkOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "ParkOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    Else
        DumpDebug "ParkOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
    
        astrOrder(0) = Order.GenesisOrderID
        astrOrder(1) = Order.BrokerID
        astrOrder(2) = ""
        astrOrder(3) = ""
        astrOrder(4) = ""
        astrOrder(5) = ""
        astrOrder(6) = ""
        astrOrder(7) = ""
        astrOrder(8) = ""
        astrOrder(9) = ""
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = ""
        astrOrder(13) = ""
        astrOrder(14) = ""
        astrOrder(15) = ""
        astrOrder(16) = ""
        astrOrder(17) = ""
    
        SendPatsMessage eGDPatsMessageType_CancelOrder, astrOrder.JoinFields(vbTab)
        
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
    RaiseError "cPats.ParkOrder"
    
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
    RaiseError "cPats.CancelOrderFromBrokerID"
    
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
    RaiseError "cPats.CancelWorkingOrders"
    
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
    RaiseError "cPats.FlattenPosition"
    
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
        DumpDebug "Getting Account Information"
        m.bRefreshingAccounts = True
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for account information...", , , m.strBrokerName & " Refresh", True
        End If
        SendPatsMessage eGDPatsMessageType_GetAccounts, ""
    Else
        DumpDebug "Account Refresh Requested but already in Progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.GetAccounts"
    
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

    Dim lIndex As Long                  ' Index into a for loop

    If Not Refreshing Then
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for order information...", , , m.strBrokerName & " Refresh", True
        End If
        If Len(strAccount) = 0 Then
            DumpDebug "Getting Orders"
            m.bRefreshingOrders = True
            For lIndex = 0 To m.astrAccounts.Size - 1
                SendPatsMessage eGDPatsMessageType_GetOrders, m.astrAccounts(lIndex)
            Next lIndex
        Else
            DumpDebug "Getting Orders for " & strAccount
            m.bRefreshingOrders = True
            SendPatsMessage eGDPatsMessageType_GetOrders, strAccount
        End If
    Else
        DumpDebug "Order Refresh Requested but already in Progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.GetOrders"
    
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

    Dim lIndex As Long                  ' Index into a for loop

    If Not Refreshing Then
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for fill information...", , , m.strBrokerName & " Refresh", True
        End If
        If Len(strAccount) = 0 Then
            DumpDebug "Getting Fills"
            m.bRefreshingFills = True
            For lIndex = 0 To m.astrAccounts.Size - 1
                SendPatsMessage eGDPatsMessageType_GetFills, m.astrAccounts(lIndex)
            Next lIndex
        Else
            DumpDebug "Getting Fills for " & strAccount
            m.bRefreshingFills = True
            SendPatsMessage eGDPatsMessageType_GetFills, strAccount
        End If
    Else
        DumpDebug "Fill Refresh Requested but already in Progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.GetFills"
    
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

    Dim lIndex As Long                  ' Index into a for loop

    If Not Refreshing Then
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for position information...", , , m.strBrokerName & " Refresh", True
        End If
        If Len(strAccount) = 0 Then
            DumpDebug "Getting Positions"
            m.bRefreshingPositions = True
            For lIndex = 0 To m.astrAccounts.Size - 1
                SendPatsMessage eGDPatsMessageType_GetPositions, m.astrAccounts(lIndex)
            Next lIndex
        Else
            DumpDebug "Getting Positions for " & strAccount
            m.bRefreshingPositions = True
            SendPatsMessage eGDPatsMessageType_GetPositions, strAccount
        End If
    Else
        DumpDebug "Position Refresh Requested but already in Progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetContracts
'' Description: Ask the server for the available contracts for this user
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetContracts()
On Error GoTo ErrSection:

    DumpDebug "Getting Contracts"
    SendPatsMessage eGDPatsMessageType_GetContracts, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.GetContracts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Refresh
'' Description: Ask the server for accounts, orders, fills, and positions
'' Inputs:      Verbose Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Refresh(Optional ByVal bVerbose As Boolean = False)
On Error GoTo ErrSection:

    If (ConnectionStatus = eGDConnectionStatus_Connected) Then
        If Not Refreshing Then
            DumpDebug "Refreshing"
            
            m.bVerboseRefresh = bVerbose
            GetAccounts
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.Refresh"
    
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
    RaiseError "cPats.ConnectionStatusForAccount"
    
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
    RaiseError "cPats.NextGenesisID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendPatsMessage
'' Description: Send an App Mail message to the PATS program
'' Inputs:      Type of Message, Message, Send Now?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendPatsMessage(ByVal nType As eGDPatsMessageTypes, ByVal strMessage As String, Optional ByVal bSendNow As Boolean = True)
On Error GoTo ErrSection:

    Dim astrMessage As New cGdArray     ' Message split out into an array

    If nType = eGDPatsMessageType_Connect Then
        astrMessage.SplitFields strMessage, vbTab
        astrMessage(1) = SetStrToChar(astrMessage(1), "*")
        astrMessage(6) = SetStrToChar(astrMessage(6), "*")
        astrMessage(7) = SetStrToChar(astrMessage(7), "*")
        DumpDebug vbTab & "Sending " & m.strBrokerName & " Message (" & MessageTypeToString(nType) & "): " & astrMessage.JoinFields(vbTab)
    Else
        DumpDebug vbTab & "Sending " & m.strBrokerName & " Message (" & MessageTypeToString(nType) & "): " & strMessage
    End If
    
    frmOnlineBroker.gdBroker.CreateMessage m.strControlID, nType, strMessage, , bSendNow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.SendPatsMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming App Mail message from the PATS program
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDPatsMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    If (nType <> eGDPatsMessageType_ConnectionInfo) And (nType <> eGDPatsMessageType_Heartbeat) Then
        DumpDebug vbTab & m.strBrokerName & " Message Received (" & MessageTypeToString(nType) & "): " & strMessage
    End If
    
    Select Case nType
        Case eGDPatsMessageType_ConnectionInfo
            HandleConnectionInfo strMessage
        
        Case eGDPatsMessageType_AppLoaded
            m.bAppLoaded = True
            If LiveTradingAllowed(m.nBroker) Then
                DumpDebug "Connecting to " & m.strBrokerName & " as " & m.strUserName & " (" & m.strHostIP & ":" & m.strHostPort & ", " & m.strPriceIP & ":" & m.strPricePort & ")"
                SendPatsMessage eGDPatsMessageType_Connect, m.strUserName & vbTab & m.strPassword & vbTab & m.strHostIP & vbTab & m.strHostPort & vbTab & m.strPriceIP & vbTab & m.strPricePort & vbTab & kAppID & vbTab & kAppLicense & vbTab & "" & vbTab & EnvChar & vbTab & kAppVersion & vbTab & m.strSuperTAS
            Else
                DumpDebug "Unloading the application since live trading is not allowed"
                SendPatsMessage eGDPatsMessageType_UnloadApp, ""
            End If
            
        Case eGDPatsMessageType_AppUnloaded
            m.bAppLoaded = False
            If m.bManualDisconnect Then
                m.dHeartBeat = 0#
            Else
                m.dHeartBeat = gdTickCount - 110000
            End If
            m.bManualDisconnect = False
            ConnectionStatus = eGDConnectionStatus_Disconnected
        
        Case eGDPatsMessageType_Heartbeat
            m.dHeartBeat = gdTickCount
            If m.dLastFillTime > 0 Then
                If gdTickCount > m.dLastFillTime + 5000 Then
                    GetPositions
                    m.dLastFillTime = 0#
                End If
            End If
        
        Case eGDPatsMessageType_Order
            m.dHeartBeat = gdTickCount
            'UpdateBrokerInfoOrder strMessage, False
            HandleOrderInfo strMessage, False
        
        Case eGDPatsMessageType_AccountR
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                If m.bVerboseRefresh Then
                    InfBox "Refreshing account information...", , , m.strBrokerName & " Refresh", True
                End If
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
        
        Case eGDPatsMessageType_OrderR
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                If m.bVerboseRefresh Then
                    InfBox "Refreshing order information...", , , m.strBrokerName & " Refresh", True
                End If
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
        
        Case eGDPatsMessageType_FillR
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                If m.bVerboseRefresh Then
                    InfBox "Refreshing fill information...", , , m.strBrokerName & " Refresh", True
                End If
                m.BrokerInfo.RefreshingFills = True
                m.BrokerInfo.ClearFills
            ElseIf UCase(strMessage) = "END" Then
                m.BrokerInfo.RefreshingFills = False
                m.bRefreshingFills = False
                GetPositions
            Else
                'AddFillToBrokerInfo strMessage, True
                HandleOrderInfo strMessage, True
            End If
        
        Case eGDPatsMessageType_PositionR
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                If m.bVerboseRefresh Then
                    InfBox "Refreshing position information...", , , m.strBrokerName & " Refresh", True
                End If
                m.BrokerInfo.RefreshingPositions = True
                m.BrokerInfo.ClearPositions
                m.astrBrokerPos.Clear
            ElseIf UCase(strMessage) = "END" Then
                If m.bVerboseRefresh Then
                    InfBox "Synchronizing position information...", , , m.strBrokerName & " Refresh", True
                End If
                m.BrokerInfo.RefreshingPositions = False
                SynchronizePositions
                m.bRefreshingPositions = False
                If m.bVerboseRefresh Then
                    InfBox ""
                    m.bVerboseRefresh = False
                End If
            Else
                AddPositionToBrokerInfo strMessage, True
                HandlePositionInfo strMessage
            End If
    
        Case eGDPatsMessageType_ContractR
            m.dHeartBeat = gdTickCount
            frmTest2.AddList strMessage
        
        Case eGDPatsMessageType_QuoteR
            m.dHeartBeat = gdTickCount
    
    End Select
    
    If Refreshing And (m.dRefreshingStart = 0#) Then
        m.dRefreshingStart = gdTickCount
    ElseIf Refreshing And (gdTickCount > m.dRefreshingStart + 30000#) Then
        m.bRefreshingAccounts = False
        m.bRefreshingOrders = False
        m.bRefreshingFills = False
        m.bRefreshingPositions = False
        If m.bVerboseRefresh Then
            InfBox ""
            m.bVerboseRefresh = False
        End If
    ElseIf (Not Refreshing) And (m.dRefreshingStart <> 0#) Then
        m.dRefreshingStart = 0#
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.HandleMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckHeartbeat
'' Description: Check the heartbeat to make sure that we are still connected
''              to the PATS servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckHeartbeat()
On Error GoTo ErrSection:

    If (m.dHeartBeat <> 0) And (g.bUnloading = False) Then
        If (m.dHeartBeat < (gdTickCount - 120000)) Then
            If m.lReconnectAttempts = 1& Then
                InfBox "Communication has been lost with " & m.strBrokerName & ".  Attempting to reconnect...", , "+-OK", m.strBrokerName & " Connection", True
            End If
            Reconnect
        ElseIf (ConnectionStatus = eGDConnectionStatus_Connected) Then
            If m.lReconnectAttempts > 1& Then
                InfBox "TradeNavigator has reconnected to " & m.strBrokerName & ".", , "+-OK", m.strBrokerName & " Connection", True
            End If
            m.lReconnectAttempts = 0&
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.CheckHeartbeat"
    
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

    Dim lDate As Long                   ' Date as a long
    Dim strTime As String               ' Time
    Dim dTime As Double                 ' Time as a fractional part of day
    Dim dReturn As Double               ' Return value for the function
    
    dReturn = 0#
    If Len(strDateTime) > 0 Then
        lDate = JulFromLong(CLng(Val(Parse(strDateTime, " ", 1))))
        strTime = Parse(strDateTime, " ", 2)
        If Len(strTime) = 6 Then
            dTime = (Val(Left(strTime, 2)) / 24#) + (Val(Mid(strTime, 3, 2)) / 1440#) + (Val(Right(strTime, 2)) / 86400#)
        End If
        dReturn = CDbl(lDate) + dTime
    End If
    
    DateToDouble = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.DateToDouble"
    
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
    RaiseError "cPats.DumpDebug"
    
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

    If ((m.bVerifyPositions = True) Or (bForce = True)) And (m.bFixCalledOnce = True) Then
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
    RaiseError "cPats.VerifyPositions"
    
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

    IsTradeableSymbol = (Len(PatsSymbol(strGenesisSymbol)) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.IsTradeableSymbol"
    
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
        For lIndex = 0 To m.astrSymbols.Size - 1
            If Parse(m.astrSymbols(lIndex), vbTab, 3) = strBaseSymbol Then
                strMask = Parse(m.astrSymbols(lIndex), vbTab, 5)
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
    RaiseError "cPats.OrderTypeAllowed"
    
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
    For lIndex = 0 To m.astrSymbols.Size - 1
        If Parse(m.astrSymbols(lIndex), vbTab, 3) = strBaseSymbol Then
            strMask = Parse(m.astrSymbols(lIndex), vbTab, 6)
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
    RaiseError "cPats.TimeInForceAllowed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixPositions
'' Description: Allow the user to attempt to fix carried position mismatches
'' Inputs:      Only do first timers
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FixPositions(Optional ByVal bOnlyFirstTime As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrEntry As cGdArray           ' Entry in the array
    
    ' Only do this if there is no modal dialog up...
    If frmMain.Enabled Then
        If m.astrToFix.Size > 0 Then
            Set astrEntry = New cGdArray
            
            For lIndex = 0 To m.astrToFix.Size - 1
                astrEntry.SplitFields m.astrToFix(lIndex), vbTab
                If (bOnlyFirstTime = False) Or (astrEntry(2) = "0") Then
                    FixPosition astrEntry(0), astrEntry(1)
                End If
            Next lIndex
        End If
        
        If m.bFixCalledOnce = False Then m.bFixCalledOnce = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPATS.FixPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixPosition
'' Description: Allow the user to attempt to fix a carried position mismatch
'' Inputs:      Account, Symbol, Fill Summary
'' Returns:     True if Fixed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FixPosition(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, Optional ByVal FillSummary As cAccountPosition = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lBrokerPos As Long              ' Broker position
    Dim lBrokerCarried As Long          ' Broker carried position
    Dim strAccount As String            ' Account number
    Dim lPos As Long                    ' Position in the array
    Dim strSymbol As String             ' Symbol
    Dim frm As frmAlertPopup            ' Alert popup form

    bReturn = True
    If FillSummary Is Nothing Then
        Set FillSummary = g.Broker.FillSummary(vAccountNumberOrID, vSymbolOrSymbolID, -1&)
    End If
    
    If Not FillSummary Is Nothing Then
        strAccount = g.Broker.GetAccountNumber(vAccountNumberOrID)
        strSymbol = GetSymbol(vSymbolOrSymbolID)
        lBrokerPos = CurrentPosition(strAccount, FillSummary.Symbol)
        lBrokerCarried = CarriedPosition(strAccount, FillSummary.Symbol)
        
        If lBrokerPos <> FillSummary.CurrentPositionSnapshot Then
            ' Only go in here if there is no modal dialog up...
            If frmMain.Enabled Then
                DumpDebug "Symbol presented to user to fix: " & strAccount & ", " & strSymbol
                
                bReturn = frmPositionFix.ShowMe(FillSummary, lBrokerPos, lBrokerCarried)
                m.BrokerInfo.CarriedMatch(strAccount, FillSummary.Symbol) = bReturn
                
                If bReturn = True Then
                    If m.astrToFix.BinarySearch(strAccount & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
                        DumpDebug "Symbol removed from To Fix list: " & strAccount & ", " & strSymbol
                        m.astrToFix.Remove lPos
                    End If
                Else
                    DumpDebug "Symbol not fixed by user: " & strAccount & ", " & strSymbol
                    
                    If m.astrToFix.BinarySearch(strAccount & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
                        If Parse(m.astrToFix(lPos), vbTab, 3) = "0" Then
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox m.strBrokerName & " is reporting that you carried a " & UCase(g.Broker.TextPosition(lBrokerCarried)) & " position for " & strSymbol & " in account " & strAccount & ", but Trade Navigator reports a carried position of " & UCase(FillSummary.CurrentPositionString) & ".||Because this inconsistency could cause incorrect orders to be placed, auto exits and automated trading strategies are being disabled for this symbol.", "Inconsistent Carried Position", vbCenter
                            
                            g.OrderStrategies.DeactivateExit FillSummary.AccountID, FillSummary.SymbolOrSymbolID, , "Position not fixed by user"
                            g.TradingItems.DisableTradeItemsForSymbol FillSummary.AccountID, FillSummary.SymbolOrSymbolID, "Position mismatch", True
                        End If
                        
                        m.astrToFix(lPos) = strAccount & vbTab & strSymbol & vbTab & "1"
                    End If
                End If
            Else
                bReturn = False
            End If
        End If
    End If
    
    FixPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPATS.FixPosition"
    
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
    Dim nPrev As eGDConnectionStatus    ' Previous Connection status
    Dim strError As String              ' Error message back
    Dim frm As New frmAlertPopup        ' Alert popup form
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
                    
                    m.astrAccounts.Clear
                    
                    m.bRefreshingAccounts = False
                    m.bRefreshingOrders = False
                    m.bRefreshingFills = False
                    m.bRefreshingPositions = False
                
                    DumpDebug "Unloading Application since we have been disconnected"
                    SendPatsMessage eGDPatsMessageType_UnloadApp, ""
                End If
            
            Case eGDConnectionStatus_Disconnecting
                m.BrokerInfo.SetConnectionStatusForUser nStatus, strLastUser
            
            Case eGDConnectionStatus_Connecting
                m.BrokerInfo.SetConnectionStatusForUser nStatus, m.strUserName
            
            Case eGDConnectionStatus_Connected
                m.bWasConnected = True
        
        End Select
        
        If m.strUserName <> strLastUser Then
            strLastUser = m.strUserName
        End If
        
        If (nStatus <> ConnectionStatus) Or (Len(strError) > 0) Then
            nPrev = ConnectionStatus
            ConnectionStatus = nStatus
            
            g.ActivityLogs.AddStatusToActivityLog m.nBroker, nStatus
            
            If (nPrev <> eGDConnectionStatus_Connected) And (ConnectionStatus = eGDConnectionStatus_Connected) Then
                SetIniFileProperty "Last", EncryptToHex(Str(Date)), "Connect", m.strIniFile
                GetAccounts
            End If
            DumpDebug vbTab & m.strBrokerName & " Message Received (" & MessageTypeToString(eGDPatsMessageType_ConnectionInfo) & "): " & strConnectionInfo
        End If
        
        If (Len(strError) > 0) And (strError <> m.strLastMessage) Then
            If InStr(UCase(strError), "INVALID LOGIN") <> 0 Then
                m.bManualDisconnect = True
                m.dHeartBeat = 0#
                m.bVerifyPositions = True
                m.bAskedPassword = False
            End If
            
            Set frm = New frmAlertPopup
            frm.ShowMessageBox strError, m.strBrokerName & " Message", vbCenter
        End If
        m.strLastMessage = strError
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.HandleConnectionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleAccountInfo
'' Description: Do the appropriate thing with the given account information
'' Inputs:      Account Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleAccountInfo(ByVal strAccountInfo As String)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strAccountNumber As String      ' Account Number from the callback
    Dim lAccountID As Long              ' Account ID
    Dim lPos As Long                    ' Position of account in the array

    If Len(Trim(strAccountInfo)) > 0 Then
        strAccountNumber = Parse(strAccountInfo, vbTab, 1)
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
            rs!SecTypeMask = 1
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
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.HandleAccountInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleOrderInfo
'' Description: Do the appropriate thing with the given order information
'' Inputs:      Order Information
'' Returns:     None
''
'' RecordFrom:  Genesis ID, Order ID, Fill ID, Account, Base Symbol, Symbol,
''              Order Type, B/S, Quantity, Limit, Stop, Strike, Expiration,
''              Exchange, Instrument, Time In Force, Maturity, Reserved,
''              Quantity Filled, Remaining Quantity, Quantity Cum, Fill Price,
''              Fill Time, Event, Status, Message, Position ID, Open/Close
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
    Dim bFillChanged As Boolean         ' Did the fill change?
    Dim nPrevStatus As eTT_OrderStatus  ' Previous order status
    Dim lIndex As Long                  ' Index into a for loop
    Dim lOldOrderID As Long             ' Order ID for the previous version of this order
    Dim OldOrder As cPtOrder            ' Old Order
    Dim dFillTime As Double             ' Old Fill Time
    Dim dFillPrice As Double            ' Old Fill Price
    Dim lFillQty As Long                ' Old Fill Quantity
    Dim Bars As New cGdBars             ' Bars object
    Dim frm As frmAlertPopup            ' Alert popup form
    Dim nOrderType As eTT_OrderType     ' Order Type
    
    If Len(strOrderInfo) > 0 Then
        astrOrder.SplitFields strOrderInfo, vbTab
        
        strGenesisSymbol = GenesisSymbol(astrOrder(4), astrOrder(16), dMult)
        strAccountNumber = astrOrder(3)
        
        If Len(strGenesisSymbol) > 0 Then
            SetBarProperties Bars, strGenesisSymbol
            Set Order = LoadOrder(strAccountNumber, astrOrder(0), astrOrder(1), lOldOrderID)
            If Order Is Nothing Then
                DumpDebug "Order " & astrOrder(1) & ";" & astrOrder(0) & " not found -- creating new order"
                Set Order = New cPtOrder
            Else
                DumpDebug "Order " & astrOrder(1) & ";" & astrOrder(0) & " found"
            End If
            
            dFillTime = DateToDouble(astrOrder(22))
            
            With Order
                nPrevStatus = .Status
                
                If .OrderID = 0 Then .GenesisOrderID = astrOrder(0)
                .BrokerID = astrOrder(1)
                .AccountID = g.Broker.AccountIDForNumber(strAccountNumber)
                
                .Buy = (UCase(astrOrder(7)) = "B")
                .Quantity = CLng(Val(astrOrder(8)))
                .SymbolOrSymbolID = strGenesisSymbol
                
                nOrderType = GenesisOrderType(astrOrder(6))
                If Not ((nOrderType = eTT_OrderType_Market) And (.OrderType = eTT_OrderType_MIT)) Then
                    .OrderType = nOrderType
                    Select Case .OrderType
                        Case eTT_OrderType_Market
                            .LimitPrice = 0#
                            .StopPrice = 0#
                        Case eTT_OrderType_Stop
                            .LimitPrice = GenesisPrice(astrOrder(10), dMult)
                            .StopPrice = GenesisPrice(astrOrder(9), dMult)
                        Case eTT_OrderType_Limit
                            .LimitPrice = GenesisPrice(astrOrder(9), dMult)
                            .StopPrice = GenesisPrice(astrOrder(10), dMult)
                        Case eTT_OrderType_StopWithLimit
                            .LimitPrice = GenesisPrice(astrOrder(9), dMult)
                            .StopPrice = GenesisPrice(astrOrder(10), dMult)
                    End Select
                End If
                
                If .OrderDate = 0# Then .OrderDate = dFillTime
                .Message = astrOrder(25)
                .Expiration = -1&
                
                If (UCase(astrOrder(27)) = "O") Or (UCase(astrOrder(27)) = "OPEN") Then
                    .Enter = True
                ElseIf (UCase(astrOrder(27)) = "C") Or (UCase(astrOrder(27)) = "CLOSE") Then
                    .Enter = False
                End If
                
                .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.OrderDate, m.nBroker, .Symbol, False))
                .IsSnapshot = True
                
                ' If there is no order ID yet, then we need to save to that we have an order ID
                ' to send to the order ID changed function (if in an amend situation)...
                If .OrderID = 0& Then
                    .Save
                End If
                    
                If lOldOrderID <> 0& Then
                    Set OldOrder = New cPtOrder
                    If OldOrder.Load(lOldOrderID) Then
                        OldOrder.GenesisOrderID = ""
                        OldOrder.Save
                        OrderCallback OldOrder
                    End If
                    
                    .PreviousBrokerID = OldOrder.BrokerID
                    .PreviousOrder = OldOrder
                    
                    DumpDebug "Order ID for " & .GenesisOrderID & " changed from " & Str(lOldOrderID) & " to " & Str(.OrderID)
                    OrderIDChanged lOldOrderID, .OrderID
                End If
                
                ' We have a fill if there is a fill ID and the quantity exectued is greater than zero...
                If (Len(astrOrder(2)) > 0) And (Val(astrOrder(18)) > 0) Then
                    bFillReceived = True
                    bFillExisted = False
                    bFillChanged = True
                    
                    dFillPrice = GenesisPrice(astrOrder(21), dMult)
                    lFillQty = CLng(Val(astrOrder(18)))
                    
                    For lIndex = 1 To .Fills.Count
                        If .Fills(lIndex).BrokerID = astrOrder(2) Then
                            Set Fill = .Fills(lIndex)
                            bFillExisted = True
                            
                            bFillChanged = (Fill.FillDate <> dFillTime) Or (Fill.Price <> dFillPrice) Or (Fill.Quantity <> lFillQty)
                            
                            Exit For
                        End If
                    Next lIndex
                    
                    If Fill Is Nothing Then Set Fill = New cPtFill
                    With Fill
                        .AccountID = Order.AccountID
                        .AutoTradingItemID = Order.AutoTradeItemID
                        .BrokerID = astrOrder(2)
                        .BrokerOrderID = Order.BrokerID
                        .Buy = Order.Buy
                        .FillDate = dFillTime
                        .IsManual = False
                        .IsSnapshot = True
                        .OrderID = Order.OrderID
                        .Price = dFillPrice
                        .Quantity = lFillQty
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
                .StatusDate = dFillTime
                If (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_ParkPending) Then
                    .Status = eTT_OrderStatus_Parked
                ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_Parked) Then
                    .Status = eTT_OrderStatus_Parked
                ElseIf (Order.FillQuantity > Order.Quantity) Then
                    .Status = eTT_OrderStatus_OverFilled
                ElseIf (.Status = eTT_OrderStatus_Filled) And (Order.FillQuantity < Order.Quantity) Then
                    .Status = eTT_OrderStatus_Partial
                ElseIf (.Status = eTT_OrderStatus_Working) And (Order.FillQuantity = Order.Quantity) Then
                    .Status = eTT_OrderStatus_Filled
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
                
                    FillCallback Order, Fill, bFillExisted, True, bFillChanged
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
                
                If (Len(astrOrder(25)) > 0) And (UCase(astrOrder(23)) <> "REWIND") Then
                    Set frm = New frmAlertPopup
                    frm.ShowMessageBox Order.OrderText & " REJECTED||" & astrOrder(25), m.strBrokerName & " Order " & Order.BrokerID & " Rejected", vbLeftJustify
                End If
                
                If (.Status = eTT_OrderStatus_OverFilled) And (UCase(astrOrder(23)) = "REWIND") Then
                    Set frm = New frmAlertPopup
                    frm.ShowMessageBox "Trade Navigator has received duplicate fill|information for order " & Order.BrokerID & ".||This will cause this order to appear over-filled and will likely result in a position mismatch.||PLEASE CALL YOUR BROKER AND VERIFY YOUR POSITIONS.", "Over Filled Order on " & m.strBrokerName, vbCenter
                End If
                
                DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
            End With
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.HandleOrderInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandlePositionInfo
'' Description: Do the appropriate thing with the given Position information
'' Inputs:      Position Information
'' Returns:     None
''
'' RecordFrom:  Account, Symbol, Symbol Long, Position, Realized, Unrealized,
''              Avg Entry, Carried Position, Carried Realized, Carried Unrealized,
''              Carried Avg Entry
'' RecordTo:    Account, Symbol, Current Position, Carried Position, Average
''              Entry, Entry Prices
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandlePositionInfo(ByVal strPositionInfo As String)
On Error GoTo ErrSection:

    Dim astrPosition As New cGdArray    ' Array of position information
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim lPos As Long                    ' Position in the array
    Dim astrBrokerPos As cGdArray       ' Broker position split out into an array
    
    If Len(strPositionInfo) > 0 Then
        astrPosition.SplitFields strPositionInfo, vbTab
        
        strGenesisSymbol = GenesisSymbol(Parse(astrPosition(2), "-", 1), Parse(astrPosition(2), "-", 2))
        
        If Len(strGenesisSymbol) > 0 Then
            If m.astrBrokerPos.BinarySearch(astrPosition(0) & vbTab & strGenesisSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                Set astrBrokerPos = New cGdArray
                astrBrokerPos.SplitFields m.astrBrokerPos(lPos), vbTab
                m.astrBrokerPos(lPos) = astrPosition(0) & vbTab & strGenesisSymbol & vbTab & astrPosition(3) & vbTab & astrBrokerPos(3) & vbTab & astrBrokerPos(4) & vbTab & astrBrokerPos(5)
            Else
                m.astrBrokerPos.Add astrPosition(0) & vbTab & strGenesisSymbol & vbTab & astrPosition(3) & vbTab & "" & vbTab & "" & vbTab & "", lPos
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.HandlePositionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SynchronizePositions
'' Description: Synchronize the positions between Genesis and Pats
'' Inputs:      Symbol
'' Returns:     None
''
'' Position:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SynchronizePositions(Optional ByVal strSymbol As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPosition As Long               ' Position from the broker
    Dim strPosition As String           ' Position information
    Dim frm As frmAlertPopup            ' Alert popup form
    Dim FillSumms As cAccountPositions  ' Collection of fill summaries
    Dim strAccount As String            ' Account number
    Dim lBrokerPos As Long              ' Broker position
    Dim lBrokerCarried As Long          ' Broker carried position
    Dim lPos As Long                    ' Position in an array
    
    m.bSyncInProgress = True
    
    BuildPositions
    frmPositionConfirm.ShowMe m.nBroker, "", "", False, False
    
    Set FillSumms = m.BrokerInfo.FillSummaries(True)
    If Not FillSumms Is Nothing Then
        For lIndex = 1 To FillSumms.Count
            With FillSumms(lIndex)
                If (.AutoTradeItemID = -1&) And ((Len(strSymbol) = 0) Or (strSymbol = .Symbol)) Then
                    g.OrderStrategies.RefreshPosition .AccountID, .SymbolOrSymbolID
                    
                    strAccount = g.Broker.AccountNumberForID(.AccountID)
                    
                    ' Retrieve the broker position and broker carried position from the array...
                    lBrokerPos = CurrentPosition(strAccount, .Symbol)
                    lBrokerCarried = CarriedPosition(strAccount, .Symbol)
                    lPosition = lBrokerCarried + .NumBuysSnapshot - .NumSellsSnapshot
                    
                    ' Verify first that broker information is consistent with itself...
                    DumpDebug "Position Verification(" & strAccount & ", " & .Symbol & "): " & Str(lBrokerCarried) & " + " & Str(.NumBuysSnapshot) & " - " & Str(.NumSellsSnapshot) & " = " & Str(lBrokerPos)
                    
                    m.BrokerInfo.ConsistentBroker(strAccount, .Symbol) = True
                        
                    ' Now verify that the carried position information matches...
                    If lBrokerCarried <> .CurrentPosition Then
                        m.BrokerInfo.CarriedMatch(strAccount, .Symbol) = False
                        
                        If m.astrToFix.BinarySearch(strAccount & vbTab & .Symbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = False Then
                            DumpDebug "Symbol added to To Fix List: " & strAccount & ", " & .Symbol
                            m.astrToFix.Add strAccount & vbTab & .Symbol & vbTab & "0", lPos
                        End If
                    Else
                        m.BrokerInfo.CarriedMatch(strAccount, .Symbol) = True
                        
                        If m.astrToFix.BinarySearch(strAccount & vbTab & .Symbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
                            DumpDebug "Symbol removed from To Fix list: " & strAccount & ", " & .Symbol
                            m.astrToFix.Remove lPos
                        End If
                    End If
                End If
            End With
        Next lIndex
    End If

    m.bSyncInProgress = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.SynchronizePositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildPositions
'' Description: Build the positions from the position and fill summary info
'' Inputs:      None
'' Returns:     None
''
'' RecordFrom:  Account, Symbol, Current Position, Carried Position, Average
''              Entry, Entry Prices
'' Position:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildPositions()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrInfo As cGdArray            ' Information split out into arrays
    Dim astrToAdd As cGdArray           ' Array to add
    Dim lPos As Long                    ' Position in an array
    Dim strOvernight As String          ' Overnight position
    Dim lPosition As Long               ' Position
    Dim FillSumms As cAccountPositions  ' Fill summary collection
    Dim strAccount As String            ' Account number
    Dim astrBrokerPos As cGdArray       ' Broker position entry
    Dim lCurrentPos As Long             ' Current position
    
    If Not Refreshing Then
        Set astrInfo = New cGdArray
        Set astrToAdd = New cGdArray
        Set astrBrokerPos = New cGdArray
        
        ' Clear out the broker info positions...
        m.BrokerInfo.RefreshingPositions = True
        m.BrokerInfo.ClearPositions
        
        ' Start with the overnight positions...
        For lIndex = 0 To m.astrBrokerPos.Size - 1
            astrInfo.Clear
            astrInfo.SplitFields m.astrBrokerPos(lIndex)
            
            astrToAdd.Clear
            astrToAdd(0) = astrInfo(0)
            astrToAdd(1) = astrInfo(1)
            astrToAdd(2) = astrInfo(2)
            astrToAdd(3) = astrInfo(4)
            astrToAdd(4) = astrInfo(3)
            
            DumpDebug "BuildPositions - Position: " & astrToAdd.JoinFields(vbTab)
            m.BrokerInfo.AddPosition astrToAdd.JoinFields(vbTab), False
        Next lIndex
        
        ' Now walk through the broker summary information...
        Set FillSumms = m.BrokerInfo.FillSummaries(True)
        If Not FillSumms Is Nothing Then
            For lIndex = 1 To FillSumms.Count
                With FillSumms(lIndex)
                    If .AutoTradeItemID = -1& Then
                        strAccount = g.Broker.AccountNumberForID(.AccountID)
                        
                        If m.astrBrokerPos.BinarySearch(strAccount & vbTab & .Symbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
                            astrBrokerPos.SplitFields m.astrBrokerPos(lPos)
                            lCurrentPos = CLng(Val(astrBrokerPos(2)))
                            lPosition = lCurrentPos - .NumBuysSnapshot + .NumSellsSnapshot
                            strOvernight = Str(lPosition)
                            astrBrokerPos(4) = strOvernight
                            m.astrBrokerPos(lPos) = astrBrokerPos.JoinFields(vbTab)
                        Else
                            lCurrentPos = 0&
                            lPosition = .NumSellsSnapshot - .NumBuysSnapshot
                            strOvernight = Str(lPosition)
                            m.astrBrokerPos.Add strAccount & vbTab & .Symbol & vbTab & Str(lCurrentPos) & vbTab & Str(lPosition) & vbTab & "" & vbTab & ""
                        End If
                        
                        astrToAdd.Clear
                        astrToAdd(0) = strAccount
                        astrToAdd(1) = .Symbol
                        astrToAdd(2) = Str(lCurrentPos)
                        astrToAdd(3) = ""
                        astrToAdd(4) = strOvernight
                        
                        m.BrokerInfo.AddPosition astrToAdd.JoinFields(vbTab), True
                    End If
                End With
            Next lIndex
        End If
    
        m.BrokerInfo.RefreshingPositions = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPATS.BuildPositions"
    
End Sub

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
    
    ' First, try to look it up from the Trade Console in case it is already loaded...
    Set Order = m.BrokerInfo.Order(strBrokerID, strGenesisOrderID)
    If Not Order Is Nothing Then
        If (Len(strBrokerID) = 0) Or (Len(Order.BrokerID) = 0) Or (Order.BrokerID = strBrokerID) Then
            bFound = True
        Else
            bFound = False
            lOldOrderID = Order.OrderID
        End If
    End If
    
    ' If not found, we will need to attempt to load it from the database...
    If bFound = False Then
        ' First, try to load it from the database with the broker's order ID...
        If Len(strBrokerID) > 0 Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                        "WHERE [BrokerOrderID]='" & strBrokerID & "' AND [AccountID]=" & lAccountID & ";", dbOpenDynaset)
            If Not (rs.EOF And rs.BOF) Then
                Set Order = New cPtOrder
                If Order.Load(rs!OrderID) Then
                    bFound = True
                    lOldOrderID = 0&
                End If
            End If
        End If
        
        ' If not found, then try to look it up with the Genesis Order ID...
        If (bFound = False) And (Len(strGenesisOrderID) > 0) Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                        "WHERE [GenesisOrderID]='" & strGenesisOrderID & "' AND [AccountID]=" & lAccountID & ";", dbOpenDynaset)
            If Not (rs.EOF And rs.BOF) Then
                Set Order = New cPtOrder
                If Order.Load(rs!OrderID) Then
                    If (Len(strBrokerID) = 0) Or (Len(Order.BrokerID) = 0) Or (Order.BrokerID = strBrokerID) Then
                        bFound = True
                        lOldOrderID = 0&
                    Else
                        bFound = False
                        lOldOrderID = Order.OrderID
                    End If
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
    RaiseError "cPats.LoadOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PatsSymbol
'' Description: Convert a Genesis symbol to a Pats symbol
'' Inputs:      Genesis symbol, Multiplier, Exchange
'' Returns:     Pats Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PatsSymbol(ByVal strGenesisSymbol As String, Optional dMult As Double, Optional strExchange As String) As String
On Error GoTo ErrSection:

    Dim strBase As String               ' Base symbol
    Dim strContract As String           ' Contract
    Dim strReturn As String             ' Return from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim strYear As String               ' Year of the contract
    Dim iMonth As Integer               ' Month of the contract
    
    strReturn = ""
    If Len(strGenesisSymbol) > 0 Then
        strBase = Parse(strGenesisSymbol, "-", 1)
        strContract = Parse(strGenesisSymbol, "-", 2)
        
        For lIndex = 0 To m.astrSymbols.Size - 1
            If Parse(m.astrSymbols(lIndex), vbTab, 3) = strBase Then
                strYear = Left(strContract, 4)
                iMonth = CInt(Val(Right(strContract, 2)))
                
                strReturn = Parse(m.astrSymbols(lIndex), vbTab, 1) & vbTab & UCase(MonthName(iMonth, True, True)) & Right(strYear, 2)
                dMult = Val(Parse(m.astrSymbols(lIndex), vbTab, 4))
                strExchange = Parse(m.astrSymbols(lIndex), vbTab, 2)
                
                Exit For
            End If
        Next lIndex
    End If

    PatsSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.PatsSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Convert a Pats symbol to a Genesis symbol
'' Inputs:      Pats symbol, Exchange, Contract, Multiplier
'' Returns:     Genesis Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisSymbol(ByVal strPatsSymbol As String, ByVal strContract As String, Optional dMult As Double) As String
On Error GoTo ErrSection:

    Dim strBase As String               ' Base symbol
    Dim lPos As Long                    ' Position of base symbol in the array
    Dim strReturn As String             ' Return from the function
    Dim lYear As Long                   ' Year of the contract
    Dim lMonth As Long                  ' Month of the contract
            
    strReturn = ""
    If Len(strPatsSymbol) > 0 Then
        If m.astrSymbols.BinarySearch(strPatsSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
            lYear = CLng(Val(Right(strContract, 2))) + 2000
            lMonth = MonthFromName(Left(strContract, 3))
            Do While (lYear * 100 + lMonth) < (Year(Date) * 100 + Month(Date))
                lYear = lYear + 10
            Loop
        
            strReturn = Parse(m.astrSymbols(lPos), vbTab, 3) & "-" & Str(lYear) & Format(lMonth, "00")
            dMult = Val(Parse(m.astrSymbols(lPos), vbTab, 4))
        End If
    End If

    GenesisSymbol = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.GenesisSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PatsOrderType
'' Description: Convert a Genesis order type to a Pats order type
'' Inputs:      Genesis order type
'' Returns:     Pats Order Type (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PatsOrderType(ByVal nOrderType As eTT_OrderType) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case nOrderType
        Case eTT_OrderType_Market
            strReturn = "Market"
        Case eTT_OrderType_Stop
            strReturn = "Stop" ' "Stop!"
        Case eTT_OrderType_Limit
            strReturn = "Limit"
        Case eTT_OrderType_StopWithLimit
            strReturn = "StopLimit"
        Case Else
            strReturn = ""
    End Select
    
    PatsOrderType = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.PatsOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderType
'' Description: Convert a Pats order type to a Genesis order type
'' Inputs:      Pats order type
'' Returns:     Genesis Order Type (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderType(ByVal strPatsOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderType        ' Return value for the function

    Select Case UCase(strPatsOrderType)
        Case "MARKET"
            nReturn = eTT_OrderType_Market
        Case "LIMIT"
            nReturn = eTT_OrderType_Limit
        Case "STOP!", "STOP"
            nReturn = eTT_OrderType_Stop
        Case "STOPLIMIT!", "STOPLIMIT"
            nReturn = eTT_OrderType_StopWithLimit
    End Select
    
    GenesisOrderType = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.GenesisOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PatsPrice
'' Description: Convert a Genesis price to a Pats price
'' Inputs:      Genesis Price, Multiplier
'' Returns:     Pats Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PatsPrice(ByVal dGenesisPrice As Double, ByVal dMult As Double) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from the function
    
    If dMult <> 0 Then
        Select Case dMult
            Case 8
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 8, "0")
            
            Case 32
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 32, "00")
            
            Case 32.25
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 3200, "0000")
                ' 06/09/2009 DAJ: Customer is getting rejected because we are doing this, so I am going to
                ' comment it out for now.
                'strReturn = Left(strReturn, Len(strReturn) - 1)
            
            Case 32.5
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 320, "000")
            
            Case Else
                strReturn = Str(dGenesisPrice / dMult)

        End Select
    End If
    
    PatsPrice = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.PatsPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisPrice
'' Description: Convert a Pats price to a Genesis price
'' Inputs:      Pats Price, Multiplier
'' Returns:     Genesis Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisPrice(ByVal strPatsPrice As String, ByVal dMult As Double) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value from the function
    Dim strWhole As String              ' Whole part of the price
    Dim strFraction As String           ' Fractional part of the price
    
    If strPatsPrice = "0" Then
        dReturn = 0#
    ElseIf (dMult <> 0) And (Len(strPatsPrice) > 0) Then
        strWhole = Parse(strPatsPrice, ".", 1)
        strFraction = Parse(strPatsPrice, ".", 2)
        
        Select Case dMult
            Case 8
                If Len(strFraction) = 0 Then
                    strFraction = "0"
                End If
                dReturn = Val(strWhole) + (Val(strFraction) / 8)
            
            Case 32
                If Len(strFraction) = 0 Then
                    strFraction = "00"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "0"
                End If
                dReturn = Val(strWhole) + (Val(strFraction) / 32)
            
            Case 32.25
                If Len(strFraction) = 0 Then
                    strFraction = "0000"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "000"
                ElseIf Len(strFraction) = 2 Then
                    strFraction = strFraction & "00"
                ElseIf Len(strFraction) = 3 Then
                    If Right(strFraction, 1) = "2" Or Right(strFraction, 1) = "7" Then
                        strFraction = strFraction & "5"
                    Else
                        strFraction = strFraction & "0"
                    End If
                End If
                dReturn = Val(strWhole) + ((Val(strFraction) / 100) / 32)
            
            Case 32.5
                If Len(strFraction) = 0 Then
                    strFraction = "000"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "00"
                ElseIf Len(strFraction) = 2 Then
                    strFraction = strFraction & "0"
                End If
                dReturn = Val(strWhole) + ((Val(strFraction) / 10) / 32)
                
            Case 64
                If Len(strFraction) = 0 Then
                    strFraction = "00"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "0"
                End If
                dReturn = Val(strWhole) + (Val(strFraction) / 64)
            
            Case 64.5
                If Len(strFraction) = 0 Then
                    strFraction = "000"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "00"
                ElseIf Len(strFraction) = 2 Then
                    strFraction = strFraction & "0"
                End If
                dReturn = Val(strWhole) + ((Val(strFraction) / 10) / 64)
            
            Case Else
                dReturn = Val(strPatsPrice) * dMult
                
        End Select
    End If
    
    GenesisPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.GenesisPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderStatus
'' Description: Convert a Pats order status to a Genesis order status
'' Inputs:      Pats order status, Event for Callback
'' Returns:     Genesis order status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderStatus(ByVal strPatsOrderStatus As String, ByVal strPatsEvent As String) As eTT_OrderStatus
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderStatus      ' Return value from the function

    Select Case UCase(strPatsOrderStatus)
        Case "PENDING"
            nReturn = eTT_OrderStatus_Working
        Case "FILLED"
            nReturn = eTT_OrderStatus_Filled
        Case "BALCANCEL"
            If UCase(strPatsEvent) = "REPLACE" Then
                nReturn = eTT_OrderStatus_Amended
            Else
                nReturn = eTT_OrderStatus_BalCancelled
            End If
        Case "CANCELED"
            If UCase(strPatsEvent) = "REPLACE" Then
                nReturn = eTT_OrderStatus_Amended
            Else
                nReturn = eTT_OrderStatus_Cancelled
            End If
        Case "REJECTED"
            nReturn = eTT_OrderStatus_Rejected
    End Select
    
    GenesisOrderStatus = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.GenesisOrderStatus"
    
End Function

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
    astrToAdd(4) = GenesisSymbol(astrOrder(4), astrOrder(16), dMult)
    astrToAdd(5) = Str(GenesisOrderType(astrOrder(6)))
    If UCase(astrOrder(7)) = "B" Then
        astrToAdd(6) = Str(CLng(True))
    Else
        astrToAdd(6) = Str(CLng(False))
    End If
    astrToAdd(7) = astrOrder(8)
    astrToAdd(8) = Str(GenesisPrice(astrOrder(9), dMult))
    astrToAdd(9) = Str(GenesisPrice(astrOrder(10), dMult))
    astrToAdd(10) = "-1"
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
    RaiseError "cPats.UpdateBrokerInfoOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderToBrokerInfo
'' Description: Add the Genesis order string to the broker info class
'' Inputs:      Order String, Status
'' Returns:     None
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
    RaiseError "cPats.GenesisOrderToBrokerInfo"
    
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
    astrToAdd(4) = GenesisSymbol(astrFill(4), astrFill(16), dMult)
    astrToAdd(5) = Str(DateToDouble(astrFill(22)))
    If (UCase(astrFill(7)) = "B") Then
        astrToAdd(6) = Str(CLng(True))
    Else
        astrToAdd(6) = Str(CLng(False))
    End If
    astrToAdd(7) = astrFill(18)
    astrToAdd(8) = Str(GenesisPrice(astrFill(21), dMult))
    
    m.BrokerInfo.AddFill astrToAdd.JoinFields(vbTab), bRefresh

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.AddFillToBrokerInfo"
    
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
    Dim strPatsSymbol As String         ' Pats Symbol
    Dim dMult As Double                 ' Multiplier for converting prices
    Dim strExchange As String           ' Exchange

    If Len(strOrder) > 0 Then
        Set astrOrder = New cGdArray
        astrOrder.Create eGDARRAY_Strings
        Set astrToCancel = New cGdArray
        astrToCancel.Create eGDARRAY_Strings
    
        astrOrder.SplitFields strOrder, vbTab
        strPatsSymbol = PatsSymbol(astrOrder(4), dMult, strExchange)
        If Len(strPatsSymbol) > 0 Then
            astrToCancel(0) = astrOrder(1)
            astrToCancel(1) = astrOrder(0)
            astrToCancel(2) = ""
            astrToCancel(3) = ""
            astrToCancel(4) = ""
            astrToCancel(5) = ""
            astrToCancel(6) = ""
            astrToCancel(7) = ""
            astrToCancel(8) = ""
            astrToCancel(9) = ""
            astrToCancel(10) = ""
            astrToCancel(11) = ""
            astrToCancel(12) = ""
            astrToCancel(13) = ""
            astrToCancel(14) = ""
            astrToCancel(15) = ""
            astrToCancel(16) = ""
            
            SendPatsMessage eGDPatsMessageType_CancelOrder, astrToCancel.JoinFields(vbTab)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.CancelBrokerInfoOrder"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddPositionToBrokerInfo
'' Description: Add the position string to the broker info class
'' Inputs:      Position String, Refresh?
'' Returns:     None
''
'' RecordFrom:  Account, Symbol, Symbol Long, Position, Realized,
''              Unrealized, Avg Entry, Carried Position, Carried Realized,
''              Carried Unrealized, Carried Avg Entry
'' RecordTo:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddPositionToBrokerInfo(ByVal strPositionString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim astrPosition As cGdArray        ' Array of position information split out from string
    Dim astrToAdd As cGdArray           ' Array of information to pass to broker info class
    Dim dMult As Double                 ' Price Multiplier
    
    Set astrPosition = New cGdArray
    astrPosition.Create eGDARRAY_Strings
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings
    
    astrPosition.SplitFields strPositionString, vbTab
    
    astrToAdd(0) = astrPosition(0)
    astrToAdd(1) = GenesisSymbol(Parse(astrPosition(1), "-", 1), Parse(astrPosition(1), "-", 2), dMult)
    astrToAdd(2) = astrPosition(3)
    astrToAdd(3) = GenesisPrice(astrPosition(6), dMult)
    astrToAdd(4) = astrPosition(7)
    
    m.BrokerInfo.AddPosition astrToAdd.JoinFields(vbTab), bRefresh

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.AddPositionToBrokerInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MonthFromName
'' Description: Convert a month abbreviation to the month number
'' Inputs:      Month Name
'' Returns:     Month Number
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MonthFromName(ByVal strMonthName As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value from the function

    Select Case strMonthName
        Case "JAN"
            lReturn = 1
        Case "FEB"
            lReturn = 2
        Case "MAR"
            lReturn = 3
        Case "APR"
            lReturn = 4
        Case "MAY"
            lReturn = 5
        Case "JUN"
            lReturn = 6
        Case "JUL"
            lReturn = 7
        Case "AUG"
            lReturn = 8
        Case "SEP"
            lReturn = 9
        Case "OCT"
            lReturn = 10
        Case "NOV"
            lReturn = 11
        Case "DEC"
            lReturn = 12
    End Select
    
    MonthFromName = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.MonthFromName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnvChar
'' Description: Return the environment character for the environment string
'' Inputs:      None
'' Returns:     Environment Character
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EnvChar() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case UCase(m.strEnvironment)
        Case "GATEWAY"
            strReturn = "G"
        Case "CLIENT"
            strReturn = "C"
        Case "TEST CLIENT"
            strReturn = "T"
        Case "TEST GATEWAY"
            strReturn = "g"
        Case "DEMO CLIENT"
            strReturn = "D"
        Case Else
            strReturn = ""
    End Select
    
    EnvChar = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.EnvChar"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MessageTypeToString
'' Description: Convert a message type enumeration to a string
'' Inputs:      Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MessageTypeToString(ByVal nType As eGDPatsMessageTypes) As String
On Error GoTo ErrSection:

    Select Case nType
        Case eGDPatsMessageType_Connect
            MessageTypeToString = "eGDPatsMessageType_Connect"
        Case eGDPatsMessageType_Disconnect
            MessageTypeToString = "eGDPatsMessageType_Disconnect"
        Case eGDPatsMessageType_AddOrder
            MessageTypeToString = "eGDPatsMessageType_AddOrder"
        Case eGDPatsMessageType_AmendOrder
            MessageTypeToString = "eGDPatsMessageType_AmendOrder"
        Case eGDPatsMessageType_CancelOrder
            MessageTypeToString = "eGDPatsMessageType_CancelOrder"
        Case eGDPatsMessageType_UnloadApp
            MessageTypeToString = "eGDPatsMessageType_UnloadApp"
        Case eGDPatsMessageType_GetAccounts
            MessageTypeToString = "eGDPatsMessageType_GetAccounts"
        Case eGDPatsMessageType_GetOrders
            MessageTypeToString = "eGDPatsMessageType_GetOrders"
        Case eGDPatsMessageType_GetFills
            MessageTypeToString = "eGDPatsMessageType_GetFills"
        Case eGDPatsMessageType_GetPositions
            MessageTypeToString = "eGDPatsMessageType_GetPositions"
        Case eGDPatsMessageType_GetContracts
            MessageTypeToString = "eGDPatsMessageType_GetContracts"
        Case eGDPatsMessageType_Subscribe
            MessageTypeToString = "eGDPatsMessageType_Subscribe"
        Case eGDPatsMessageType_Unsubscribe
            MessageTypeToString = "eGDPatsMessageType_Unsubscribe"
        
        Case eGDPatsMessageType_ConnectionInfo
            MessageTypeToString = "eGDPatsMessageType_ConnectionInfo"
        Case eGDPatsMessageType_AppLoaded
            MessageTypeToString = "eGDPatsMessageType_AppLoaded"
        Case eGDPatsMessageType_AppUnloaded
            MessageTypeToString = "eGDPatsMessageType_AppUnloaded"
        Case eGDPatsMessageType_Heartbeat
            MessageTypeToString = "eGDPatsMessageType_Heartbeat"
        Case eGDPatsMessageType_Order
            MessageTypeToString = "eGDPatsMessageType_Order"
        Case eGDPatsMessageType_AccountR
            MessageTypeToString = "eGDPatsMessageType_AccountR"
        Case eGDPatsMessageType_OrderR
            MessageTypeToString = "eGDPatsMessageType_OrderR"
        Case eGDPatsMessageType_FillR
            MessageTypeToString = "eGDPatsMessageType_FillR"
        Case eGDPatsMessageType_PositionR
            MessageTypeToString = "eGDPatsMessageType_PositionR"
        Case eGDPatsMessageType_ContractR
            MessageTypeToString = "eGDPatsMessageType_ContractR"
        Case eGDPatsMessageType_QuoteR
            MessageTypeToString = "eGDPatsMessageType_QuoteR"
            
        Case Else
            MessageTypeToString = Str(nType)
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPats.MessageTypeToString"
    
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
    Open m.strLogPath & "\TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append Shared As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " - " & strMessage
        Close #fh
    End If

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

    ConnectionStatus = eGDConnectionStatus_Disconnected
    
    Set m.astrToFix = New cGdArray
    m.astrToFix.Create eGDARRAY_Strings
    
    m.bWasConnected = False
    m.bVerifyPositions = True
    m.bSyncInProgress = False
    m.bFixCalledOnce = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.Class_Initialize"
    
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
    Set m.astrAccounts = Nothing
    Set m.astrBrokerPos = Nothing
    Set m.astrSymbols = Nothing
    Set m.astrToFix = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPats.Class_Terminate"
    
End Sub
