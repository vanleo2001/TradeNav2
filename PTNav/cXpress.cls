VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXpress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cXpress.cls
'' Description: Wrapper class for routines to access the LindXpress API
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 01/09/2009   DAJ         Put code in to better handle Balance Cancelled orders
'' 02/04/2009   DAJ         Dump some additional information when amending an order
'' 02/05/2009   DAJ         Added support for 1/2 64ths
'' 03/26/2009   DAJ         Use CommonFilesPath function instead of hard-coded
'' 04/28/2009   DAJ         If BrokerID is Zero, blank it out
'' 05/19/2009   DAJ         Use the new Option Navigator status variable
'' 05/27/2009   DAJ         If Add or Amend failed, send error to Option Nav
'' 06/01/2009   DAJ         Added FO and SO to security type mask for trading
'' 06/09/2009   DAJ         Set the BrokerInfo connection status
'' 09/01/2009   DAJ         Use new Parked order status
'' 10/27/2009   DAJ         Allow GenesisSymbolForPosition to handle option symbol
'' 02/22/2010   DAJ         Call order BalCancelled if Amended and partially filled
'' 03/11/2010   DAJ         Use global collections
'' 11/18/2010   DAJ         Added Starting Genesis ID for Option Navigator
'' 12/15/2010   DAJ         Added ShowAccountOnline for Xpress
'' 03/07/2011   DAJ         Changed OptionNav calls, Added Change Password calls,
''                          Broker Disconnect Alerts
'' 03/15/2011   DAJ         For positions check both "customer" symbol and "internal" symbol
'' 05/11/2011   DAJ         Fixed BuildPositions for options
'' 05/25/2011   DAJ         Changed the refresh start time to last refresh message
'' 05/27/2011   DAJ         Send Manual/Automated flag along with an order
'' 06/21/2011   DAJ         Last connection status change
'' 01/18/2012   DAJ         Enhanced logging for automated trading
'' 01/31/2012   DAJ         Handle unsolicted cancels in TradeSense order groups
'' 02/14/2012   DAJ         New status alerts for position mismatch / auto trade disabled
'' 03/21/2012   DAJ         Fix for amending a partially filled order
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 12/11/2012   DAJ         Use the flatten queue for position reversals
'' 12/11/2012   DAJ         Handle broker status in automated trading item
'' 04/17/2013   DAJ         Added direction filter to CancelWorkingOrders call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strIniFile As String                ' Ini file for LindXpress
    strIpIniFile As String              ' Ini file for IP information
    nBroker As eTT_AccountType          ' Broker type for the TT broker
    strBrokerName As String             ' Broker name for the TT broker
    strProcessName As String            ' Name of the process to check
    strProcessPath As String            ' Full path of the process to start
    strControlID As String              ' Control ID
    strLogPath As String                ' Path for the log files
    
    strUserName As String               ' User Name to get onto the servers
    strLoginAccount As String           ' Account to get onto the servers
    strPassword As String               ' Password to get onto the servers
    strFirm As String                   ' Firm to use with this logon
    strHostIP As String                 ' Host to use with this logon
    strHostPort As String               ' Port to use with this logon
    iSubsystem As Integer               ' SubSystem to use with this logon
    strVersion As String                ' Trade Navigator Version to pass along
    
    nStatus As eGDConnectionStatus      ' Connection status to the LindXpress servers
    bWasConnected As Boolean            ' This broker was has been connected
    dHeartBeat As Double                ' Tick count as of last message receieved
    dLastStatusMsg As Double            ' Tick count of the last status message
    bReconnecting As Boolean            ' Are we currently reconnecting?
    lReconnectAttempts As Long          ' Reconnect attempts that have been made
    bAppLoaded As Boolean               ' Is the LindXpress application loaded?
    bManualDisconnect As Boolean        ' Is the user manually disconnecting?
    bAskedPassword As Boolean           ' Have we asked for the password yet?
    dLastFillTime As Double             ' Tick count of when we received the last fill
    strLastMessage As String            ' Last message from the connection status message
    bVerifyPositions As Boolean         ' Do we want to verify the positions with the user?
    bVerboseRefresh As Boolean          ' Do we want to tell the user about the refresh?
    bConfirmCancel As Boolean           ' Does the user want to confirm each order cancel?
    bSyncInProgress As Boolean          ' Are we in the process of synchronizing positions?
    bFixCalledOnce As Boolean           ' Has the fix positions call been done at least once?
    lLastFromDateFills As Long          ' Last from date for a fill refresh
    
    astrAccounts As cGdArray            ' List of accounts currently connected to
    astrSymbols As cGdArray             ' Array of base symbol information
    astrBrokerPos As cGdArray           ' List of current positions
    astrAmendPending As cGdArray        ' Pending order amend information array
    astrToFix As cGdArray               ' List of account/symbol pairs that need their position fixed
    
    dLastRefreshMesage As Double        ' Time of the last refresh message
    bRefreshingAccounts As Boolean      ' Are we refreshing accounts?
    bRefreshingOrders As Boolean        ' Are we refreshing orders?
    bRefreshingFills As Boolean         ' Are we refreshing fills?
    bRefreshingPositions As Boolean     ' Are we refreshing positions?
        
    BrokerInfo As cBrokerInfo           ' Object of broker information
End Type
Private m As mPrivate

Public Property Get UserName() As String
    UserName = m.strUserName
End Property
Public Property Let UserName(ByVal strUserName As String)
    m.strUserName = strUserName
End Property

Public Property Get LoginAccount() As String
    LoginAccount = m.strLoginAccount
End Property
Public Property Let LoginAccount(ByVal strLoginAccount As String)
    m.strLoginAccount = strLoginAccount
End Property

Public Property Get Login() As String
    If Len(m.strUserName) > 0 Then
        Login = "U:" & m.strUserName
    ElseIf Len(m.strLoginAccount) > 0 Then
        Login = "A:" & m.strLoginAccount
    Else
        Login = ""
    End If
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

Public Property Get Firm() As String
    Firm = m.strFirm
End Property
Public Property Let Firm(ByVal strFirm As String)
    m.strFirm = strFirm
End Property

Public Property Get Subsystem() As Integer
    Subsystem = m.iSubsystem
End Property
Public Property Let Subsystem(ByVal iSubsystem As Integer)
    m.iSubsystem = iSubsystem
End Property

Public Property Get ConfirmCancel() As Boolean
    ConfirmCancel = m.bConfirmCancel
End Property
Public Property Let ConfirmCancel(ByVal bConfirmCancel As Boolean)
    m.bConfirmCancel = bConfirmCancel
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

Public Property Get EnabledSymbols() As cGdArray
    Set EnabledSymbols = Nothing
End Property

Public Property Get ProcessName() As String
    ProcessName = m.strProcessName
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
    RaiseError "cXpress.CurrentPosition.Get"
    
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
    RaiseError "cXpress.CurrentPosition.Let"
    
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
    RaiseError "cXpress.CarriedPosition.Get"
    
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
    RaiseError "cXpress.CarriedPosition.Let"
    
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
        Case eTT_AccountType_LindWaldock
            m.strIniFile = AddSlash(App.Path) & "LindWaldock.INI"
            m.strIpIniFile = AddSlash(App.Path) & "Provided\LwIps.INI"
            m.nBroker = eTT_AccountType_LindWaldock
            m.strBrokerName = "Lind Waldock"
            m.strProcessName = "GenLindWaldock"
            m.strProcessPath = AddSlash(App.Path) & "..\Brokers\LindWaldock\GenLindWaldock.EXE"
            m.strControlID = "LindWaldock"
            m.strLogPath = AddSlash(App.Path) & "LindWaldock"
            strSymbolFile = AddSlash(App.Path) & "Provided\LwToGen.TXT"
        
        Case eTT_AccountType_ManExpress
            m.strIniFile = AddSlash(App.Path) & "ManExpress.INI"
            m.strIpIniFile = AddSlash(App.Path) & "Provided\MxIps.INI"
            m.nBroker = eTT_AccountType_ManExpress
            m.strBrokerName = "Man Express"
            m.strProcessName = "GenManExpress"
            m.strProcessPath = AddSlash(App.Path) & "..\Brokers\ManExpress\GenManExpress.EXE"
            m.strControlID = "ManExpress"
            m.strLogPath = AddSlash(App.Path) & "ManExpress"
            strSymbolFile = AddSlash(App.Path) & "Provided\MxToGen.TXT"
            
    End Select
            
    If Not DirExist(m.strLogPath) Then MkDir m.strLogPath
    KillFile m.strLogPath & "\*.LOG /o=-30"
    
    Set m.BrokerInfo = New cBrokerInfo
    m.BrokerInfo.Init m.nBroker, m.strLogPath
    
    Set m.astrAccounts = New cGdArray
    m.astrAccounts.Create eGDARRAY_Strings
    
    Set m.astrBrokerPos = New cGdArray
    m.astrBrokerPos.Create eGDARRAY_Strings
    
    Set m.astrAmendPending = New cGdArray
    m.astrAmendPending.Create eGDARRAY_Strings
    
    Set m.astrSymbols = New cGdArray
    m.astrSymbols.Create eGDARRAY_Strings
    m.astrSymbols.FromFile strSymbolFile
    m.astrSymbols.Sort
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.Init"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowConfig
'' Description: Determines whether we need to show the configuration
'' Inputs:      None
'' Returns:     True if need to show configuration, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowConfig() As Boolean
On Error GoTo ErrSection:

    If Len(m.strUserName) = 0 Or Len(m.strHostIP) = 0 Or Len(m.strHostPort) = 0 Then
        ShowConfig = True
    Else
        ShowConfig = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.ShowConfig"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Connect
'' Description: Attempt a connection to the Lind Express servers
'' Inputs:      New Login, Are we switching?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Connect(Optional ByVal strNewLogin As String = "", Optional ByVal bSwitching As Boolean = False)
On Error GoTo ErrSection:

    Dim bConnect As Boolean             ' Should we still try to connect?
    Dim lTimeOut As Long                ' Timeout variable
    
    If g.bUnloading = False Then
        If (ConnectionStatus = eGDConnectionStatus_Disconnected) Then
            If (m.bAskedPassword = True) And ((strNewLogin = Login) Or (Len(strNewLogin) = 0)) Then
                bConnect = True
            Else
                bConnect = frmLwLogin.ShowMe(m.nBroker, strNewLogin)
                m.bAskedPassword = bConnect
            End If
        Else
            If ((strNewLogin <> Login) And (Len(strNewLogin) > 0)) Or (bSwitching = True) Then
                bConnect = frmLwLogin.ShowMe(m.nBroker, strNewLogin, False, bSwitching)
                If bConnect = True Then
                    m.bAskedPassword = True
                    Disconnect False, "Changing login to " & strNewLogin
                    
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
                m.strVersion = "Trade Navigator" & " " & App.Major & "." & App.Minor & "." & App.Revision
    
                ConnectionStatus = eGDConnectionStatus_Connecting
                
                If ((frmOnlineBroker.gdBroker.FindControlID(m.strControlID) = 0) Or (KillProcess(m.strProcessName, True) = 0)) Then
                    If KillProcess(m.strProcessName) > 0 Then
                        Sleep 3
                    End If
                    
                    FileCopy AddSlash(App.Path) & "..\Brokers\GenLindXpress.EXE", m.strProcessPath, True
                    
                    DumpDebug "Trade Navigator version: " & FormatVersion & "." & Str(App.Revision) & " " & DateFormat(FileDate(App.Path & "\" & App.EXEName & ".EXE"), MM_DD_YYYY, HH_MM, AMPM_UPPER)
                    DumpDebug "Loading application " & m.strProcessPath
                    RunProcess m.strProcessPath, , , vbHide
                Else
                    DumpDebug "Connecting to " & m.strBrokerName & " as " & Login & " (" & m.strHostIP & ":" & m.strHostPort & ", " & m.strFirm & ", " & Str(m.iSubsystem) & ")"
                    SendLindXpressMessage eGDLindXpressMessageType_Connect, m.strUserName & vbTab & m.strPassword & vbTab & m.strHostIP & vbTab & m.strHostPort & vbTab & m.strFirm & vbTab & Str(m.iSubsystem) & vbTab & m.strLoginAccount & vbTab & m.strVersion
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.Connect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disconnect
'' Description: Disconnect form the Lind Express servers
'' Inputs:      Force Reconnect?, Reason for Disconnect
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Disconnect(Optional ByVal bForceReconnect As Boolean = False, Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    If (ConnectionStatus = eGDConnectionStatus_Connected) Then
        ConnectionStatus = eGDConnectionStatus_Disconnecting
        
        m.bManualDisconnect = Not bForceReconnect
        If m.bManualDisconnect Then
            m.dHeartBeat = 0#
            g.OrderStrategies.DeactivateExitsForBroker m.nBroker, "Disconnecting from " & m.strBrokerName
        End If
        
        If Len(strReason) = 0 Then DumpDebug "Disconnecting from " & m.strBrokerName Else DumpDebug "Disconnecting from " & m.strBrokerName & ": " & strReason
        SendLindXpressMessage eGDLindXpressMessageType_Disconnect, ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.Disconnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reconnect
'' Description: Attempt to reconnect to the Lind Express servers
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
            Disconnect True, "Reconnecting"
        End If
        
        lTimeOut = 0&
        Do While (m.bAppLoaded = True) And (lTimeOut < 30&)
            Sleep 1#
            lTimeOut = lTimeOut + 1&
        Loop
                        
        If (ConnectionStatus <> eGDConnectionStatus_Connected) And (ConnectionStatus <> eGDConnectionStatus_Connecting) Then
            If (m.bAppLoaded = False) Or (lTimeOut >= 30&) Then
                m.lReconnectAttempts = m.lReconnectAttempts + 1&
                ConnectionStatus = eGDConnectionStatus_Disconnected
                Connect
            End If
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
    RaiseError "cXpress.Reconnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowAccountOnline
'' Description: Show the users account information online
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowAccountOnline()
On Error GoTo ErrSection:

    Dim strAddress As String            ' Address to Xpress URL
    
    strAddress = GetIniFileProperty("Address", "", "Online", m.strIpIniFile)
    If Len(strAddress) > 0 Then
        RunProcess InternetBrowser, Chr(34) & strAddress & Chr(34)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.ShowAccountOnline"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendOrder
'' Description: Attempt to add or amend the order as appropriate
'' Inputs:      Order to Send
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendOrder(Order As cPtOrder, ByVal strPrevGenesisOrderID As String)
On Error GoTo ErrSection:
    
    Dim nStatus As eTT_OrderStatus      ' New order status

    If NotSent(Order.Status) Or (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        If AddOrder(Order) = False Then
            nStatus = eTT_OrderStatus_Error
        Else
            nStatus = eTT_OrderStatus_Sent
        End If
    Else
        If AmendOrder(Order, strPrevGenesisOrderID) = False Then
            nStatus = eTT_OrderStatus_Error
        Else
            nStatus = eTT_OrderStatus_AmendPending
        End If
    End If
    
    Order.ChangeOrderStatus nStatus, m.BrokerInfo

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.SendOrder"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOrder
'' Description: Send a new order to the Lind Express servers
'' Inputs:      Order
'' Returns:     True if successfully converted, False otherwise
''
'' Order:       Account, Genesis ID, Broker ID, Message, NumLegs, B/S, Quantity,
''              Symbol, Order Type, Price, With Limit Price, Expiration,
''              Manual Order Flag
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AddOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Array of order information
    Dim strLindXpressSymbol As String   ' LindXpress version of the Genesis symbol
    Dim dMult As Double                 ' Multiplier
    Dim strLindXpressType As String     ' LindXpress order type
    Dim strMessage As String            ' Message to send to the fill server
    Dim bReturn As Boolean              ' Return value from the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    Set astrOrder = New cGdArray
    astrOrder.Create eGDARRAY_Strings
    
    bReturn = False
    strLindXpressSymbol = LindXpressSymbol(Order.Symbol, dMult)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strLindXpressType = LindXpressOrderType(Order.OrderType)
    
    If Len(strLindXpressSymbol) = 0 Then
        DumpDebug "AddOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
    ElseIf Len(strLindXpressType) = 0 Then
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
        
        strMessage = ""
        If m.strUserName = "GENFIL" Then
            strMessage = InfBox("Please specify the Order Message:", , , "Order Message", , , , , , "string", "")
        End If
        
        astrOrder(0) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(1) = Order.GenesisOrderID
        astrOrder(2) = ""
        astrOrder(3) = strMessage
        astrOrder(4) = "1"
        If Order.Buy Then astrOrder(5) = "1" Else astrOrder(5) = "2"
        astrOrder(6) = Str(Order.Quantity)
        astrOrder(7) = strLindXpressSymbol
        astrOrder(8) = strLindXpressType
        Select Case Order.OrderType
            Case eTT_OrderType_Market, eTT_OrderType_MarketOnClose
                astrOrder(9) = ""
                astrOrder(10) = ""
            Case eTT_OrderType_Stop, eTT_OrderType_StopCloseOnly
                astrOrder(9) = Str(LindXpressPrice(Order.StopPrice, dMult))
                astrOrder(10) = ""
            Case eTT_OrderType_Limit
                astrOrder(9) = Str(LindXpressPrice(Order.LimitPrice, dMult))
                astrOrder(10) = ""
            Case eTT_OrderType_StopWithLimit
                astrOrder(9) = Str(LindXpressPrice(Order.StopPrice, dMult))
                astrOrder(10) = Str(LindXpressPrice(Order.LimitPrice, dMult))
        End Select
        If Order.Expiration < 0& Then
            astrOrder(11) = "        "
        ElseIf Order.Expiration = 0& Then
            astrOrder(11) = "CANCEL  "
        Else
            astrOrder(11) = Format(Order.Expiration, "YYYYMMDD")
        End If
        If Order.IsAutomated Then
            astrOrder(12) = "A"
        Else
            astrOrder(12) = "M"
        End If
        
        SendLindXpressMessage eGDLindXpressMessageType_AddOrder, astrOrder.JoinFields(vbTab)
        bReturn = True
    End If
        
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
    AddOrder = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.AddOrder"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddAmendPendingInfo
'' Description: Add the given information to the amend pending array
'' Inputs:      Broker ID, Old Access and Genesis ID's, New Access and Genesis ID's
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddAmendPendingInfo(ByVal strBrokerID As String, ByVal lOldAccessID As Long, ByVal strOldGenesisID As String, ByVal lNewAccessID As Long, ByVal strNewGenesisID As String)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array

    If m.astrAmendPending.BinarySearch(strBrokerID & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = False Then
        m.astrAmendPending.Add strBrokerID & vbTab & Str(lOldAccessID) & vbTab & strOldGenesisID & vbTab & Str(lNewAccessID) & vbTab & strNewGenesisID, lPos
        DumpDebug "Amend Pending Info Added: " & m.astrAmendPending(lPos)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.AddAmendPendingInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AmendOrder
'' Description: Send a modify order to the Lind Express servers
'' Inputs:      Order, Previous Order ID
'' Returns:     True if successfully converted, False otherwise
''
'' Order:       Account, Genesis ID, Broker ID, Message, NumLegs, B/S, Quantity,
''              Symbol, Order Type, Price, With Limit Price, Expiration,
''              Previous Genesis Order ID, Manual Order Flag
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AmendOrder(ByVal Order As cPtOrder, ByVal strPrevGenesisOrderID As String) As Boolean
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Array of order information
    Dim strLindXpressSymbol As String   ' LindXpress version of the Genesis symbol
    Dim dMult As Double                 ' Multiplier
    Dim strLindXpressType As String     ' LindXpress order type
    Dim strMessage As String            ' Message to send to the fill server
    Dim bReturn As Boolean              ' Return value from the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    Set astrOrder = New cGdArray
    astrOrder.Create eGDARRAY_Strings
    
    bReturn = False
    strLindXpressSymbol = LindXpressSymbol(Order.Symbol, dMult)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strLindXpressType = LindXpressOrderType(Order.OrderType)
    
    If Len(strLindXpressSymbol) = 0 Then
        DumpDebug "AmendOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
    ElseIf Len(strLindXpressType) = 0 Then
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
        DumpDebug "AmendOrder: " & Order.OrderText & "(" & Order.GenesisOrderID & ", " & Order.BrokerID & ")"
        
        strMessage = ""
        If m.strUserName = "GENFIL" Then
            strMessage = InfBox("Please specify the Order Message:", , , "Order Message", , , , , , "string", "")
        End If
        
        astrOrder(0) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(1) = Order.GenesisOrderID
        If InStr(Order.PreviousBrokerID, "-") = 0 Then
            astrOrder(2) = Order.PreviousBrokerID
        Else
            astrOrder(2) = Parse(Order.PreviousBrokerID, "-", 2)
        End If
        astrOrder(3) = strMessage
        astrOrder(4) = "1"
        If Order.Buy Then astrOrder(5) = "1" Else astrOrder(5) = "2"
        astrOrder(6) = Str(Order.Quantity)
        astrOrder(7) = strLindXpressSymbol
        astrOrder(8) = strLindXpressType
        Select Case Order.OrderType
            Case eTT_OrderType_Market, eTT_OrderType_MarketOnClose
                astrOrder(9) = ""
                astrOrder(10) = ""
            Case eTT_OrderType_Stop, eTT_OrderType_StopCloseOnly
                astrOrder(9) = Str(LindXpressPrice(Order.StopPrice, dMult))
                astrOrder(10) = ""
            Case eTT_OrderType_Limit
                astrOrder(9) = Str(LindXpressPrice(Order.LimitPrice, dMult))
                astrOrder(10) = ""
            Case eTT_OrderType_StopWithLimit
                astrOrder(9) = Str(LindXpressPrice(Order.StopPrice, dMult))
                astrOrder(10) = Str(LindXpressPrice(Order.LimitPrice, dMult))
        End Select
        If Order.Expiration < 0& Then
            astrOrder(11) = "        "
        ElseIf Order.Expiration = 0& Then
            astrOrder(11) = "CANCEL  "
        Else
            astrOrder(11) = Format(Order.Expiration, "YYYYMMDD")
        End If
        astrOrder(12) = strPrevGenesisOrderID
        If Order.IsAutomated Then
            astrOrder(13) = "A"
        Else
            astrOrder(13) = "M"
        End If
        
        SendLindXpressMessage eGDLindXpressMessageType_AmendOrder, astrOrder.JoinFields(vbTab)
        bReturn = True
    End If
        
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
    AmendOrder = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.AmendOrder"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Send a cancel order to the Lind Express servers
'' Inputs:      Order
'' Returns:     True if successfully converted, False otherwise
''
'' Order:       Account, Genesis ID, Broker ID, Confirmation, Manual Order Flag
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim strAutomated As String          ' Automated flag to send

    If g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True Then
        DumpDebug "CancelOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
        
        Order.Status = eTT_OrderStatus_Cancelled
        Order.StatusDate = ConvertToBrokerDate(CurrentTime, g.Broker.AccountTypeForID(Order.AccountID), Order.Symbol, True)
        Order.Save
        
        g.TradingItems.OrderCallback Order.AutoTradeItemID, Order
        g.Broker.AddOrder Order
        OrderCallback Order
        
        g.Broker.SyntheticOrders.Remove Str(Order.OrderID)
        
        DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
    Else
        If Not g.TsoGroups Is Nothing Then
            g.TsoGroups.CancelRequested Order
        End If
        
        DumpDebug "CancelOrder: " & Order.OrderText & "(" & Order.GenesisOrderID & ", " & Order.BrokerID & ")"
        
        If Order.IsAutomated Then
            strAutomated = "A"
        Else
            strAutomated = "M"
        End If
        
        If InStr(Order.BrokerID, "-") = 0 Then
            SendLindXpressMessage eGDLindXpressMessageType_CancelOrder, g.Broker.AccountNumberForID(Order.AccountID) & vbTab & Order.GenesisOrderID & vbTab & Order.BrokerID & vbTab & "0" & vbTab & strAutomated
        Else
            SendLindXpressMessage eGDLindXpressMessageType_CancelOrder, g.Broker.AccountNumberForID(Order.AccountID) & vbTab & Order.GenesisOrderID & vbTab & Parse(Order.BrokerID, "-", 2) & vbTab & "0" & vbTab & strAutomated
        End If
        
        Order.Status = eTT_OrderStatus_CancelPending
        Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
        Order.Save
        
        m.BrokerInfo.AddOrder Order, False
    End If
    
    CancelOrder = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.CancelOrder"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Park the given order and Cancel it on the broker server
'' Inputs:      Order
'' Returns:     True if successfully converted, False otherwise
''
'' Order:       Account, Genesis ID, Broker ID, Confirmation, Manual Order Flag
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ParkOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim strAutomated As String          ' Automated flag to send

    If g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True Then
        DumpDebug "ParkOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
        
        Order.Status = eTT_OrderStatus_Parked
        Order.StatusDate = ConvertToBrokerDate(CurrentTime, g.Broker.AccountTypeForID(Order.AccountID), Order.Symbol, True)
        Order.Save
        
        g.TradingItems.OrderCallback Order.AutoTradeItemID, Order
        g.Broker.AddOrder Order
        OrderCallback Order
        
        g.Broker.SyntheticOrders.Remove Str(Order.OrderID)
        
        DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
    Else
        DumpDebug "ParkOrder: " & Order.OrderText & "(" & Order.GenesisOrderID & ", " & Order.BrokerID & ")"
        
        If Order.IsAutomated Then
            strAutomated = "A"
        Else
            strAutomated = "M"
        End If
        
        If InStr(Order.BrokerID, "-") = 0 Then
            SendLindXpressMessage eGDLindXpressMessageType_CancelOrder, g.Broker.AccountNumberForID(Order.AccountID) & vbTab & Order.GenesisOrderID & vbTab & Order.BrokerID & vbTab & "0" & vbTab & strAutomated
        Else
            SendLindXpressMessage eGDLindXpressMessageType_CancelOrder, g.Broker.AccountNumberForID(Order.AccountID) & vbTab & Order.GenesisOrderID & vbTab & Parse(Order.BrokerID, "-", 2) & vbTab & "0" & vbTab & strAutomated
        End If
        
        Order.Status = eTT_OrderStatus_ParkPending
        Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
        Order.Save
        
        m.BrokerInfo.AddOrder Order, False
    End If
    
    ParkOrder = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.ParkOrder"

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
    RaiseError "cXpress.CancelOrderFromBrokerID"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelWorkingOrders
'' Description: Cancel all working orders for the given account and symbol
'' Inputs:      Account, Symbol, Auto Trade Item ID, Skip Order ID, Direction
'' Returns:     Number of orders cancelled
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelWorkingOrders(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, Optional ByVal lSkipOrderID As Long = -1&, Optional ByVal nDirection As eGDFilterDirection = eGDFilterDirection_All) As Long
On Error GoTo ErrSection:
    
    Dim Orders As cPtOrders             ' Collection of working orders
    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value for the function
    
    DumpDebug "Cancelling Working Orders for: " & strAccount & ", " & strSymbol & ", " & Str(lAtID)
    Set Orders = m.BrokerInfo.OrdersForSymbol(strAccount, strSymbol, lAtID, True, nDirection)
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
    RaiseError "cXpress.CancelWorkingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenPosition
'' Description: Flatten the position for the given account and symbol
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
    RaiseError "cXpress.FlattenPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReversePosition
'' Description: Reverse the position for the given account and symbol
'' Inputs:      Account, Symbol, Auto Trade Item ID, Confirm?, Genesis ID
'' Returns:     Did user Confirm?
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReversePosition(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, Optional ByVal bConfirm As Boolean = True, Optional strGenesisOrderID As String) As Boolean
On Error GoTo ErrSection:

    Dim lPosition As Long               ' Current position
    Dim strMessage As String            ' Message to display to the user
    Dim strReturn As String             ' Return from user question
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    DumpDebug "Reversing Position for: " & strAccount & ", " & strSymbol & ", " & Str(lAtID)
    lPosition = m.BrokerInfo.CurrentPosition(strAccount, strSymbol, lAtID)
    If lPosition <> 0& Then
        If bConfirm Then
            If lPosition < 0 Then
                strMessage = "BUY " & Abs(lPosition * 2) & " " & strSymbol & " at the MARKET"
            Else
                strMessage = "SELL " & Abs(lPosition * 2) & " " & strSymbol & " at the MARKET"
            End If
            strReturn = InfBox("Are you sure you want to|" & strMessage & "|in account '" & strAccount & "'|to reverse your position?", "?", "+Yes|-No", "Reverse for " & strSymbol & " in " & strAccount)
        Else
            strReturn = "Y"
        End If
        
        If strReturn = "Y" Then
            bReturn = True
            DumpDebug "User Confirmed: " & strMessage & " in account " & strAccount
            ReversePositionForSymbol strAccount, strSymbol, lAtID, lPosition, strGenesisOrderID
        End If
    End If
    
    ReversePosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.ReversePosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Send a request for the accounts on the LindXpress servers
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
        SendLindXpressMessage eGDLindXpressMessageType_GetAccounts, ""
    Else
        DumpDebug "Account Refresh Requested but already in Progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.GetAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Send a request for the orders on a specific account
'' Inputs:      Account, From Date
'' Returns:     None
''
'' Fields:      Account (blank for all), From Date (blank for snapshot only)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetOrders(Optional ByVal strAccount As String = "", Optional ByVal strFromDate As String = "")
On Error GoTo ErrSection:

    If Not Refreshing Then
        DumpDebug "Getting Orders for " & strAccount
        m.bRefreshingOrders = True
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for order information...", , , m.strBrokerName & " Refresh", True
        End If
        If Len(strFromDate) > 0 Then
            SendLindXpressMessage eGDLindXpressMessageType_GetOrders, strAccount & vbTab & strFromDate
        Else
            SendLindXpressMessage eGDLindXpressMessageType_GetOrders, strAccount & vbTab & FromDateOrders
        End If
    Else
        DumpDebug "Order Refresh for " & strAccount & " requested but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.GetOrders"
    
End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllOrders
'' Description: Send a request for orders on all of the accounts
'' Inputs:      None
'' Returns:     None
''
'' Fields:      Account, From Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAllOrders()
On Error GoTo ErrSection:

    If Not Refreshing Then
        DumpDebug "Getting Orders for all accounts"
        m.bRefreshingOrders = True
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for order information...", , , m.strBrokerName & " Refresh", True
        End If
        SendLindXpressMessage eGDLindXpressMessageType_GetAllOrders, ""
    Else
        DumpDebug "Order Refresh for all accounts requested but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.GetAllOrders"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFills
'' Description: Send a request for the fills on a specific account
'' Inputs:      Account, From Date
'' Returns:     None
''
'' Fields:      Account (blank for all), From Date (blank for snapshot only)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetFills(Optional ByVal strAccount As String = "", Optional ByVal strFromDate As String = "")
On Error GoTo ErrSection:

    If Not Refreshing Then
        DumpDebug "Getting Fills for " & strAccount
        m.bRefreshingFills = True
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for fill information...", , , m.strBrokerName & " Refresh", True
        End If
        If Len(strFromDate) > 0 Then
            SendLindXpressMessage eGDLindXpressMessageType_GetFills, strAccount & vbTab & strFromDate
        Else
            SendLindXpressMessage eGDLindXpressMessageType_GetFills, strAccount & vbTab & FromDateFills
        End If
    Else
        DumpDebug "Fill Refresh requested for " & strAccount & " but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.GetFills"
    
End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllFills
'' Description: Send a request for fills on all of the accounts
'' Inputs:      None
'' Returns:     None
''
'' Fields:      Account, From Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAllFills()
On Error GoTo ErrSection:

    If Not Refreshing Then
        DumpDebug "Getting Fills for all accounts"
        m.bRefreshingFills = True
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for fill information...", , , m.strBrokerName & " Refresh", True
        End If
        SendLindXpressMessage eGDLindXpressMessageType_GetAllFills, ""
    Else
        DumpDebug "Fill Refresh for all accounts requested but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.GetAllFills"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositions
'' Description: Send a request for the positions on a specific account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPositions(ByVal strAccount As String)
On Error GoTo ErrSection:

    If Not Refreshing Then
        DumpDebug "Getting Positions for " & strAccount
        m.bRefreshingPositions = True
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for position information...", , , m.strBrokerName & " Refresh", True
        End If
        SendLindXpressMessage eGDLindXpressMessageType_GetPositions, strAccount
    Else
        DumpDebug "Position Refresh requested for " & strAccount & " but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllPositions
'' Description: Send a request for positions on all of the accounts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAllPositions()
On Error GoTo ErrSection:

    If Not Refreshing Then
        DumpDebug "Getting Positions for all accounts"
        m.bRefreshingPositions = True
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for position information...", , , m.strBrokerName & " Refresh", True
        End If
        SendLindXpressMessage eGDLindXpressMessageType_GetAllPositions, ""
    Else
        DumpDebug "Position Refresh for all accounts requested but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.GetAllPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccountInfo
'' Description: Send a request for account information for an account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAccountInfo(ByVal strAccount As String)
On Error GoTo ErrSection:

    DumpDebug "Getting account information for " & strAccount
    SendLindXpressMessage eGDLindXpressMessageType_GetAccountInfo, strAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.GetAccountInfo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Refresh
'' Description: Send a request for the orders, fills, and positions for account
'' Inputs:      Account
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
    RaiseError "cXpress.Refresh"
    
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
    RaiseError "cXpress.ConnectionStatusForAccount"
    
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
    lCustomerID = RI_GetLastDataServiceID \ 1000
    
    If lDate <> Date Then
        lDate = Date
        lCounter = 0&
    Else
        lCounter = lCounter + 1
    End If

    SetIniFileProperty "Date", lDate, "ID", m.strIniFile
    SetIniFileProperty "Counter", lCounter, "ID", m.strIniFile
    
    NextGenesisID = Format(lCustomerID, "0000000") & Format(lDate, "00000") & Format(lCounter, "00000")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.NextGenesisID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartingGenesisIdForOptNav
'' Description: Determine where Option Navigator should start its Genesis ID
'' Inputs:      None
'' Returns:     Starting ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StartingGenesisIdForOptNav() As String
On Error GoTo ErrSection:

    Dim lCustomerID As Long             ' Customer ID from the registry
    
    lCustomerID = (RI_GetLastDataServiceID \ 1000) Mod 1000000
    
    StartingGenesisIdForOptNav = Format(lCustomerID, "0000000") & Format(Date, "00000") & "90000"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.StartingGenesisIdForOptNav"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendLindXpressMessage
'' Description: Send an App Mail message to the LindXpress program
'' Inputs:      Type of Message, Message, Send Now?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendLindXpressMessage(ByVal nType As eGDLindXpressMessageTypes, ByVal strMessage As String, Optional ByVal bSendNow As Boolean = True)
On Error GoTo ErrSection:

    Dim astrMessage As New cGdArray     ' Message split out into an array

    If nType = eGDLindXpressMessageType_Connect Then
        astrMessage.SplitFields strMessage, vbTab
        astrMessage(1) = SetStrToChar(astrMessage(1), "*")
        DumpDebug vbTab & "Sending " & m.strBrokerName & " (" & MessageTypeToString(nType) & "): " & astrMessage.JoinFields(vbTab)
    ElseIf nType = eGDLindXpressMessageType_ChangePassword Then
        astrMessage.SplitFields strMessage, vbTab
        astrMessage(0) = SetStrToChar(astrMessage(0), "*")
        astrMessage(1) = SetStrToChar(astrMessage(1), "*")
        DumpDebug vbTab & "Sending " & m.strBrokerName & " (" & MessageTypeToString(nType) & "): " & astrMessage.JoinFields(vbTab)
    Else
        DumpDebug vbTab & "Sending " & m.strBrokerName & " (" & MessageTypeToString(nType) & "): " & strMessage
    End If
    
    frmOnlineBroker.gdBroker.CreateMessage m.strControlID, nType, strMessage, , bSendNow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.SendLindXpressMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming App Mail message from the LindXpress program
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDLindXpressMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    If (nType <> eGDLindXpressMessageType_ConnectionInfo) And (nType <> eGDLindXpressMessageType_Heartbeat) Then
        DumpDebug vbTab & m.strBrokerName & " Message Received (" & MessageTypeToString(nType) & "): " & strMessage
    End If

    Select Case nType
        Case eGDLindXpressMessageType_ConnectionInfo
            HandleConnectionInfo strMessage
        
        Case eGDLindXpressMessageType_AppLoaded
            m.bAppLoaded = True
            If LiveTradingAllowed(m.nBroker) Then
                DumpDebug "Connecting to " & m.strBrokerName & " as " & Login & " (" & m.strHostIP & ":" & m.strHostPort & ", " & m.strFirm & ", " & Str(m.iSubsystem) & ")"
                SendLindXpressMessage eGDLindXpressMessageType_Connect, m.strUserName & vbTab & m.strPassword & vbTab & m.strHostIP & vbTab & m.strHostPort & vbTab & m.strFirm & vbTab & Str(m.iSubsystem) & vbTab & m.strLoginAccount & vbTab & m.strVersion
            Else
                DumpDebug "Unloading the application since live trading is not allowed"
                SendLindXpressMessage eGDLindXpressMessageType_UnloadApp, ""
            End If
        
        Case eGDLindXpressMessageType_AppUnloaded
            m.bAppLoaded = False
            If m.bManualDisconnect = True Then
                m.dHeartBeat = 0#
            Else
                m.dHeartBeat = gdTickCount - 45000
            End If
            m.bManualDisconnect = False
            ConnectionStatus = eGDConnectionStatus_Disconnected
            
        Case eGDLindXpressMessageType_Heartbeat
            m.dHeartBeat = gdTickCount
            If m.dLastFillTime > 0 Then
                If gdTickCount > m.dLastFillTime + 2500 Then
                    m.dLastFillTime = 0#
                    SynchronizePositions
                End If
            End If
            
        Case eGDLindXpressMessageType_Account
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingAccounts = True
                m.BrokerInfo.ClearAccounts
                m.astrAccounts.Clear
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " account information...", , , m.strBrokerName & " Refresh", True
                End If
            ElseIf UCase(strMessage) = "END" Then
                m.BrokerInfo.RefreshingAccounts = False
                m.bRefreshingAccounts = False
                GetOrders
            Else
                If m.bRefreshingAccounts Then
                    m.dLastRefreshMesage = gdTickCount
                End If
                HandleAccountInfo strMessage
                m.BrokerInfo.AddAccount strMessage
                m.BrokerInfo.SetConnectionStatusForAccount eGDConnectionStatus_Connected, Parse(strMessage, vbTab, 1)
            End If
            
        Case eGDLindXpressMessageType_Order
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingOrders = True
                m.BrokerInfo.ClearOrders
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " order information...", , , m.strBrokerName & " Refresh", True
                End If
                SendBeginToOptionNav eGDOptNav_Order, m.nBroker
            ElseIf UCase(strMessage) = "END" Then
                SetIniFileProperty "Order", CurrentTime("CHI"), "Refresh", m.strIniFile
                m.BrokerInfo.RefreshingOrders = False
                m.bRefreshingOrders = False
                SendEndToOptionNav eGDOptNav_Order, m.nBroker
                GetFills
            Else
                If m.bRefreshingOrders Then
                    m.dLastRefreshMesage = gdTickCount
                End If
                'UpdateBrokerInfoOrder strMessage, False, False
                HandleOrderInfo strMessage, False, m.bRefreshingOrders
            End If
        
        Case eGDLindXpressMessageType_Fill
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingFills = True
                m.BrokerInfo.ClearFills
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " fill information...", , , m.strBrokerName & " Refresh", True
                End If
                SendBeginToOptionNav eGDOptNav_Fill, m.nBroker
            ElseIf UCase(strMessage) = "END" Then
                SetIniFileProperty "Fill", CurrentTime("CHI"), "Refresh", m.strIniFile
                SendEndToOptionNav eGDOptNav_Fill, m.nBroker
                
                m.BrokerInfo.RefreshingFills = False
                m.bRefreshingFills = False
                
                GetAllPositions
            Else
                If m.bRefreshingFills Then
                    m.dLastRefreshMesage = gdTickCount
                End If
                'UpdateBrokerInfoOrder strMessage, False, True
                HandleOrderInfo strMessage, True, m.bRefreshingFills
            End If
            
        Case eGDLindXpressMessageType_Position
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingPositions = True
                m.BrokerInfo.ClearPositions
                m.astrBrokerPos.Clear
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " position information...", , , m.strBrokerName & " Refresh", True
                End If
            ElseIf UCase(strMessage) = "END" Then
                SetIniFileProperty "Position", CurrentTime("CHI"), "Refresh", m.strIniFile
                m.BrokerInfo.RefreshingPositions = False
                m.bRefreshingPositions = False
            Else
                If m.bRefreshingPositions Then
                    m.dLastRefreshMesage = gdTickCount
                End If
                AddPositionToBrokerInfo strMessage, False
                HandlePositionInfo strMessage
            End If
        
        Case eGDLindXpressMessageType_Alert
            m.dHeartBeat = gdTickCount
            HandleAlert strMessage
            
        Case eGDLindXpressMessageType_AccountInfo
            m.dHeartBeat = gdTickCount
            g.Broker.UpdateAccountInfoForm Parse(strMessage, vbTab, 1), strMessage
            
    End Select

    If Refreshing And (m.dLastRefreshMesage = 0#) Then
        m.dLastRefreshMesage = gdTickCount
    ElseIf Refreshing And (gdTickCount > m.dLastRefreshMesage + 30000#) Then
        m.bRefreshingAccounts = False
        m.bRefreshingOrders = False
        m.bRefreshingFills = False
        m.bRefreshingPositions = False
        If m.bVerboseRefresh Then
            InfBox ""
            m.bVerboseRefresh = False
        End If
    ElseIf (Not Refreshing) And (m.dLastRefreshMesage <> 0#) Then
        m.dLastRefreshMesage = 0#
        If m.bVerboseRefresh Then
            InfBox "Synchronizing " & m.strBrokerName & " position information...", , , m.strBrokerName & " Refresh", True
        End If
        SynchronizePositions
        If m.bVerboseRefresh Then
            InfBox ""
            m.bVerboseRefresh = False
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.HandleMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckHeartbeat
'' Description: Check the heartbeat to make sure that we are still connected
''              to the LindXpress servers
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
                DumpDebug "Communication has been lost with " & m.strBrokerName & ".  Attempting to reconnect..."
            End If
            Reconnect
        ElseIf (ConnectionStatus = eGDConnectionStatus_Connected) Then
            If m.lReconnectAttempts > 1& Then
                InfBox "TradeNavigator has reconnected to " & m.strBrokerName & ".", , "+-OK", m.strBrokerName & " Connection", True
                DumpDebug "TradeNavigator has reconnected to " & m.strBrokerName & "."
            End If
            m.lReconnectAttempts = 0&
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.CheckHeartbeat"
    
End Sub

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
    RaiseError "cXpress.DumpDebug", eGDRaiseError_Raise
    
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
        If ((m.bVerifyPositions = True) Or (bForce = True)) And (m.bFixCalledOnce = True) Then
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
    RaiseError "cXpress.VerifyPositions"
    
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

    IsTradeableSymbol = (Len(LindXpressSymbol(strGenesisSymbol)) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.IsTradeableSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsEnabledSymbol
'' Description: Is the given symbol enabled for trading for the user?
'' Inputs:      Symbol, Broker Base, Broker Exchange
'' Returns:     True if enabled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEnabledSymbol(ByVal strGenesisSymbol As String, Optional strBrokerBase As String, Optional strBrokerExchange As String) As Boolean
On Error GoTo ErrSection:

    IsEnabledSymbol = (Len(LindXpressSymbol(strGenesisSymbol)) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.IsEnabledSymbol"
    
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
                
                Exit For
            End If
        Next lIndex
    End If
    
    OrderTypeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.OrderTypeAllowed"
    
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
            If Len(strMask) = 0 Then strMask = "100"
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
    RaiseError "cXpress.TimeInForceAllowed"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.FixPositions"
    
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
        End If
    End If
    
    FixPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.FixPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolInformation
'' Description: Get the symbol information for the given symbol
'' Inputs:      Genesis Symbol
'' Returns:     Symbol Information
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SymbolInformation(ByVal strGenesisSymbol As String) As String
On Error GoTo ErrSection:

    Dim strBase As String               ' Base symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If InStr(strGenesisSymbol, " ") = 0 Then
        strBase = Parse(strGenesisSymbol, "-", 1)
    Else
        strBase = "O:" & Parse(Parse(strGenesisSymbol, " ", 1), "-", 1)
    End If
    
    For lIndex = 0 To m.astrSymbols.Size - 1
        If Parse(m.astrSymbols(lIndex), vbTab, 3) = strBase Then
            strReturn = Parse(m.astrSymbols(lIndex), vbTab, 5) & vbTab & Parse(m.astrSymbols(lIndex), vbTab, 6)
            Exit For
        End If
    Next lIndex
    
    SymbolInformation = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.SymbolInformation"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangePassword
'' Description: Attempt to change the password on the Xpress servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangePassword()
On Error GoTo ErrSection:

    Dim strNewPassword As String        ' New Password

    strNewPassword = frmPassword.ShowChange(m.strPassword)
    If Len(strNewPassword) > 0 Then
        SendLindXpressMessage eGDLindXpressMessageType_ChangePassword, m.strPassword & vbTab & strNewPassword
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.ChangePassword"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleConnectionInfo
'' Description: Handle the connection information
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
    Dim frm As frmAlertPopup            ' Alert popup message
    Static strLastUser As String        ' Last user name
    
    dTickCount = Val(Parse(strConnectionInfo, vbTab, 1))
    If Len(Parse(strConnectionInfo, vbTab, 2)) > 0 Then
        nStatus = CLng(Val(Parse(strConnectionInfo, vbTab, 2)))
        strError = Parse(strConnectionInfo, vbTab, 3)
    
        Select Case nStatus
            Case eGDConnectionStatus_Disconnected
                m.BrokerInfo.SetConnectionStatus eGDConnectionStatus_Disconnected
            
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
                    SendLindXpressMessage eGDLindXpressMessageType_UnloadApp, ""
                End If
            
            Case eGDConnectionStatus_Disconnecting
                m.BrokerInfo.SetConnectionStatus eGDConnectionStatus_Disconnecting
                m.dLastStatusMsg = dTickCount
                
            Case eGDConnectionStatus_Connecting
                m.BrokerInfo.SetConnectionStatus eGDConnectionStatus_Connecting
                m.dLastStatusMsg = dTickCount
                
            Case eGDConnectionStatus_Connected
                m.dLastStatusMsg = dTickCount
                m.bWasConnected = True
                
        End Select
    
        If m.strUserName <> strLastUser Then
            strLastUser = m.strUserName
        End If
        
        If ConnectionStatus <> nStatus Then
            nPrev = ConnectionStatus
            ConnectionStatus = nStatus
            
            g.Broker.ConnectionStatusChanged m.nBroker, nStatus
        
            If (nPrev <> eGDConnectionStatus_Connected) And (ConnectionStatus = eGDConnectionStatus_Connected) Then
                SetIniFileProperty "Last", EncryptToHex(Str(Date)), "Connect", m.strIniFile
                GetAccounts
            End If
            DumpDebug vbTab & m.strBrokerName & " Message Received (" & MessageTypeToString(eGDLindXpressMessageType_ConnectionInfo) & "): " & strConnectionInfo
        End If
        
        If (Len(strError) > 0) And (strError <> m.strLastMessage) Then
            If InStr(UCase(strError), "INVALID LOGON") <> 0 Then
                m.bManualDisconnect = True
                m.dHeartBeat = 0#
                m.bVerifyPositions = True
                m.bAskedPassword = False
            End If
            If (InStr(UCase(strError), "ERROR 6:") <> 0) Then
                ' This is a generic socket error - set the reconnect attempts to 1 so that the
                ' "you have lost connection" dialog comes up on the first actual reconnect attempt...
                m.lReconnectAttempts = 1&
            ElseIf (InStr(UCase(strError), "ERROR 7:") = 0) Then
                Set frm = New frmAlertPopup
                frm.ShowMessageBox strError, m.strBrokerName & " Message", vbCenter
            End If
        End If
        m.strLastMessage = strError
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.HandleConnectionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleAccountInfo
'' Description: Handle account information back from the Xpress servers
'' Inputs:      Account Info
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleAccountInfo(ByVal strAccountInfo As String)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strAccountNumber As String      ' Account Number from the callback
    Dim lAccountID As Long              ' Account ID
    Dim lPos As Long                    ' Position of account in the array
    
    If Len(strAccountInfo) > 0 Then
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
            rs!UserName = Login
            EncryptField rs!Password, m.strPassword
            rs!AccountType = m.nBroker
            rs!Broker = ""
            rs!Comms = 0#
            rs!SecTypeMask = 9
            rs!FillRT = False
            rs!ClosedProfit = 0#
            lAccountID = rs!AccountID
            rs.Update
            
            ' If we had to create an account, reset the fill and refresh times so
            ' that we get full refresh of a week or two of data (10/09/2008 DAJ)...
            SetIniFileProperty "Fill", 0#, "Refresh", m.strIniFile
            SetIniFileProperty "Order", 0#, "Refresh", m.strIniFile
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
    RaiseError "cXpress.HandleAccountInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleOrderInfo
'' Description: Handle order information back from the LindXpress servers
'' Inputs:      Order Info, Has Fill Info
'' Returns:     None
''
'' RecordFrom:  Genesis ID, Broker ID, Original ID, Status, Account, Order Time,
''              Business Date, Error Code, Error Text, Symbol, B/S, Order Type,
''              Price, With Limit Price, Quantity, Qty Filled, Qty Working,
''              Expiration, Floor Recv Time, Broker Recv Time, Broker Exec Time,
''              Fill ID, Trade Date, Trade Price, Fill Qty, Snapshot
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleOrderInfo(ByVal strOrderInfo As String, ByVal bHasFillInfo As Boolean, Optional ByVal bRefresh As Boolean = False)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    Dim Fill As cPtFill                 ' Fill object
    Dim astrOrder As New cGdArray       ' Array of order information from callback
    Dim strLindXpressSymbol As String   ' LindXpress symbol
    Dim strGenesisSymbol As String      ' Genesis version of LindXpress symbol
    Dim dMult As Double                 ' Price multiplier
    Dim nPrevStatus As eTT_OrderStatus  ' Previous order status for the order
    Dim lIndex As Long                  ' Index into a for loop
    Dim strFillID As String             ' Fill identifier
    Dim bFillExisted As Boolean         ' Did the fill already exist?
    Dim frm As frmAlertPopup            ' Alert popup message
    Dim strTradeDate As String          ' Trading date
    Dim bReplace As Boolean             ' Is this a special replace message?
    Dim lPos As Long                    ' Position in an array
    Dim strGenesisID As String          ' Genesis Order ID
    Dim Bars As New cGdBars             ' Bars object
    Dim bNewOrder As Boolean            ' Was this a new order?
    Dim nOrderType As eTT_OrderType     ' Order Type
    Dim bSnapshot As Boolean            ' Is this a snapshot order/fill?
    Dim PrevOrder As cPtOrder           ' Previous order

    astrOrder.SplitFields strOrderInfo, vbTab
    
    bReplace = False
    strLindXpressSymbol = astrOrder(9)
    If UCase(astrOrder(0)) = "REPLACE" Then
        astrOrder(0) = ""
        bReplace = True
    ElseIf UCase(astrOrder(0)) = "SPECIAL" Then
        astrOrder(0) = ""
    End If
    
    ' If the broker ID comes back as zero from Xpress, blank it out...
    If astrOrder(1) = "0" Then
        astrOrder(1) = ""
    End If
    
    If astrOrder(2) = "0" Then
        astrOrder(2) = ""
    End If
    
    If Len(astrOrder(5)) >= 8 Then
        astrOrder(1) = Left(astrOrder(5), 8) & "-" & astrOrder(1)
        If Len(astrOrder(2)) > 0 Then
            astrOrder(2) = Left(astrOrder(5), 8) & "-" & astrOrder(2)
        End If
    End If
    
    If (Len(astrOrder(25)) > 0) And (Val(astrOrder(25)) = 0) Then
        bSnapshot = False
    Else
        bSnapshot = True
    End If
    
    If Len(strLindXpressSymbol) = 0 Then
        If bReplace Then
            DumpDebug "Searching for Amend Pending Info for " & astrOrder(2)
            If m.astrAmendPending.BinarySearch(astrOrder(2) & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
                DumpDebug "Amend Pending Info Found: " & m.astrAmendPending(lPos)
                
                Set Order = LoadOrder(astrOrder(4), Parse(m.astrAmendPending(lPos), vbTab, 3), astrOrder(2), "")
                If Not Order Is Nothing Then
                    Order.Status = GenesisOrderStatus(astrOrder(3))
                    If Len(astrOrder(5)) > 0 Then
                        Order.StatusDate = DateToDouble(astrOrder(5))
                    Else
                        Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
                    End If
                    Order.Save
                    
                    m.BrokerInfo.AddOrder Order, bRefresh, bSnapshot
                    
                    OrderCallback Order
                    If bRefresh = False Then
                        g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                    End If
                
                    DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                    SendOrderToOptionNav Order, bRefresh
                End If
                
                Set Order = LoadOrder(astrOrder(4), Parse(m.astrAmendPending(lPos), vbTab, 5), astrOrder(2), "")
                If Not Order Is Nothing Then
                    Order.BrokerID = astrOrder(1)
                    Order.Status = eTT_OrderStatus_Sent
                    If Len(astrOrder(5)) > 0 Then
                        Order.StatusDate = DateToDouble(astrOrder(5))
                    Else
                        Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
                    End If
                    Order.Save
                    
                    m.BrokerInfo.AddOrder Order, bRefresh, bSnapshot
                    
                    OrderCallback Order
                    If bRefresh = False Then
                        g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                    End If
                
                    DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                    SendOrderToOptionNav Order, bRefresh
                End If
                
                OrderIDChanged CLng(Val(Parse(m.astrAmendPending(lPos), vbTab, 2))), CLng(Val(Parse(m.astrAmendPending(lPos), vbTab, 4)))
                DumpDebug "Removing Amend Pending Info: " & m.astrAmendPending(lPos)
                m.astrAmendPending.Remove lPos
            End If
        Else
            Set Order = LoadOrder(astrOrder(4), astrOrder(0), astrOrder(1), astrOrder(2))
            If Not Order Is Nothing Then
                With Order
                    nPrevStatus = .Status
                    .Status = GenesisOrderStatus(astrOrder(3))
                    If (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_ParkPending) Then
                        .Status = eTT_OrderStatus_Parked
                    ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_Parked) Then
                        .Status = eTT_OrderStatus_Parked
                    ElseIf (.Status = eTT_OrderStatus_Cancelled) And (Order.Fills.Count > 0) Then
                        ' If Max Express is calling this order cancelled, but it has fills associated
                        ' with it, then call it Balance Cancelled (01/09/2009 DAJ)...
                        .Status = eTT_OrderStatus_BalCancelled
                    End If
                    If Len(astrOrder(5)) > 0 Then
                        Order.StatusDate = DateToDouble(astrOrder(5))
                    Else
                        Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
                    End If
                    .Save
                    
                    m.BrokerInfo.AddOrder Order, bRefresh, bSnapshot
                    
                    OrderCallback Order
                    If bRefresh = False Then
                        g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                    End If
                
                    DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                    SendOrderToOptionNav Order, bRefresh
                End With
            End If
        End If
    ElseIf GenesisSymbol(strLindXpressSymbol, strGenesisSymbol, dMult) Then
        SetBarProperties Bars, strGenesisSymbol
        Set Order = LoadOrder(astrOrder(4), astrOrder(0), astrOrder(1), "") ', astrOrder(2))
        If Order Is Nothing Then
            DumpDebug "Order " & astrOrder(1) & ";" & astrOrder(0) & " not found -- creating new order"
            Set Order = New cPtOrder
            bNewOrder = True
        Else
            DumpDebug "Order " & astrOrder(1) & ";" & astrOrder(0) & " found"
            bNewOrder = False
        End If
        
        With Order
            nPrevStatus = .Status
            
            .AccountID = g.Broker.AccountIDForNumber(astrOrder(4))
            .BrokerID = astrOrder(1)
            Select Case astrOrder(10)
                Case "1"
                    .Buy = True
                Case "2"
                    .Buy = False
            End Select
            .DateSent = DateToDouble(astrOrder(5))
            .DateHostRecd = DateToDouble(astrOrder(6))
            '.ExchangeID = astrOrder(2)
            If Len(Trim(astrOrder(17))) = 0 Then
                .Expiration = CLng(.DateHostRecd) * -1&
            ElseIf UCase(Trim(astrOrder(17))) = "CANCEL" Then
                .Expiration = 0&
            Else
                .Expiration = JulFromLong(CLng(Val(Left(astrOrder(17), 8))))
            End If
            .GenesisOrderID = astrOrder(0)
            .Message = astrOrder(7)
            If .OrderDate = 0 Then Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
            
            nOrderType = GenesisOrderType(astrOrder(11))
            If Not ((nOrderType = eTT_OrderType_Market) And (.OrderType = eTT_OrderType_MIT)) Then
                .OrderType = nOrderType
                Select Case .OrderType
                    Case eTT_OrderType_Market
                        .LimitPrice = 0#
                        .StopPrice = 0#
                    Case eTT_OrderType_Stop
                        .LimitPrice = 0#
                        .StopPrice = GenesisPrice(Val(astrOrder(12)), dMult)
                    Case eTT_OrderType_Limit
                        .LimitPrice = GenesisPrice(Val(astrOrder(12)), dMult)
                        .StopPrice = 0#
                    Case eTT_OrderType_StopWithLimit
                        .LimitPrice = GenesisPrice(Val(astrOrder(13)), dMult)
                        .StopPrice = GenesisPrice(Val(astrOrder(12)), dMult)
                End Select
            End If
            .SymbolOrSymbolID = strGenesisSymbol
            
            .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.OrderDate, m.nBroker, .Symbol, False))
            .IsSnapshot = True
            .PreviousBrokerID = astrOrder(2)
            
            If (Len(.PreviousBrokerID) > 0) And (.PreviousBrokerID <> .BrokerID) Then
                Set PrevOrder = New cPtOrder
                PrevOrder.LoadForBrokerID .PreviousBrokerID
                .PreviousOrder = PrevOrder
            End If
                            
            If (bHasFillInfo = False) Or (.OrderID = 0) Then .Save
            
            If bHasFillInfo Then
                strFillID = astrOrder(1) & "-" & astrOrder(21)
                bFillExisted = False
                For lIndex = 1 To .Fills.Count
                    If (.Fills(lIndex).BrokerID = strFillID) Then
                        Set Fill = .Fills(lIndex)
                        bFillExisted = True
                        Exit For
                    End If
                Next lIndex
                If bFillExisted = False Then Set Fill = New cPtFill
                
                If Val(astrOrder(24)) > 0 Then
                    With Fill
                        .AccountID = Order.AccountID
                        .AutoTradingItemID = Order.AutoTradeItemID
                        .BrokerID = strFillID
                        .BrokerOrderID = Order.BrokerID
                        .Buy = Order.Buy
                        .FillDate = DateToDouble(astrOrder(20))
                        .IsManual = False
                        .IsSnapshot = True
                        .OrderID = Order.OrderID
                        .Price = GenesisPrice(Val(astrOrder(23)), dMult)
                        .Quantity = CLng(Val(astrOrder(24)))
                        .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.FillDate, m.nBroker, .Symbol, False))
                        .Symbol = Order.Symbol
                        .SymbolID = Order.SymbolID
                        
                        If .FillID = 0& Then .Save
                    End With
                    
                    Set .Fills(Str(Fill.FillID)) = Fill
                Else
                    m.BrokerInfo.RemoveFill Fill
                    Order.Fills.Remove Str(Fill.FillID)
                    g.ActivityLogs.AddFillToActivityLog m.nBroker, Fill, True
                End If
            
                DumpDebug Fill.FillText
            End If
            
            ' Since modify orders deal with remaining quantities, we need to add the
            ' remaining quantity to the fill quantity to get the absolute quantity...
            .Quantity = Val(astrOrder(16)) + .FillQuantity
            
            ' If the quantity ended up being zero, we might not have the fills yet, so
            ' use the order quantity that they gave us...
            If (.Quantity = 0&) And (Val(astrOrder(14)) > 0) Then
                .Quantity = CLng(Val(astrOrder(14)))
            End If
            
            'If bHasFillInfo = False Then
                .Status = GenesisOrderStatus(astrOrder(3))
                .StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, .Symbol)
            'End If
            
            If (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_ParkPending) Then
                .Status = eTT_OrderStatus_Parked
            ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_Parked) And (bNewOrder = False) Then
                .Status = eTT_OrderStatus_Parked
            ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_Amended) Then
                .Status = eTT_OrderStatus_Amended
            ElseIf (Order.FillQuantity > Order.Quantity) Then
                .Status = eTT_OrderStatus_OverFilled
            ElseIf (.Status = eTT_OrderStatus_Filled) And (Order.FillQuantity = 0) Then
                .Status = eTT_OrderStatus_Working
            ElseIf (.Status = eTT_OrderStatus_Filled) And (Order.FillQuantity < Order.Quantity) Then
                ' DAJ 01/09/2009: If this order was currently marked as Cancelled or Balance Cancelled,
                ' but Man Express is now calling it filled but it is not fully filled, call it Balance
                ' Cancelled...
                ' DAJ 02/22/2010: Also call is Balance Cancelled if the order was currently marked
                ' as ameneded since the remainder will not fill...
                If (nPrevStatus = eTT_OrderStatus_Cancelled) Or (nPrevStatus = eTT_OrderStatus_BalCancelled) Or (nPrevStatus = eTT_OrderStatus_Amended) Then
                    .Status = eTT_OrderStatus_BalCancelled
                Else
                    .Status = eTT_OrderStatus_Partial
                End If
            ElseIf (.Status = eTT_OrderStatus_Partial) And (Order.FillQuantity = Order.Quantity) Then
                .Status = eTT_OrderStatus_Filled
            ElseIf (.Status = eTT_OrderStatus_Cancelled) And (Order.FillQuantity > 0&) Then
                .Status = eTT_OrderStatus_BalCancelled
            ElseIf (.Status = eTT_OrderStatus_Working) And (Order.FillQuantity > 0&) Then
                If Order.FillQuantity = Order.Quantity Then
                    .Status = eTT_OrderStatus_Filled
                Else
                    .Status = eTT_OrderStatus_Partial
                End If
            End If
            
            .Save
            
            ' If we are not refreshing and the order status just went to Error or Rejected, then
            ' display a message to the user (05/08/2008 DAJ)...
            If bRefresh = False Then
                If (nPrevStatus <> Order.Status) And ((Order.Status = eTT_OrderStatus_Error) Or (Order.Status = eTT_OrderStatus_Rejected)) Then
                    If Len(astrOrder(8)) > 0 Then
                        If (Len(astrOrder(1)) > 0) And (astrOrder(1) <> "0") Then
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox "Please contact your broker and report the following problem:||Order ID: " & astrOrder(1) & "|" & "Error: " & astrOrder(8) & "|", m.strBrokerName & " Order Error", vbLeftJustify
                        Else
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox "Please contact your broker and report the following problem:||Error: " & astrOrder(8) & "|", m.strBrokerName & " Order Error", vbLeftJustify
                        End If
                    Else
                        If (Len(astrOrder(1)) > 0) And (astrOrder(1) <> "0") Then
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox "Please contact your broker and report the following problem:||Order ID " & astrOrder(1) & " was rejected.|", m.strBrokerName & " Order Error", vbLeftJustify
                        Else
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox "Please contact your broker and report the following problem:||Order was rejected.|", m.strBrokerName & " Order Error", vbLeftJustify
                        End If
                    End If
                End If
            End If
            
            If bHasFillInfo = True And Val(astrOrder(24)) > 0 Then
                m.BrokerInfo.AddOrder Order, bRefresh, bSnapshot
                m.BrokerInfo.AddFill Fill, bRefresh, , bSnapshot
            
                FillCallback Order, Fill, bFillExisted
                If bRefresh = False Then
                    g.ActivityLogs.AddFillToActivityLog m.nBroker, Fill, False
                End If
                SendFillToOptionNav Fill, bRefresh
                
                ' Don't want to resynchronize positions on each fill of a refresh because it will be done once
                ' at the end (05/08/2008 DAJ)...
                If (bFillExisted = False) And (bRefresh = False) Then
                    SynchronizePositions
                End If
            Else
                m.BrokerInfo.AddOrder Order, bRefresh, bSnapshot
            
                OrderCallback Order, False
                If bRefresh = False Then
                    g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                End If
                SendOrderToOptionNav Order, bRefresh
            End If
            
            If Order.Status = eTT_OrderStatus_OverFilled Then
                Set frm = New frmAlertPopup
                frm.ShowMessageBox "Trade Navigator has received duplicate fill|information for order " & Order.BrokerID & ".||This will cause this order to appear over-filled and will likely result in a position mismatch.||PLEASE CALL YOUR BROKER AND VERIFY YOUR POSITIONS.", "Over Filled Order on " & m.strBrokerName, vbCenter
            End If

            DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
            
            ' In case the customer is trading on multiple machines, we need to make sure that the
            ' date and counter are up to date in the INI file so we don't duplicate Genesis ID's...
            UpdateLastGenesisID astrOrder(0)
            
            ' If the original order ID is different than the order ID (the order has been modified),
            ' then make sure that the original order is set to Amended...
            If (astrOrder(1) <> astrOrder(2)) And (Len(astrOrder(2)) > 0) Then
                Set Order = LoadOrder(astrOrder(4), "", astrOrder(2), "")
                If Not Order Is Nothing Then
                    If (Order.Status <> eTT_OrderStatus_Amended) Then
                        Order.Status = eTT_OrderStatus_Amended
                        Order.Save
                        
                        m.BrokerInfo.AddOrder Order, bRefresh, bSnapshot
                        
                        OrderCallback Order
                        If bRefresh = False Then
                            g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                        End If
                    
                        DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                        SendOrderToOptionNav Order, bRefresh
                    End If
                End If
            End If
        End With
    End If

    If bRefresh = False Then
        SendNowToOptionNav
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.HandleOrderInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandlePositionInfo
'' Description: Handle a position from a callback
'' Inputs:      Position String
'' Returns:     None
''
'' RecordFrom:  Account, Symbol, B/S, Quantity, Currency, Desc, Firm,
''              Interface Type, Net Profit, Price, Price Kind, P/C, Record Type,
''              Reg Code, Seq Number, Settlement Price, Strike Price, Tick,
''              Trade Date, Validate, Version
'' RecordTo:    Account, Symbol, Current Position, Carried Position, Average
''              Entry, Entry Prices
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandlePositionInfo(ByVal strPosition As String)
On Error GoTo ErrSection:

    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim dMult As Double                 ' Price multiplier
    Dim lPos As Long                    ' Position in the array
    Dim adEntryPrices As cGdArray       ' Array of entry prices
    Dim astrPosition As cGdArray        ' Array of position information
    Dim lLindXpressPosition As Long     ' LindXpress position
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrBrokerPos As cGdArray       ' Broker position information
    Dim dAverageEntry As Double         ' Average entry price

    Set astrPosition = New cGdArray
    astrPosition.Create eGDARRAY_Strings
    Set adEntryPrices = New cGdArray
    adEntryPrices.Create eGDARRAY_Doubles
    Set astrBrokerPos = New cGdArray
    astrBrokerPos.Create eGDARRAY_Strings
    
    If Len(strPosition) > 0 Then
        astrPosition.SplitFields strPosition, vbTab
        
        If GenesisSymbolForPosition(astrPosition(1), strGenesisSymbol, dMult) Then
            Select Case UCase(astrPosition(2))
                Case "1"
                    lLindXpressPosition = CLng(Val(astrPosition(3)))
                Case "2"
                    lLindXpressPosition = CLng(Val(astrPosition(3))) * -1&
            End Select
        
            If m.astrBrokerPos.BinarySearch(astrPosition(0) & vbTab & strGenesisSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
                astrBrokerPos.SplitFields m.astrBrokerPos(lPos), vbTab
                
                adEntryPrices.SplitFields astrBrokerPos(5), ","
                For lIndex = 1 To Abs(lLindXpressPosition)
                    adEntryPrices.Add GenesisPrice(astrPosition(9), dMult)
                Next lIndex
                lLindXpressPosition = lLindXpressPosition + CLng(Val(astrBrokerPos(3)))
                
                m.astrBrokerPos(lPos) = astrBrokerPos(0) & vbTab & astrBrokerPos(1) & vbTab & astrBrokerPos(2) & vbTab & Str(lLindXpressPosition) & vbTab & Str(adEntryPrices.CalcStatistic(eGdStat_Average)) & vbTab & adEntryPrices.JoinFields(",")
                
                DumpDebug "Position Updated: " & m.astrBrokerPos(lPos)
            Else
                adEntryPrices.Clear
                For lIndex = 1 To Abs(lLindXpressPosition)
                    adEntryPrices.Add GenesisPrice(astrPosition(9), dMult)
                Next lIndex
                m.astrBrokerPos.Add astrPosition(0) & vbTab & strGenesisSymbol & vbTab & "" & vbTab & Str(lLindXpressPosition) & vbTab & Str(adEntryPrices.CalcStatistic(eGdStat_Average)) & vbTab & adEntryPrices.JoinFields(","), lPos
                
                DumpDebug "New Position Added: " & m.astrBrokerPos(lPos)
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.HandlePositionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SynchronizePositions
'' Description: Synchronize the Genesis positions with the LindXpress positions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SynchronizePositions(Optional ByVal strSymbol As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPosition As Long               ' Position from the broker
    Dim strPosition As String           ' Position information
    Dim lSymbolID As Long               ' Symbol ID
    Dim frm As frmAlertPopup            ' Alert popup form
    Dim FillSumms As cAccountPositions  ' Fill summaries
    Dim strAccount As String            ' Account number
    Dim lBrokerPos As Long              ' Broker position
    Dim lBrokerCarried As Long          ' Broker carried position
    Dim lPos As Long                    ' Position in an array
    
    m.bSyncInProgress = True
    
    BuildPositions
    frmPositionConfirm.ShowMe m.nBroker, "", strSymbol, False, False
    
    SendBeginToOptionNav eGDOptNav_Position, m.nBroker
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
                        ' If the last refresh was just a snapshot refresh, try to get a historical
                        ' refresh and see if that fixes the carried position mismatch (10/09/2008 DAJ)...
                        If m.lLastFromDateFills = Int(CurrentTime("CHI")) Then
                            DumpDebug "Calling for a big refresh since we have a carried position mismatch"
                            
                            SetIniFileProperty "Fill", 0#, "Refresh", m.strIniFile
                            SetIniFileProperty "Order", 0#, "Refresh", m.strIniFile
                            
                            Refresh
                            Exit For
                        End If
                        
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
                        
                        SendPositionToOptionNav FillSumms(lIndex), True
                    End If
                End If
            End With
        Next lIndex
    End If

    SendEndToOptionNav eGDOptNav_Position, m.nBroker, True
    
    
    m.bSyncInProgress = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.SynchronizePositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleAlert
'' Description: Handle alert back from the LindXpress servers
'' Inputs:      Alert
'' Returns:     None
''
'' Record:      Identifier, Order Number, Alert Type, Error Code, Error Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleAlert(ByVal strAlert As String)
On Error GoTo ErrSection:

    Dim astrAlert As New cGdArray       ' Alert broken out into an array
    Dim frm As New frmAlertPopup        ' Alert popup box
    Dim strMessage As String            ' Message to display to user
    Dim Order As cPtOrder               ' Order for the alert
    
    astrAlert.SplitFields strAlert, vbTab
    
    Select Case astrAlert(2)
        Case "0"                        ' AT_NONE
        Case "1"                        ' AT_INFORMATION
            frm.ShowMessageBox astrAlert(4), "Broker Message", vbCenter
            g.ActivityLogs.AddActivityLogEntry m.nBroker, astrAlert(4)
            
        Case "2"                        ' AT_ERROR
            strMessage = "Order #" & astrAlert(1) & " has the following error:|" & astrAlert(3) & ": " & astrAlert(4)
            frm.ShowMessageBox strMessage, "Order Error", vbCenter
            g.ActivityLogs.AddActivityLogEntry m.nBroker, strMessage
        
        Case "3"                        ' AT_BROKER_REJECT
            strMessage = "Order #" & astrAlert(1) & " has been rejected by the broker"
            If Len(astrAlert(4)) > 0 Then
                strMessage = strMessage & "||" & astrAlert(4)
            End If
            frm.ShowMessageBox strMessage, "Broker Rejection", vbCenter
            g.ActivityLogs.AddActivityLogEntry m.nBroker, strMessage
            Set Order = OrderForAlert(astrAlert(0), astrAlert(1))
            If Not Order Is Nothing Then
                Order.Status = eTT_OrderStatus_Rejected
                Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
                Order.Message = astrAlert(4)
                Order.Save
            
                m.BrokerInfo.AddOrder Order, False
                
                OrderCallback Order
                g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
            
                DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                SendOrderToOptionNav Order, False
            End If
        
        Case "4"                        ' AT_CANCEL_CONFIRMATION
            If m.bConfirmCancel = True Then
                strMessage = "Order #" & astrAlert(1) & " has been confirmed Cancelled"
                frm.ShowMessageBox strMessage, "Cancel Confirmation", vbCenter
                g.ActivityLogs.AddActivityLogEntry m.nBroker, strMessage
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.HandleAlert"
    
End Sub

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
'' Function:    LindXpressSymbol
'' Description: Convert a Genesis symbol to a LindXpress symbol
'' Inputs:      Genesis Symbol, Multiplier
'' Returns:     LindXpress Symbol (blank if not converted)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LindXpressSymbol(ByVal strGenesisSymbol As String, Optional dMult As Double) As String
On Error GoTo ErrSection:

    Dim strBase As String               ' Base symbol
    Dim strContract As String           ' Contract
    Dim strStrike As String             ' Strike price
    Dim strPutOrCall As String          ' Put or call
    Dim lIndex As Long                  ' Index into a for loop
    Dim strLindXpressSymbol As String   ' Symbol to return
    Dim dStrikeMult As Double           ' Strike price multiplier
    Dim dStrike As Double               ' Strike price
    
    strLindXpressSymbol = ""
    If InStr(strGenesisSymbol, " ") = 0 Then
        strBase = Parse(strGenesisSymbol, "-", 1)
        strContract = Parse(strGenesisSymbol, "-", 2)
        strStrike = ""
        strPutOrCall = ""
    Else
        strBase = "O:" & Parse(Parse(strGenesisSymbol, " ", 1), "-", 1)
        strContract = Parse(Parse(strGenesisSymbol, " ", 1), "-", 2)
        strStrike = Parse(strGenesisSymbol, " ", 2)
        strPutOrCall = Left(strStrike, 1)
        strStrike = Right(strStrike, Len(strStrike) - 1)
    End If
    
    For lIndex = 0 To m.astrSymbols.Size - 1
        If Parse(m.astrSymbols(lIndex), vbTab, 3) = strBase Then
            strLindXpressSymbol = Parse(m.astrSymbols(lIndex), vbTab, 1) & "|" & Mid(strContract, 2, 3) & "|" & Right(strContract, 2)
            dMult = Val(Parse(m.astrSymbols(lIndex), vbTab, 4))
            If Len(strStrike) > 0 Then
                dStrikeMult = Val(Parse(m.astrSymbols(lIndex), vbTab, 9))
                dStrike = Val(strStrike) * dStrikeMult
                strLindXpressSymbol = strLindXpressSymbol & "|" & Format(Int(RoundToSigDigits(dStrike)), "#") & "|" & strPutOrCall
            Else
                strLindXpressSymbol = strLindXpressSymbol & "|" & "|" & strPutOrCall
            End If
            If Left(strLindXpressSymbol, 2) = "O:" Then
                strLindXpressSymbol = Mid(strLindXpressSymbol, 3)
            End If
            Exit For
        End If
    Next lIndex
    
    LindXpressSymbol = strLindXpressSymbol

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.LindXpressSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LindXpressPrice
'' Description: Convert a Genesis price to a LindXpress price
'' Inputs:      Genesis Price, Multiplier
'' Returns:     LindXpress Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LindXpressPrice(ByVal dGenesisPrice As Double, ByVal dMult As Double) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value from the function

    Select Case dMult
        Case 8
            dReturn = CDbl(Int(dGenesisPrice) * 10) + CDbl((dGenesisPrice - Int(dGenesisPrice)) * 8)
        Case 16
            dReturn = CDbl(Int(dGenesisPrice) * 100) + CDbl((dGenesisPrice - Int(dGenesisPrice)) * 16)
        Case 32
            dReturn = CDbl(Int(dGenesisPrice) * 100) + CDbl((dGenesisPrice - Int(dGenesisPrice)) * 32)
        Case 32.25
            dReturn = CDbl(Int(Int(dGenesisPrice) * 1000) + CDbl(Int(RoundNum((dGenesisPrice - Int(dGenesisPrice)) * 320, 1))))
        Case 32.5
            dReturn = CDbl(Int(dGenesisPrice) * 1000) + CDbl((dGenesisPrice - Int(dGenesisPrice)) * 320)
        Case 64
            dReturn = CDbl(Int(dGenesisPrice) * 100) + CDbl((dGenesisPrice - Int(dGenesisPrice)) * 64)
        Case 64.5
            dReturn = CDbl(Int(dGenesisPrice) * 1000) + CDbl((dGenesisPrice - Int(dGenesisPrice)) * 640)
        Case 100.5
            dReturn = CDbl(Int(RoundNum(dGenesisPrice * 100, 1)))
        Case 1000.5
            dReturn = CDbl(Int(RoundNum(dGenesisPrice * 1000, 1)))
        Case Else
            dReturn = dGenesisPrice * dMult
    End Select
    
    LindXpressPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.LindXpressPrice"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LindXpressOrderType
'' Description: Convert a Genesis order type to a LindXpress order type
'' Inputs:      Genesis Order Type
'' Returns:     LindXpress Order Type (blank if not converted)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LindXpressOrderType(ByVal nOrderType As eTT_OrderType) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from the function
    
    strReturn = ""
    Select Case nOrderType
        Case eTT_OrderType_Market
            strReturn = "1"
        Case eTT_OrderType_Stop
            strReturn = "9"
        Case eTT_OrderType_Limit
            strReturn = "2"
        Case eTT_OrderType_StopWithLimit
            strReturn = "12"
        Case eTT_OrderType_MarketOnClose
            strReturn = "4"
        Case eTT_OrderType_StopCloseOnly
            strReturn = "10"
    End Select
    
    LindXpressOrderType = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.LindXpressOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExpirationDateForSymbol
'' Description: Given the symbol and the time of day, figure out the "proper"
''              default expiration date
'' Inputs:      None
'' Returns:     Default Expiration Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ExpirationDateForSymbol(ByVal strSymbol As String, ByVal lSymbolID As Long) As Double
On Error GoTo ErrSection:

    Dim strExchangeTimeInf As String    ' String of exchange time zone information
    Dim dNowInExchangeTime As Double    ' The current time in exchange time zone
    Dim dSessionEnd As Double           ' Session End time in exchange time zone
    Dim lReturn As Long                 ' Return value
    Dim Bars As New cGdBars             ' Temporary bars object
    Dim bReturn As Boolean              ' Return from the SetBarProperties call
    
    If lSymbolID = 0 Then
        bReturn = SetBarProperties(Bars, strSymbol)
    Else
        bReturn = SetBarProperties(Bars, lSymbolID)
    End If
    
    If (bReturn = True) Then
        strExchangeTimeInf = Bars.Prop(eBARS_ExchangeTimeZoneInf)
        dNowInExchangeTime = ConvertTimeZone(CurrentTime, , strExchangeTimeInf)
        dSessionEnd = Int(dNowInExchangeTime) + (Bars.Prop(eBARS_DefaultEndTime) / 1440#)
        
        If dNowInExchangeTime > dSessionEnd Then
            lReturn = Int(dNowInExchangeTime) + 1
            Do While Not IsWeekday(lReturn)
                lReturn = lReturn + 1
            Loop
        Else
            lReturn = Int(dNowInExchangeTime)
        End If
    Else
        lReturn = Date
    End If

    ExpirationDateForSymbol = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.ExpirationDateForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisPrice
'' Description: Convert a LindXpress price to a Genesis price
'' Inputs:      LindXpress Price, Multiplier
'' Returns:     Genesis Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisPrice(ByVal dLindXpressPrice As Double, ByVal dMult As Double) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value from the function

    Select Case dMult
        Case 0
            dReturn = 0#
        Case 8
            dReturn = CDbl(Int(dLindXpressPrice / 10)) + ((dLindXpressPrice Mod 10) / 8)
        Case 16
            dReturn = CDbl(Int(dLindXpressPrice / 100)) + ((dLindXpressPrice Mod 100) / 16)
        Case 32
            dReturn = CDbl(Int(dLindXpressPrice / 100)) + ((dLindXpressPrice Mod 100) / 32)
        Case 32.25
            If (dLindXpressPrice Mod 10 = 2) Or (dLindXpressPrice Mod 10 = 7) Then
                dReturn = CDbl(Int(dLindXpressPrice / 1000)) + (((dLindXpressPrice Mod 1000) + 0.5) / 320)
            Else
                dReturn = CDbl(Int(dLindXpressPrice / 1000)) + ((dLindXpressPrice Mod 1000) / 320)
            End If
        Case 32.5
            dReturn = CDbl(Int(dLindXpressPrice / 1000)) + ((dLindXpressPrice Mod 1000) / 320)
        Case 64
            dReturn = CDbl(Int(dLindXpressPrice / 100)) + ((dLindXpressPrice Mod 100) / 64)
        Case 64.5
            dReturn = CDbl(Int(dLindXpressPrice / 1000)) + ((dLindXpressPrice Mod 1000) / 640)
        Case 100.5
            dMult = 100
            If (dLindXpressPrice Mod 10 = 2) Or (dLindXpressPrice Mod 10 = 7) Then
                dReturn = (dLindXpressPrice + 0.5) / dMult
            Else
                dReturn = dLindXpressPrice / dMult
            End If
        Case 1000.5
            dMult = 1000
            If (dLindXpressPrice Mod 10 = 2) Or (dLindXpressPrice Mod 10 = 7) Then
                dReturn = (dLindXpressPrice + 0.5) / dMult
            Else
                dReturn = dLindXpressPrice / dMult
            End If
        Case Else
            dReturn = dLindXpressPrice / dMult
    End Select
    
    GenesisPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.GenesisPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Convert a LindXpress symbol to a Genesis symbol
'' Inputs:      LindXpress Symbol, Genesis Symbol, Multiplier
'' Returns:     True if converted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbol(ByVal strLindXpressSymbol As String, strGenesisSymbol As String, dMult As Double) As Boolean
On Error GoTo ErrSection:

    Dim strBase As String               ' Base symbol
    Dim lPos As Long                    ' Position of base symbol in the array
    Dim bReturn As Boolean              ' Return value from the function
    Dim strStrike As String             ' Strike price
    Dim strPutOrCall As String          ' Put or call
    Dim dStrikeMult As Double           ' Strike price multiplier
    Dim dStrike As Double               ' Strike price
    
    strBase = Parse(strLindXpressSymbol, "|", 1)
    strStrike = Parse(strLindXpressSymbol, "|", 4)
    If strStrike = "0" Then strStrike = ""
    strPutOrCall = Parse(strLindXpressSymbol, "|", 5)
    
    If (Len(strStrike) > 0) And (Len(strPutOrCall) > 0) Then
        strBase = "O:" & strBase
    End If
    
    If m.astrSymbols.BinarySearch(strBase & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        strGenesisSymbol = Parse(m.astrSymbols(lPos), vbTab, 3) & "-" & Parse(strLindXpressSymbol, "|", 3) & Format(Parse(strLindXpressSymbol, "|", 2), "00")
        dMult = Val(Parse(m.astrSymbols(lPos), vbTab, 4))
        If Left(strGenesisSymbol, 2) = "O:" Then
            strGenesisSymbol = Mid(strGenesisSymbol, 3)
            dStrikeMult = Val(Parse(m.astrSymbols(lPos), vbTab, 9))
            If (dStrikeMult <> 0) And (Len(strStrike) > 0) Then
                dStrike = Val(strStrike) / dStrikeMult
            Else
                dStrike = 0#
            End If
            If (Len(strPutOrCall) > 0) And (dStrike <> 0#) Then
                strGenesisSymbol = strGenesisSymbol & " " & strPutOrCall & Format(dStrike, "0.#######")
                If Right(strGenesisSymbol, 1) = "." Then strGenesisSymbol = Left(strGenesisSymbol, Len(strGenesisSymbol) - 1)
            End If
        End If
        bReturn = True
    Else
        strGenesisSymbol = ""
        dMult = 0#
        bReturn = False
    End If

    GenesisSymbol = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.GenesisSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbolForPosition
'' Description: Convert a LindXpress position symbol to a Genesis symbol
'' Inputs:      LindXpress Symbol, Genesis Symbol, Multiplier
'' Returns:     True if converted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbolForPosition(ByVal strLindXpressSymbol As String, strGenesisSymbol As String, dMult As Double) As Boolean
On Error GoTo ErrSection:

    Dim strBase As String               ' Base symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrLine As cGdArray            ' Line in the symbols array
    Dim bReturn As Boolean              ' Return value from the function
    Dim strStrike As String             ' Strike price
    Dim strPutOrCall As String          ' Put or call
    Dim dStrikeMult As Double           ' Strike price multiplier
    Dim dStrike As Double               ' Strike price
    Dim bFound As Boolean               ' Has the symbol been found?
    
    Set astrLine = New cGdArray
    astrLine.Create eGDARRAY_Strings
    
    strBase = Parse(strLindXpressSymbol, "|", 1)
    strStrike = Parse(strLindXpressSymbol, "|", 4)
    If strStrike = "0" Then strStrike = ""
    strPutOrCall = Parse(strLindXpressSymbol, "|", 5)
    strGenesisSymbol = ""
    dMult = 0#
    bReturn = False
    
    If (Len(strStrike) > 0) And (Len(strPutOrCall) > 0) Then
        strBase = "O:" & strBase
    End If
    
    For lIndex = 0 To m.astrSymbols.Size - 1
        astrLine.SplitFields m.astrSymbols(lIndex), vbTab
        
        ' 03/15/2011 DAJ: As per Joe Crowe, if the user is logged in as a "User", the position will
        ' be sent with the "internal symbol", if the user is logged in as an "Account", the position
        ' will be sent with the "customer symbol"
        bFound = False
        If Len(m.strUserName) > 0 Then
            bFound = (astrLine(9) = strBase)
        ElseIf Len(m.strLoginAccount) > 0 Then
            bFound = (astrLine(7) = strBase)
        End If
        
        If bFound Then
            strGenesisSymbol = astrLine(2) & "-" & Parse(strLindXpressSymbol, "|", 3) & Format(Parse(strLindXpressSymbol, "|", 2), "00")
            dMult = Val(astrLine(3))
            
            If Left(strGenesisSymbol, 2) = "O:" Then
                strGenesisSymbol = Mid(strGenesisSymbol, 3)
                dStrikeMult = Val(astrLine(8))
                If (dStrikeMult <> 0) And (Len(strStrike) > 0) Then
                    dStrike = Val(strStrike) / dStrikeMult
                Else
                    dStrike = 0#
                End If
                If (Len(strPutOrCall) > 0) And (dStrike <> 0#) Then
                    strGenesisSymbol = strGenesisSymbol & " " & strPutOrCall & Format(dStrike, "0.#######")
                    If Right(strGenesisSymbol, 1) = "." Then strGenesisSymbol = Left(strGenesisSymbol, Len(strGenesisSymbol) - 1)
                End If
            End If
            
            bReturn = True
            Exit For
        End If
    Next lIndex

    GenesisSymbolForPosition = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.GenesisSymbolForPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderStatus
'' Description: Convert a LindXpress order status to a Genesis order status
'' Inputs:      LindXpress Order Status
'' Returns:     Genesis Order Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderStatus(ByVal strLindXpressOrderStatus As String) As eTT_OrderStatus
On Error GoTo ErrSection:

    Select Case strLindXpressOrderStatus
        Case "0"            ' OS_NONE
            GenesisOrderStatus = eTT_OrderStatus_Open
        Case "1"            ' OS_CANCELED
            GenesisOrderStatus = eTT_OrderStatus_Cancelled
        Case "2"            ' OS_REPLACED
            GenesisOrderStatus = eTT_OrderStatus_Amended
        Case "3"            ' OS_WORKING
            GenesisOrderStatus = eTT_OrderStatus_Working
        Case "4"            ' OS_EXPIRED
            GenesisOrderStatus = eTT_OrderStatus_Expired
        Case "5"            ' OS_FILLED
            GenesisOrderStatus = eTT_OrderStatus_Filled
        Case "6"            ' OS_SPECIAL
            GenesisOrderStatus = eTT_OrderStatus_Cancelled
        Case "7"            ' OS_CONFIRMED_OUT
            GenesisOrderStatus = eTT_OrderStatus_Cancelled
        Case "8"            ' OS_REJECTED
            GenesisOrderStatus = eTT_OrderStatus_Rejected
        Case "9"            ' OS_ERROR
            GenesisOrderStatus = eTT_OrderStatus_Error
        Case "10"           ' OS_DEAD
            GenesisOrderStatus = eTT_OrderStatus_Expired
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.GenesisOrderStatus"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderType
'' Description: Convert a LindXpress order type to a Genesis order type
'' Inputs:      LindXpress Order Type
'' Returns:     Genesis Order Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderType(ByVal strLindXpressOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Select Case strLindXpressOrderType
        Case "0"            ' OT_NONE
        Case "1"            ' OT_MARKET
            GenesisOrderType = eTT_OrderType_Market
        Case "2"            ' OT_LIMIT
            GenesisOrderType = eTT_OrderType_Limit
        Case "3"            ' OT_MIT
        Case "4"            ' OT_MOC
            GenesisOrderType = eTT_OrderType_MarketOnClose
        Case "5"            ' OT_FOK
        Case "6"            ' OT_OB
        Case "7"            ' OT_OBOO
        Case "8"            ' OT_OBCO
        Case "9"            ' OT_STOP
            GenesisOrderType = eTT_OrderType_Stop
        Case "10"           ' OT_SCO
            GenesisOrderType = eTT_OrderType_StopCloseOnly
        Case "11"           ' OT_SOO
        Case "12"           ' OT_STWL
            GenesisOrderType = eTT_OrderType_StopWithLimit
        Case "13"           ' OT_STL
        Case "14"           ' OT_IOC
        Case "15"           ' OT_AON
        Case "16"           ' OT_OO
        Case "17"           ' OT_FAK
        Case "18"           ' OT_GTD
        Case "19"           ' OT_GTC
        Case "20"           ' OT_OCO
        Case "21"           ' OT_ONLYMARKET
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.GenesisOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrder
'' Description: Load the order from the database for Genesis ID or Broker ID
'' Inputs:      Account Number, Genesis Order ID, Broker ID, Original ID
'' Returns:     Order (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadOrder(ByVal strAccountNumber As String, ByVal strGenesisOrderID As String, ByVal strBrokerID As String, ByVal strOriginalID As String) As cPtOrder
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
    If (bFound = False) And (Len(strOriginalID) > 0) Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                    "WHERE [BrokerOrderID]='" & strOriginalID & "' AND [AccountID]=" & lAccountID & ";", dbOpenDynaset)
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
    RaiseError "cXpress.LoadOrder"
    
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

    Dim dDate As Double                 ' Date portion
    Dim strTime As String               ' Time portion
    Dim dTime As Double                 ' Time portion of the date
    Dim dReturn As Double               ' Return value from the function
    
    If Len(strDate) > 0 Then
        dDate = JulFromLong(Val(Left(strDate, 8)))
        strTime = Right(Parse(strDate, " ", 1), 6)
    
        dTime = Val(Left(strTime, 2)) * 3600 + Val(Mid(strTime, 3, 2)) * 60 + Val(Right(strTime, 2))
        dTime = dTime / 86400#
        
        dReturn = dDate + dTime
    End If
    
    DateToDouble = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.DateToDouble"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MessageTypeToString
'' Description: Convert a message type enumeration to a string
'' Inputs:      Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MessageTypeToString(ByVal nType As eGDLindXpressMessageTypes) As String
On Error GoTo ErrSection:

    Select Case nType
        Case eGDLindXpressMessageType_Connect
            MessageTypeToString = "eGDLindXpressMessageType_Connect"
        Case eGDLindXpressMessageType_Disconnect
            MessageTypeToString = "eGDLindXpressMessageType_Disconnect"
        Case eGDLindXpressMessageType_AddOrder
            MessageTypeToString = "eGDLindXpressMessageType_AddOrder"
        Case eGDLindXpressMessageType_AmendOrder
            MessageTypeToString = "eGDLindXpressMessageType_AmendOrder"
        Case eGDLindXpressMessageType_CancelOrder
            MessageTypeToString = "eGDLindXpressMessageType_CancelOrder"
        Case eGDLindXpressMessageType_UnloadApp
            MessageTypeToString = "eGDLindXpressMessageType_UnloadApp"
        Case eGDLindXpressMessageType_GetAccounts
            MessageTypeToString = "eGDLindXpressMessageType_GetAccounts"
        Case eGDLindXpressMessageType_GetOrders
            MessageTypeToString = "eGDLindXpressMessageType_GetOrders"
        Case eGDLindXpressMessageType_GetFills
            MessageTypeToString = "eGDLindXpressMessageType_GetFills"
        Case eGDLindXpressMessageType_GetPositions
            MessageTypeToString = "eGDLindXpressMessageType_GetPositions"
        Case eGDLindXpressMessageType_GetAllOrders
            MessageTypeToString = "eGDLindXpressMessageType_GetAllOrders"
        Case eGDLindXpressMessageType_GetAllFills
            MessageTypeToString = "eGDLindXpressMessageType_GetAllFills"
        Case eGDLindXpressMessageType_GetAllPositions
            MessageTypeToString = "eGDLindXpressMessageType_GetAllPositions"
        Case eGDLindXpressMessageType_GetAccountInfo
            MessageTypeToString = "eGDLindXpressMessageType_GetAccountInfo"
        
        Case eGDLindXpressMessageType_ConnectionInfo
            MessageTypeToString = "eGDLindXpressMessageType_ConnectionInfo"
        Case eGDLindXpressMessageType_AppLoaded
            MessageTypeToString = "eGDLindXpressMessageType_AppLoaded"
        Case eGDLindXpressMessageType_AppUnloaded
            MessageTypeToString = "eGDLindXpressMessageType_AppUnloaded"
        Case eGDLindXpressMessageType_Heartbeat
            MessageTypeToString = "eGDLindXpressMessageType_Heartbeat"
        Case eGDLindXpressMessageType_Account
            MessageTypeToString = "eGDLindXpressMessageType_Account"
        Case eGDLindXpressMessageType_Order
            MessageTypeToString = "eGDLindXpressMessageType_Order"
        Case eGDLindXpressMessageType_Fill
            MessageTypeToString = "eGDLindXpressMessageType_Fill"
        Case eGDLindXpressMessageType_Position
            MessageTypeToString = "eGDLindXpressMessageType_Position"
        Case eGDLindXpressMessageType_Alert
            MessageTypeToString = "eGDLindXpressMessageType_Alert"
        Case eGDLindXpressMessageType_AccountInfo
            MessageTypeToString = "eGDLindXpressMessageType_AccountInfo"
        
        Case Else
            MessageTypeToString = Str(nType)
            
    End Select
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.MessageTypeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StringToMessageType
'' Description: Convert a string to a message type enumeration
'' Inputs:      String
'' Returns:     Message Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StringToMessageType(ByVal strMessageType As String) As eGDLindXpressMessageTypes
On Error GoTo ErrSection:

    Dim nReturn As eGDLindXpressMessageTypes    ' Return value for the function

    Select Case strMessageType
        Case "eGDLindXpressMessageType_Connect"
            nReturn = eGDLindXpressMessageType_Connect
        Case "eGDLindXpressMessageType_Disconnect"
            nReturn = eGDLindXpressMessageType_Disconnect
        Case "eGDLindXpressMessageType_AddOrder"
            nReturn = eGDLindXpressMessageType_AddOrder
        Case "eGDLindXpressMessageType_AmendOrder"
            nReturn = eGDLindXpressMessageType_AmendOrder
        Case "eGDLindXpressMessageType_CancelOrder"
            nReturn = eGDLindXpressMessageType_CancelOrder
        Case "eGDLindXpressMessageType_UnloadApp"
            nReturn = eGDLindXpressMessageType_UnloadApp
        Case "eGDLindXpressMessageType_GetAccounts"
            nReturn = eGDLindXpressMessageType_GetAccounts
        Case "eGDLindXpressMessageType_GetOrders"
            nReturn = eGDLindXpressMessageType_GetOrders
        Case "eGDLindXpressMessageType_GetFills"
            nReturn = eGDLindXpressMessageType_GetFills
        Case "eGDLindXpressMessageType_GetPositions"
            nReturn = eGDLindXpressMessageType_GetPositions
        Case "eGDLindXpressMessageType_GetAllOrders"
            nReturn = eGDLindXpressMessageType_GetAllOrders
        Case "eGDLindXpressMessageType_GetAllFills"
            nReturn = eGDLindXpressMessageType_GetAllFills
        Case "eGDLindXpressMessageType_GetAllPositions"
            nReturn = eGDLindXpressMessageType_GetAllPositions
        Case "eGDLindXpressMessageType_GetAccountInfo"
            nReturn = eGDLindXpressMessageType_GetAccountInfo
    
        Case "eGDLindXpressMessageType_ConnectionInfo"
            nReturn = eGDLindXpressMessageType_ConnectionInfo
        Case "eGDLindXpressMessageType_AppLoaded"
            nReturn = eGDLindXpressMessageType_AppLoaded
        Case "eGDLindXpressMessageType_AppUnloaded"
            nReturn = eGDLindXpressMessageType_AppUnloaded
        Case "eGDLindXpressMessageType_Heartbeat"
            nReturn = eGDLindXpressMessageType_Heartbeat
        Case "eGDLindXpressMessageType_Account"
            nReturn = eGDLindXpressMessageType_Account
        Case "eGDLindXpressMessageType_Order"
            nReturn = eGDLindXpressMessageType_Order
        Case "eGDLindXpressMessageType_Fill"
            nReturn = eGDLindXpressMessageType_Fill
        Case "eGDLindXpressMessageType_Position"
            nReturn = eGDLindXpressMessageType_Position
        Case "eGDLindXpressMessageType_Alert"
            nReturn = eGDLindXpressMessageType_Alert
        Case "eGDLindXpressMessageType_AccountInfo"
            nReturn = eGDLindXpressMessageType_AccountInfo
    End Select
    
    StringToMessageType = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.StringToMessageType"
        
End Function

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBrokerInfoOrder
'' Description: Add the order string to the broker info class
'' Inputs:      Order String, Refresh?
'' Returns:     None
''
'' RecordFrom:  Genesis ID, Broker ID, Original ID, Status, Account, Order Time,
''              Business Date, Error Code, Error Text, Symbol, B/S, Order Type,
''              Price, With Limit Price, Quantity, Qty Filled, Qty Working,
''              Expiration, Floor Recv Time, Broker Recv Time, Broker Exec Time,
''              Fill ID, Trade Date, Trade Price, Fill Qty
'' Order:       Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date
'' Fill:        Broker ID, Genesis ID, Fill ID, Account, Symbol, Fill Date,
''              B/S, Fill Quantity, Fill Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateBrokerInfoOrder(ByVal strOrderString As String, ByVal bRefresh As Boolean, ByVal bHasFillInfo As Boolean)
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Order information split out from string
    Dim astrToAdd As cGdArray           ' Order information to add to broker info
    Dim strGenesisSymbol As String      ' Genesis symbol from the LindXpress symbol
    Dim dMult As Double                 ' Multiplier
    Dim nOrderType As eTT_OrderType     ' Order type for the order
    
    Set astrOrder = New cGdArray
    astrOrder.Create eGDARRAY_Strings
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings
    
    astrOrder.SplitFields strOrderString, vbTab
    
    If Len(astrOrder(5)) >= 8 Then
        astrOrder(1) = Left(astrOrder(5), 8) & "-" & astrOrder(1)
        If (Len(astrOrder(2)) > 0) And (astrOrder(2) <> "0") Then
            astrOrder(2) = Left(astrOrder(5), 8) & "-" & astrOrder(2)
        End If
    End If
    
    If UCase(astrOrder(0)) = "SPECIAL" Then
        astrToAdd(0) = astrOrder(1)
        astrToAdd(1) = astrOrder(0)
        astrToAdd(2) = astrOrder(2)
        astrToAdd(3) = astrOrder(4)
        astrToAdd(4) = Str(GenesisOrderStatus(astrOrder(3)))
        astrToAdd(5) = Str(DateToDouble(astrOrder(5)))
        
        m.BrokerInfo.OrderCancelled astrToAdd.JoinFields(vbTab), bRefresh
    ElseIf UCase(astrOrder(0)) = "REPLACE" Then
        astrToAdd(0) = astrOrder(1)
        astrToAdd(1) = ""
        astrToAdd(2) = astrOrder(2)
        astrToAdd(3) = astrOrder(4)
        astrToAdd(4) = Str(GenesisOrderStatus(astrOrder(3)))
        astrToAdd(5) = Str(DateToDouble(astrOrder(5)))
        
        m.BrokerInfo.OrderReplaced astrToAdd.JoinFields(vbTab), bRefresh
    Else
        If GenesisSymbol(astrOrder(9), strGenesisSymbol, dMult) Then
            nOrderType = GenesisOrderType(astrOrder(11))
            
            astrToAdd(0) = astrOrder(1)
            astrToAdd(1) = astrOrder(0)
            If (Len(astrOrder(2)) = 0) Or (astrOrder(2) = "0") Then astrToAdd(2) = "" Else astrToAdd(2) = astrOrder(2)
            astrToAdd(3) = astrOrder(4)
            astrToAdd(4) = strGenesisSymbol
            astrToAdd(5) = Str(nOrderType)
            Select Case astrOrder(10)
                Case "1"
                    astrToAdd(6) = "1"
                Case "2"
                    astrToAdd(6) = "0"
            End Select
            astrToAdd(7) = astrOrder(14)
            Select Case nOrderType
                Case eTT_OrderType_Market, eTT_OrderType_MarketOnClose
                    astrToAdd(8) = ""
                    astrToAdd(9) = ""
                Case eTT_OrderType_Stop, eTT_OrderType_StopCloseOnly
                    astrToAdd(8) = ""
                    astrToAdd(9) = GenesisPrice(astrOrder(12), dMult)
                Case eTT_OrderType_Limit, eTT_OrderType_LimitCloseOnly
                    astrToAdd(8) = GenesisPrice(astrOrder(12), dMult)
                    astrToAdd(9) = ""
                Case eTT_OrderType_StopWithLimit, eTT_OrderType_StopWithLimitCloseOnly
                    astrToAdd(8) = GenesisPrice(astrOrder(12), dMult)
                    astrToAdd(9) = GenesisPrice(astrOrder(13), dMult)
            End Select
            astrToAdd(10) = astrOrder(17)
            astrToAdd(11) = Str(GenesisOrderStatus(astrOrder(3)))
            astrToAdd(12) = astrOrder(16)
            astrToAdd(13) = Str(DateToDouble(astrOrder(5)))
            
DumpDebug "ADDING ORDER: " & astrToAdd.JoinFields(vbTab)
            m.BrokerInfo.AddOrder astrToAdd.JoinFields(vbTab), bRefresh
            
            If bHasFillInfo Then
                If CLng(Val(astrOrder(24))) > 0 Then
                    astrToAdd.Clear
                    astrToAdd(0) = astrOrder(1)
                    astrToAdd(1) = astrOrder(0)
                    astrToAdd(2) = astrOrder(1) & "-" & astrOrder(21)
                    astrToAdd(3) = astrOrder(4)
                    astrToAdd(4) = strGenesisSymbol
                    astrToAdd(5) = Str(DateToDouble(astrOrder(20)))
                    Select Case astrOrder(10)
                        Case "1"
                            astrToAdd(6) = "1"
                        Case "2"
                            astrToAdd(6) = "0"
                    End Select
                    astrToAdd(7) = astrOrder(24)
                    astrToAdd(8) = Str(GenesisPrice(Val(astrOrder(23)), dMult))
                    
                    m.BrokerInfo.AddFill astrToAdd.JoinFields(vbTab), bRefresh
                Else
                    m.BrokerInfo.RemoveFill astrOrder(1) & "-" & astrOrder(21)
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.UpdateBrokerInfoOrder"
    
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
    
DumpDebug "ADDING ORDER: " & astrToAdd.JoinFields(vbTab)
    m.BrokerInfo.AddOrder astrToAdd.JoinFields(vbTab), False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.GenesisOrderToBrokerInfo"
    
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
'' RecordTo:    Account, Genesis ID, Broker ID, Confirmation
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelBrokerInfoOrder(ByVal strOrder As String)
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Order broken out into an array

    If Len(strOrder) > 0 Then
        Set astrOrder = New cGdArray
        astrOrder.Create eGDARRAY_Strings
    
        astrOrder.SplitFields strOrder, vbTab
        
        If InStr(astrOrder(0), "-") = 0 Then
            SendLindXpressMessage eGDLindXpressMessageType_CancelOrder, astrOrder(3) & vbTab & astrOrder(1) & vbTab & astrOrder(0) & vbTab & "0"
        Else
            SendLindXpressMessage eGDLindXpressMessageType_CancelOrder, astrOrder(3) & vbTab & astrOrder(1) & vbTab & Parse(astrOrder(0), "-", 2) & vbTab & "0"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.CancelBrokerInfoOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFillToBrokerInfo
'' Description: Add the fill string to the broker info class
'' Inputs:      Fill String, Refresh?
'' Returns:     None
''
'' RecordFrom:  Genesis ID, Broker ID, Original ID, Status, Account, Order Time,
''              Business Date, Error Code, Error Text, Symbol, B/S, Order Type,
''              Price, With Limit Price, Quantity, Qty Filled, Qty Working,
''              Expiration, Floor Recv Time, Broker Recv Time, Broker Exec Time,
''              Fill ID, Trade Date, Trade Price, Fill Qty
'' RecordTo:    Broker ID, Genesis ID, Fill ID, Account, Symbol, Fill Date,
''              B/S, Fill Quantity, Fill Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddFillToBrokerInfo(ByVal strFillString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.AddFillToBrokerInfo"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddPositionToBrokerInfo
'' Description: Add the position string to the broker info class
'' Inputs:      Position String, Refresh?
'' Returns:     None
''
'' RecordFrom:  Account, Symbol, B/S, Quantity, Currency, Desc, Firm,
''              Interface Type, Net Profit, Price, Price Kind, P/C, Record Type,
''              Reg Code, Seq Number, Settlement Price, Strike Price, Tick,
''              Trade Date, Validate, Version
'' RecordTo:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddPositionToBrokerInfo(ByVal strPositionString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim astrPosition As cGdArray        ' Array of position information split out from string
    Dim astrToAdd As cGdArray           ' Array of information to pass to broker info class
    Dim strGenesisSymbol As String      ' Genesis symbol from the LindXpress symbol
    Dim dMult As Double                 ' Multiplier
    
    Set astrPosition = New cGdArray
    astrPosition.Create eGDARRAY_Strings
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings
    
    astrPosition.SplitFields strPositionString, vbTab
    If GenesisSymbolForPosition(astrPosition(1), strGenesisSymbol, dMult) Then
        astrToAdd(0) = astrPosition(0)
        astrToAdd(1) = strGenesisSymbol
        Select Case astrPosition(2)
            Case "1"
                astrToAdd(2) = astrPosition(3)
            Case "2"
                astrToAdd(2) = Str(CLng(Val(astrPosition(3))) * -1)
        End Select
        astrToAdd(3) = Str(GenesisPrice(astrPosition(9), dMult))
        astrToAdd(4) = astrToAdd(2)
        
        m.BrokerInfo.AddPosition astrToAdd.JoinFields(vbTab), bRefresh
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.AddPositionToBrokerInfo"
    
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
            astrInfo.SplitFields m.astrBrokerPos(lIndex), vbTab
            
            astrToAdd.Clear
            astrToAdd(0) = astrInfo(0)
            astrToAdd(1) = astrInfo(1)
            astrToAdd(2) = astrInfo(3)
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
                            astrBrokerPos.SplitFields m.astrBrokerPos(lPos), vbTab
                            strOvernight = astrBrokerPos(3)
                            lPosition = CLng(Val(strOvernight))
                            lCurrentPos = lPosition + .NumBuysSnapshot - .NumSellsSnapshot
                            astrBrokerPos(2) = Str(lCurrentPos)
                            m.astrBrokerPos(lPos) = astrBrokerPos.JoinFields(vbTab)
                        Else
                            lPosition = 0&
                            strOvernight = "0"
                            lCurrentPos = .NumBuysSnapshot - .NumSellsSnapshot
                            m.astrBrokerPos.Add strAccount & vbTab & .Symbol & vbTab & Str(lCurrentPos) & vbTab & Str(lPosition) & vbTab & "" & vbTab & "", lPos
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
    RaiseError "cXpress.BuildPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromDateOrders
'' Description: Get the from date for an orders request
'' Inputs:      None
'' Returns:     From Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FromDateOrders() As String
On Error GoTo ErrSection:

    Dim dFromDate As Double             ' Date of the last order refresh
    Dim strReturn As String             ' Return value for the function

    dFromDate = GetIniFileProperty("Order", 0#, "Refresh", m.strIniFile)
    If dFromDate = 0# Then
        dFromDate = CurrentTime("CHI") - 6
    End If
    
    ' Since the historical fills and orders work off of a calendar date instead
    ' of a session date, if the from date is not today, then we need to back up
    ' to the closest Sunday to make sure everything is synced. (DAJ 07/14/2008)...
    If Int(dFromDate) < Int(CurrentTime("CHI")) Then
        Do While Weekday(dFromDate) <> vbSunday
            dFromDate = dFromDate - 1
        Loop
    End If
    strReturn = Format(dFromDate, "YYYYMMDD")
    
    FromDateOrders = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.FromDateOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromDateFills
'' Description: Get the from date for a fill request
'' Inputs:      None
'' Returns:     From Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FromDateFills() As String
On Error GoTo ErrSection:

    Dim dFromDate As Double             ' Date of the last fill refresh
    Dim strReturn As String             ' Return value for the function

    dFromDate = GetIniFileProperty("Fill", 0#, "Refresh", m.strIniFile)
    If dFromDate = 0# Then
        dFromDate = CurrentTime("CHI") - 6
    End If
    
    ' Since the historical fills and orders work off of a calendar date instead
    ' of a session date, if the from date is not today, then we need to back up
    ' to the closest Sunday to make sure everything is synced. (DAJ 07/14/2008)...
    If Int(dFromDate) < Int(CurrentTime("CHI")) Then
        Do While Weekday(dFromDate) <> vbSunday
            dFromDate = dFromDate - 1
        Loop
    End If
    
    m.lLastFromDateFills = Int(dFromDate)
    strReturn = Format(dFromDate, "YYYYMMDD")
        
    FromDateFills = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.FromDateFills"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateLastGenesisID
'' Description: Update the INI file with the last known Genesis ID if necessary
'' Inputs:      Genesis Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateLastGenesisID(ByVal strGenesisID As String)
On Error GoTo ErrSection:

    Dim strCustomerID As String         ' Customer ID out of the ID
    Dim strDate As String               ' Date out of the ID
    Dim strCounter As String            ' Counter out of the ID
    Dim lDate As Long                   ' Date out of the INI file
    Dim lCounter As Long                ' Counter out of the INI file
    
    If Len(strGenesisID) = 17 Then
        strCustomerID = Left(strGenesisID, 7)
        strDate = Mid(strGenesisID, 8, 5)
        strCounter = Right(strGenesisID, 5)
        
        lDate = GetIniFileProperty("Date", 0&, "ID", m.strIniFile)
        lCounter = GetIniFileProperty("Counter", 0&, "ID", m.strIniFile)
        
        If CLng(Val(strDate)) >= lDate Then
            If CLng(Val(strCounter)) > lCounter Then
                SetIniFileProperty "Date", CLng(Val(strDate)), "ID", m.strIniFile
                SetIniFileProperty "Counter", CLng(Val(strCounter)), "ID", m.strIniFile
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.UpdateLastGenesisID"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderForAlert
'' Description: Get the order for the given information from an alert
'' Inputs:      Genesis ID, Broker ID
'' Returns:     Order (or Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderForAlert(ByVal strGenesisID As String, ByVal strBrokerID As String) As cPtOrder
On Error GoTo ErrSection:

    Dim Orders As cPtOrders
    Dim lIndex As Long
    Dim Order As cPtOrder
    
    Set Order = Nothing
    Set Orders = m.BrokerInfo.Orders
    
    If Not Orders Is Nothing Then
        For lIndex = Orders.Count To 1 Step -1
            If InStr(strBrokerID, "-") = 0 Then
                If Parse(Orders(lIndex).BrokerID, "-", 2) = strBrokerID Then
                    If Orders(lIndex).GenesisOrderID = strGenesisID Then
                        Set Order = Orders(lIndex)
                        Exit For
                    End If
                End If
            Else
                If Orders(lIndex).BrokerID = strBrokerID Then
                    If Orders(lIndex).GenesisOrderID = strGenesisID Then
                        Set Order = Orders(lIndex)
                        Exit For
                    End If
                End If
            End If
        Next lIndex
    End If
    
    Set OrderForAlert = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cXpress.OrderForAlert"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when a new object is created
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Dim strDrive As String              ' Drive for the Xpress OCX and DLL's
    
    ConnectionStatus = eGDConnectionStatus_Disconnected
    
    Set m.astrToFix = New cGdArray
    m.astrToFix.Create eGDARRAY_Strings
    
    m.bWasConnected = False
    m.bVerifyPositions = True
    m.bSyncInProgress = False
    m.bFixCalledOnce = False
    
    ' Attempt to register the appropriate DLL every time Trade Navigator is started...
    strDrive = Parse(App.Path, ":", 1) & ":"
    ChangePath AddSlash(CommonFilesPath) & "LindWaldockAPI"
    gdRegisterFiles AddSlash(CommonFilesPath) & "LindWaldockAPI\LindFinancialAPICOM.DLL", ByVal 0
    ChangePath App.Path
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Destroy class members when the object is destroyed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.astrAccounts = Nothing
    Set m.astrBrokerPos = Nothing
    Set m.astrToFix = Nothing

    Set m.BrokerInfo = Nothing
    Set m.astrSymbols = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cXpress.Class_Terminate"
    
End Sub

