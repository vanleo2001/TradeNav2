VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPFG.cls
'' Description: Wrapper class for interacting with the PFG server
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 01/08/2009   DAJ         Added support for prices reported in 64ths
'' 01/15/2009   DAJ         Changed amend order to send order qty instead of
''                          remaining qty as per Dave Rettger e-mail today
'' 01/22/2009   DAJ         Have the 64ths conversion take the format into account
'' 02/05/2009   DAJ         Added support for 1/2 64ths
'' 02/23/2009   DAJ         Set ConsistentBroker to False even if first time
'' 02/23/2009   DAJ         Prepend WeekNum to PFG orders instead of calendar date
'' 02/24/2009   DAJ         PFG Stock symbols are prepended with an "EQ-"
'' 03/17/2009   DAJ         Fixed GenesisPrice for 64ths when three digits come back
'' 03/18/2009   DAJ         Truncate the strike price when converting to PFG symbol
'' 04/08/2009   DAJ         If order not found, try previous week numbers (up to 5)
'' 04/27/2009   DAJ         Don't handle order/fill if the date comes in blank
'' 05/19/2009   DAJ         Use the new Option Navigator status variable
'' 05/27/2009   DAJ         If Add or Amend failed, send error to Option Nav
'' 06/01/2009   DAJ         Added FO and SO to security type mask for trading
'' 06/09/2009   DAJ         Set the BrokerInfo connection status
'' 08/07/2009   DAJ         Handle old Alaron accounts (no time on fills, status X)
'' 08/14/2009   DAJ         Handle stocks and stock options
'' 08/21/2009   DAJ         Fix for converting stock option symbol
'' 08/25/2009   DAJ         Added support for PFG account information
'' 09/01/2009   DAJ         Use new Parked order status
'' 09/16/2009   DAJ         Save the access key information to registry
'' 10/07/2009   DAJ         Added support for Linked Orders at broker
'' 10/14/2009   DAJ         Fixed a problem with creating a carried position fill
'' 10/27/2009   DAJ         Fix for handling expired orders, SendNow flag
'' 11/10/2009   DAJ         Be sure to add space back in when converting SO
'' 12/01/2009   DAJ         Enhancements for Auto Exits held at broker
'' 12/08/2009   DAJ         Attempt to kill GenPFG upon restart if still running
'' 01/27/2010   DAJ         Utilize new cOptionSymbol object for parsing option symbols
'' 02/08/2010   DAJ         Fix for converting to a Genesis symbol
'' 03/11/2010   DAJ         Use global collections
'' 05/25/2010   DAJ         Set HoldOcoAtBroker in the HandleAccount routine on a new account
'' 08/13/2010   DAJ         Added in account flags, handle account with delayed streaming
'' 09/24/2010   DAJ         Fixed the DateToDouble routine to ignore the milliseconds
'' 10/28/2010   DAJ         Added PFG spread calls
'' 11/01/2010   DAJ         Dump TickCount with log information
'' 11/18/2010   DAJ         Added Starting Genesis ID for Option Navigator
'' 02/11/2011   DAJ         Don't adjust date/time on a refresh fill
'' 03/07/2011   DAJ         Changed Option Nav calls, Broker Disconnect Alerts
'' 03/09/2011   DAJ         Reset ManualDisconnect on connect, handle multiple connected accounts
'' 04/01/2011   DAJ         If we get a single digit for 32nds, append a zero
'' 04/04/2011   DAJ         Added the "Queued" order status
'' 05/04/2011   DAJ         If order comes in "Queued" it has the calendar date instead of session
'' 05/09/2011   DAJ         Prepend GenesisID with "TN" so PFG can distinguish order source
'' 05/27/2011   DAJ         Send "Application ID" with Add, Amend, and Cancel order calls
'' 06/21/2011   DAJ         Last connection status change
'' 07/15/2011   DAJ         Handle order link failure from PFG servers
'' 12/14/2011   DAJ         Added Capital Trading Group and Fintec for PFG
'' 01/18/2012   DAJ         Enhanced logging for automated trading, don't overwrite order quantity with zero
'' 01/30/2012   DAJ         Fix for handling an order link failure
'' 01/31/2012   DAJ         Handle unsolicted cancels in TradeSense order groups
'' 02/14/2012   DAJ         New status alerts for position mismatch / auto trade disabled
'' 03/21/2012   DAJ         Fix for amending a partially filled order
'' 06/18/2012   DAJ         Fix for getting a one digit price for ZN (32.25)
'' 06/25/2012   DAJ         Process an unsolicited cancel even if status date is older
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 12/11/2012   DAJ         Use the flatten queue for position reversals
'' 12/11/2012   DAJ         Handle broker status in automated trading item
'' 04/17/2013   DAJ         Added direction filter to CancelWorkingOrders call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strIniFile As String                ' Ini file for the broker information
    nBroker As eTT_AccountType          ' Broker type
    strBrokerName As String             ' Broker name
    strProcessName As String            ' Name of the process to check
    strProcessPath As String            ' Full path of the process to start
    strControlID As String              ' Control ID
    strLogPath As String                ' Path for the log files
    strConnectIni As String             ' Name and path of the ini file with connection information
    
    strUserName As String               ' User name to logon to the server with
    strPassword As String               ' Password to logon to the server with
    strHostIP As String                 ' Host IP for the PFG server
    strHostPort As String               ' Host port for the PFG server
    strServer As String                 ' Server that is being connected to
    
    nStatus As eGDConnectionStatus      ' Connection status to the PFG servers
    bWasConnected As Boolean            ' This broker was has been connected
    dHeartBeat As Double                ' Tick count as of last message receieved
    bReconnecting As Boolean            ' Are we currently reconnecting?
    lReconnectAttempts As Long          ' Reconnect attempts that have been made
    bAppLoaded As Boolean               ' Is the PFG application loaded?
    bManualDisconnect As Boolean        ' Is the user manually disconnecting?
    bAskedPassword As Boolean           ' Have we asked for the password yet?
    dLastFillTime As Double             ' Tick count of when we received the last fill
    strLastMessage As String            ' Last message from the connection status message
    bVerifyPositions As Boolean         ' Do we want to verify the positions with the user?
    bVerboseRefresh As Boolean          ' Do we want to tell the user about the refresh?
    bSyncInProgress As Boolean          ' Are we in the process of synchronizing positions?
    bFixCalledOnce As Boolean           ' Has the fix positions call been done at least once?
    lLastConnectDate As Long            ' Last connection date
    
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

Public Property Get IniFile() As String
    IniFile = m.strIniFile
End Property
Public Property Let IniFile(ByVal strIniFile As String)
    m.strIniFile = strIniFile
End Property

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property

Public Property Get BrokerName() As String
    BrokerName = m.strBrokerName
End Property
Public Property Let BrokerName(ByVal strBrokerName As String)
    m.strBrokerName = strBrokerName
End Property

Public Property Get ProcessName() As String
    ProcessName = m.strProcessName
End Property
Public Property Let ProcessName(ByVal strProcessName As String)
    m.strProcessName = strProcessName
End Property

Public Property Get ProcessPath() As String
    ProcessPath = m.strProcessPath
End Property
Public Property Let ProcessPath(ByVal strProcessPath As String)
    m.strProcessPath = strProcessPath
End Property

Public Property Get ControlID() As String
    ControlID = m.strControlID
End Property
Public Property Let ControlID(ByVal strControlID As String)
    m.strControlID = strControlID
End Property

Public Property Get LogPath() As String
    LogPath = m.strLogPath
End Property

Public Property Get ConnectIni() As String
    ConnectIni = m.strConnectIni
End Property
Public Property Let ConnectIni(ByVal strConnectIni As String)
    m.strConnectIni = strConnectIni
End Property

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

Public Property Get Server() As String
    Server = m.strServer
End Property
Public Property Let Server(ByVal strServer As String)
    m.strServer = strServer
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
    RaiseError "cPFG.CurrentPosition.Get"
    
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
        m.astrBrokerPos(lPos) = strAccountNumber & vbTab & strSymbol & vbTab & Str(lCurrentPosition) & vbTab & Parse(m.astrBrokerPos(lPos), vbTab, 4)
    Else
        m.astrBrokerPos.Add strAccountNumber & vbTab & strSymbol & vbTab & Str(lCurrentPosition) & vbTab & "0", lPos
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPFG.CurrentPosition.Let"
    
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
    RaiseError "cPFG.CarriedPosition.Get"
    
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
        m.astrBrokerPos(lPos) = strAccountNumber & vbTab & strSymbol & vbTab & Parse(m.astrBrokerPos(lPos), vbTab, 3) & vbTab & Str(lCarriedPosition)
    Else
        m.astrBrokerPos.Add strAccountNumber & vbTab & strSymbol & vbTab & "0" & vbTab & Str(lCarriedPosition), lPos
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPFG.CarriedPosition.Let"
    
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
        Case eTT_AccountType_PFG
            m.strIniFile = AddSlash(App.Path) & "PFG.INI"
            m.nBroker = eTT_AccountType_PFG
            m.strBrokerName = "PFG"
            m.strProcessName = "GenPFG"
            m.strProcessPath = AddSlash(App.Path) & "..\Brokers\PFG\GenPFG.EXE"
            m.strControlID = "PFG"
            m.strLogPath = AddSlash(App.Path) & "PFG"
            m.strConnectIni = AddSlash(App.Path) & "Provided\PfgIps.INI"
            strSymbolFile = AddSlash(App.Path) & "Provided\PfgToGen2.TXT"
        
        Case eTT_AccountType_CtgPfg
            m.strIniFile = AddSlash(App.Path) & "CtgPfg.INI"
            m.nBroker = eTT_AccountType_CtgPfg
            m.strBrokerName = "CTG3"
            m.strProcessName = "GenCtgPfg"
            m.strProcessPath = AddSlash(App.Path) & "..\Brokers\CtgPfg\GenCtgPfg.EXE"
            m.strControlID = "CtgPfg"
            m.strLogPath = AddSlash(App.Path) & "CtgPfg"
            m.strConnectIni = AddSlash(App.Path) & "Provided\CtgPfgIps.INI"
            strSymbolFile = AddSlash(App.Path) & "Provided\CtgPfgToGen.TXT"
            
        Case eTT_AccountType_FintecPfg
            m.strIniFile = AddSlash(App.Path) & "FintecPfg.INI"
            m.nBroker = eTT_AccountType_FintecPfg
            m.strBrokerName = "Fintec"
            m.strProcessName = "GenFintecPfg"
            m.strProcessPath = AddSlash(App.Path) & "..\Brokers\FintecPfg\GenFintecPfg.EXE"
            m.strControlID = "FintecPfg"
            m.strLogPath = AddSlash(App.Path) & "FintecPfg"
            m.strConnectIni = AddSlash(App.Path) & "Provided\FintecPfgIps.INI"
            strSymbolFile = AddSlash(App.Path) & "Provided\FintecPfgToGen.TXT"
            
    End Select
    
    FileCopy AddSlash(App.Path) & "..\Brokers\GenPFG.EXE", m.strProcessPath, True
                    
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
    RaiseError "cPFG.Init"
    
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

    If (Len(m.strUserName) = 0) Or (Len(m.strHostIP) = 0) Or (Len(m.strHostPort) = 0) Then
        ShowConfig = True
    Else
        ShowConfig = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.ShowConfig"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Connect
'' Description: Attempt a connection to the PFG servers if not already connected
'' Inputs:      New Account, Are we switching?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Connect(Optional ByVal strNewAccount As String = "", Optional ByVal bSwitching As Boolean = False)
On Error GoTo ErrSection:

    Dim bConnect As Boolean             ' Should we still try to connect?
    Dim lTimeOut As Long                ' Timeout variable
    
    If g.bUnloading = False Then
        If (ConnectionStatus = eGDConnectionStatus_Disconnected) Then
            If (m.bAskedPassword = True) And ((strNewAccount = m.strUserName) Or (Len(strNewAccount) = 0)) Then
                bConnect = True
            Else
                bConnect = frmPfgLogin.ShowMe(m.nBroker, strNewAccount)
                m.bAskedPassword = bConnect
            End If
        Else
            If ((strNewAccount <> m.strUserName) And (Len(strNewAccount) > 0)) Or (bSwitching = True) Then
                bConnect = frmPfgLogin.ShowMe(m.nBroker, strNewAccount, False, bSwitching)
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
                    
                    DumpDebug "Trade Navigator version: " & FormatVersion & "." & Str(App.Revision) & " " & DateFormat(FileDate(App.Path & "\" & App.EXEName & ".EXE"), MM_DD_YYYY, HH_MM, AMPM_UPPER)
                    DumpDebug "Loading application " & m.strProcessPath
                    RunProcess m.strProcessPath, , , vbHide
                Else
                    DumpDebug "Connecting to " & m.strBrokerName & " as " & m.strUserName & " (" & m.strHostIP & ":" & m.strHostPort & ")"
                    SendPfgMessage eGDPfgMessageType_Connect, m.strUserName & vbTab & m.strPassword & vbTab & m.strHostIP & vbTab & m.strHostPort & vbTab & "GBT" & vbTab & "TS" & vbTab & FormatVersion & "." & App.Revision
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.Connect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disconnect
'' Description: Disconnect form the PFG servers if connected
'' Inputs:      Force Reconnect?, Reason for Disconnect
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Disconnect(Optional ByVal bForceReconnect As Boolean = False, Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    If (ConnectionStatus = eGDConnectionStatus_Connected) Then
        m.bManualDisconnect = Not bForceReconnect
        DumpDebug "ManualDisconnect = " & Str(m.bManualDisconnect) & " (Disconnect)"
        If m.bManualDisconnect Then
            DumpDebug vbTab & "Heartbeat cleared -- Disconnect"
            m.dHeartBeat = 0#
            g.OrderStrategies.DeactivateExitsForBroker m.nBroker, "Disconnecting from " & m.strBrokerName, True
        End If
        
        If Len(strReason) = 0 Then DumpDebug "Disconnecting from " & m.strBrokerName Else DumpDebug "Disconnecting from " & m.strBrokerName & ": " & strReason
        SendPfgMessage eGDPfgMessageType_Disconnect, ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.Disconnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reconnect
'' Description: Attempt to reconnect to the PFG servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Reconnect()
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Time out variable
    
    DumpDebug "Reconnect(" & Str(m.bReconnecting) & ", " & Str(ConnectionStatus) & ", " & Str(m.bAppLoaded) & ", " & Str(m.lReconnectAttempts) & ")"

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
        
        DumpDebug vbTab & "AppLoaded = " & Str(m.bAppLoaded) & "; TimeOut = " & Str(lTimeOut) & "; Status = " & Str(ConnectionStatus)
                        
        If (ConnectionStatus <> eGDConnectionStatus_Connected) And (ConnectionStatus <> eGDConnectionStatus_Connecting) Then
            If (m.bAppLoaded = False) Or (lTimeOut >= 30&) Then
                ' 12/08/2009 DAJ: If the GenPFG program is still alive for some reason, try to kill it...
                If KillProcess(m.strProcessName) > 0 Then
                    DumpDebug m.strProcessName & " killed"
                    Sleep 3
                End If
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
    RaiseError "cPFG.Reconnect"
    
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

    Dim strAddress As String            ' Address to PFG URL
    
    strAddress = GetIniFileProperty("Address2", "", "Online", m.strConnectIni)
    If Len(strAddress) > 0 Then
        strAddress = Replace(strAddress, "[account_number]", m.strUserName)
        strAddress = Replace(strAddress, "[access_key]", m.strPassword)
        strAddress = Replace(strAddress, "[system]", m.strServer)
    
        RunProcess InternetBrowser, Chr(34) & strAddress & Chr(34)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.ShowAccountOnline"
    
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

    Dim nStatus As eTT_OrderStatus      ' Order status

    If NotSent(Order.Status) Or (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        If AddOrder(Order) = False Then
            nStatus = eTT_OrderStatus_Error
        Else
            nStatus = eTT_OrderStatus_Sent
        End If
    Else
        If AmendOrder(Order) = False Then
            nStatus = eTT_OrderStatus_Error
        Else
            nStatus = eTT_OrderStatus_AmendPending
        End If
    End If
    
    Order.ChangeOrderStatus nStatus, m.BrokerInfo

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.SendOrder"

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
    Dim strPfgSymbol As String          ' Pfg Symbol
    Dim strPfgOrderType As String       ' Pfg Order Type
    Dim dMult As Double                 ' Multiplier for converting prices
    Dim strFormat As String             ' Price format
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    strPfgSymbol = PfgSymbol(Order.Symbol, dMult, strFormat)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strPfgOrderType = PfgOrderType(Order.OrderType)
    
    If Len(strPfgSymbol) = 0 Then
        DumpDebug "AddOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
    ElseIf Len(strPfgOrderType) = 0 Then
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
        If IsForex(Order.Symbol) Then
            astrOrder(3) = strPfgSymbol
        Else
            astrOrder(3) = Left(strPfgSymbol, Len(strPfgSymbol) - 2)
        End If
        astrOrder(4) = strPfgSymbol
        astrOrder(5) = strPfgOrderType
        If Order.Buy Then astrOrder(6) = "1" Else astrOrder(6) = "2"
        astrOrder(7) = Str(Order.Quantity)
        Select Case Order.OrderType
            Case eTT_OrderType_Market
                astrOrder(8) = ""
                astrOrder(9) = ""
            
            Case eTT_OrderType_Stop
                astrOrder(8) = ""
                astrOrder(9) = PfgPrice(Order.StopPrice, dMult, strFormat)
            
            Case eTT_OrderType_Limit
                astrOrder(8) = PfgPrice(Order.LimitPrice, dMult, strFormat)
                astrOrder(9) = ""
            
            Case eTT_OrderType_StopWithLimit
                astrOrder(8) = PfgPrice(Order.LimitPrice, dMult, strFormat)
                astrOrder(9) = PfgPrice(Order.StopPrice, dMult, strFormat)
                
        End Select
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = "CME"
        astrOrder(13) = ""
        If Order.Expiration = 0 Then
            astrOrder(14) = "1"
        Else
            astrOrder(14) = "0"
        End If
        astrOrder(15) = ""
        astrOrder(16) = ""
        astrOrder(17) = ""
        astrOrder(18) = ""
        astrOrder(19) = "TN-" & FormatVersion & "." & App.Revision
    
        'GenesisOrderToBrokerInfo Order, eTT_OrderStatus_Sent
        SendPfgMessage eGDPfgMessageType_AddOrder, astrOrder.JoinFields(vbTab)
        AddOrder = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.AddOrder"
    
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
    Dim strPfgSymbol As String          ' Pfg Symbol
    Dim strPfgOrderType As String       ' Pfg Order Type
    Dim dMult As Double                 ' Multiplier for converting prices
    Dim strFormat As String             ' Price format
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    strPfgSymbol = PfgSymbol(Order.Symbol, dMult, strFormat)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strPfgOrderType = PfgOrderType(Order.OrderType)
    
    If Len(strPfgSymbol) = 0 Then
        DumpDebug "AmendOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
    ElseIf Len(strPfgOrderType) = 0 Then
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
        astrOrder(1) = Parse(Order.BrokerID, "|", 2) ' Parse(Order.BrokerID, ":", 2)
        astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
        If IsForex(Order.Symbol) Then
            astrOrder(3) = strPfgSymbol
        Else
            astrOrder(3) = Left(strPfgSymbol, Len(strPfgSymbol) - 2)
        End If
        astrOrder(4) = strPfgSymbol
        astrOrder(5) = strPfgOrderType
        If Order.Buy Then astrOrder(6) = "1" Else astrOrder(6) = "2"
        
        ' Dave Rettger at PFG e-mailed me saying that this is causing at least some of the "Cancel Pending"
        ' order problems.  On an amend (especially when only the price is changing), we need to send the
        ' original order quantity instead of the remaining order quantity.  (01/15/2009 DAJ)...
        'astrOrder(7) = Str(Order.RemainingQuantity)
        astrOrder(7) = Str(Order.Quantity)
        
        Select Case Order.OrderType
            Case eTT_OrderType_Market
                astrOrder(8) = ""
                astrOrder(9) = ""
            
            Case eTT_OrderType_Stop
                astrOrder(8) = ""
                astrOrder(9) = PfgPrice(Order.StopPrice, dMult, strFormat)
            
            Case eTT_OrderType_Limit
                astrOrder(8) = PfgPrice(Order.LimitPrice, dMult, strFormat)
                astrOrder(9) = ""
            
            Case eTT_OrderType_StopWithLimit
                astrOrder(8) = PfgPrice(Order.LimitPrice, dMult, strFormat)
                astrOrder(9) = PfgPrice(Order.StopPrice, dMult, strFormat)
                
        End Select
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = "CME"
        astrOrder(13) = ""
        If Order.Expiration = 0 Then
            astrOrder(14) = "1"
        Else
            astrOrder(14) = "0"
        End If
        astrOrder(15) = ""
        astrOrder(16) = ""
        astrOrder(17) = ""
        astrOrder(18) = ""
        astrOrder(19) = "TN-" & FormatVersion & "." & App.Revision
    
        SendPfgMessage eGDPfgMessageType_AmendOrder, astrOrder.JoinFields(vbTab)
        AmendOrder = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.AmendOrder"
    
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
    Dim strPfgSymbol As String          ' Pfg Symbol
    Dim strPfgOrderType As String       ' Pfg Order Type
    Dim dMult As Double                 ' Multiplier for converting prices
    Dim strFormat As String             ' Price format
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    strPfgSymbol = PfgSymbol(Order.Symbol, dMult, strFormat)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strPfgOrderType = PfgOrderType(Order.OrderType)
        
    If NotSent(Order.Status) Then
        DumpDebug "CancelOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
        Order.ChangeOrderStatus eTT_OrderStatus_Cancelled, m.BrokerInfo
    ElseIf Len(strPfgSymbol) = 0 Then
        DumpDebug "CancelOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
    ElseIf Len(strPfgOrderType) = 0 Then
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
        astrOrder(1) = Parse(Order.BrokerID, "|", 2) ' Parse(Order.BrokerID, ":", 2)
        astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
        If IsForex(Order.Symbol) Then
            astrOrder(3) = strPfgSymbol
        Else
            astrOrder(3) = Left(strPfgSymbol, Len(strPfgSymbol) - 2)
        End If
        astrOrder(4) = strPfgSymbol
        astrOrder(5) = strPfgOrderType
        If Order.Buy Then astrOrder(6) = "1" Else astrOrder(6) = "2"
        astrOrder(7) = Str(Order.Quantity)
        Select Case Order.OrderType
            Case eTT_OrderType_Market
                astrOrder(8) = ""
                astrOrder(9) = ""
            
            Case eTT_OrderType_Stop
                astrOrder(8) = ""
                astrOrder(9) = PfgPrice(Order.StopPrice, dMult, strFormat)
            
            Case eTT_OrderType_Limit
                astrOrder(8) = PfgPrice(Order.LimitPrice, dMult, strFormat)
                astrOrder(9) = ""
            
            Case eTT_OrderType_StopWithLimit
                astrOrder(8) = PfgPrice(Order.LimitPrice, dMult, strFormat)
                astrOrder(9) = PfgPrice(Order.StopPrice, dMult, strFormat)
                
        End Select
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = "CME"
        astrOrder(13) = ""
        If Order.Expiration = 0 Then
            astrOrder(14) = "1"
        Else
            astrOrder(14) = "0"
        End If
        astrOrder(15) = ""
        astrOrder(16) = ""
        astrOrder(17) = ""
        astrOrder(18) = ""
        astrOrder(19) = "TN-" & FormatVersion & "." & App.Revision
    
        SendPfgMessage eGDPfgMessageType_CancelOrder, astrOrder.JoinFields(vbTab)
        
        Order.Status = eTT_OrderStatus_CancelPending
        Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
        Order.Save
        m.BrokerInfo.AddOrder Order, False
        CancelOrder = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.CancelOrder"
    
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
    Dim strPfgSymbol As String          ' Pfg Symbol
    Dim strPfgOrderType As String       ' Pfg Order Type
    Dim dMult As Double                 ' Multiplier for converting prices
    Dim strFormat As String             ' Price format
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    strPfgSymbol = PfgSymbol(Order.Symbol, dMult, strFormat)
    
    nPrevOrderType = Order.OrderType
    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    strPfgOrderType = PfgOrderType(Order.OrderType)
        
    If Len(strPfgSymbol) = 0 Then
        DumpDebug "ParkOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
    ElseIf Len(strPfgOrderType) = 0 Then
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
        astrOrder(1) = Parse(Order.BrokerID, "|", 2) ' Parse(Order.BrokerID, ":", 2)
        astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
        If IsForex(Order.Symbol) Then
            astrOrder(3) = strPfgSymbol
        Else
            astrOrder(3) = Left(strPfgSymbol, Len(strPfgSymbol) - 2)
        End If
        astrOrder(4) = strPfgSymbol
        astrOrder(5) = strPfgOrderType
        If Order.Buy Then astrOrder(6) = "1" Else astrOrder(6) = "2"
        astrOrder(7) = Str(Order.Quantity)
        Select Case Order.OrderType
            Case eTT_OrderType_Market
                astrOrder(8) = ""
                astrOrder(9) = ""
            
            Case eTT_OrderType_Stop
                astrOrder(8) = ""
                astrOrder(9) = PfgPrice(Order.StopPrice, dMult, strFormat)
            
            Case eTT_OrderType_Limit
                astrOrder(8) = PfgPrice(Order.LimitPrice, dMult, strFormat)
                astrOrder(9) = ""
            
            Case eTT_OrderType_StopWithLimit
                astrOrder(8) = PfgPrice(Order.LimitPrice, dMult, strFormat)
                astrOrder(9) = PfgPrice(Order.StopPrice, dMult, strFormat)
                
        End Select
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = "CME"
        astrOrder(13) = ""
        If Order.Expiration = 0 Then
            astrOrder(14) = "1"
        Else
            astrOrder(14) = "0"
        End If
        astrOrder(15) = ""
        astrOrder(16) = ""
        astrOrder(17) = ""
        astrOrder(18) = ""
        astrOrder(19) = "TN-" & FormatVersion & "." & App.Revision
    
        SendPfgMessage eGDPfgMessageType_CancelOrder, astrOrder.JoinFields(vbTab)
        
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
    RaiseError "cPFG.ParkOrder"
    
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
    RaiseError "cPFG.CancelOrderFromBrokerID"
    
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
    
    lReturn = 0&
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
    RaiseError "cPFG.CancelWorkingOrders"
    
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
        If lPosition < 0 Then
            strMessage = "BUY " & Abs(lPosition) & " " & strSymbol & " at the MARKET"
        Else
            strMessage = "SELL " & Abs(lPosition) & " " & strSymbol & " at the MARKET"
        End If
        
        If bConfirm Then
            strReturn = InfBox("Are you sure you want to|" & strMessage & "|in account '" & strAccount & "'|to flatten your position?", "?", "+Yes|-No", "Flatten for " & strSymbol & " in " & strAccount)
        Else
            strReturn = "Y"
        End If
        
        If strReturn = "Y" Then
            bReturn = True
            If bConfirm Then
                DumpDebug "User Informed: " & strMessage & " in account " & strAccount
            Else
                DumpDebug "User Not Informed: " & strMessage & " in account " & strAccount
            End If
            ExitPositionForSymbol strAccount, strSymbol, lAtID, lPosition, strGenesisOrderID
        End If
    End If
    
    FlattenPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.FlattenPosition"
    
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
        If lPosition < 0 Then
            strMessage = "BUY " & Abs(lPosition * 2) & " " & strSymbol & " at the MARKET"
        Else
            strMessage = "SELL " & Abs(lPosition * 2) & " " & strSymbol & " at the MARKET"
        End If
        
        If bConfirm Then
            strReturn = InfBox("Are you sure you want to|" & strMessage & "|in account '" & strAccount & "'|to reverse your position?", "?", "+Yes|-No", "Reverse for " & strSymbol & " in " & strAccount)
        Else
            strReturn = "Y"
        End If
        
        If strReturn = "Y" Then
            bReturn = True
            If bConfirm Then
                DumpDebug "User Informed: " & strMessage & " in account " & strAccount
            Else
                DumpDebug "User Not Informed: " & strMessage & " in account " & strAccount
            End If
            ReversePositionForSymbol strAccount, strSymbol, lAtID, lPosition, strGenesisOrderID
        End If
    End If
    
    ReversePosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.ReversePosition"
    
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
        SendPfgMessage eGDPfgMessageType_GetAccounts, ""
    Else
        DumpDebug "Account Refresh Requested but already in Progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.GetAccounts"
    
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

    If Not Refreshing Then
        If Len(strAccount) = 0 Then DumpDebug "Getting Orders" Else DumpDebug "Getting Orders for " & strAccount
        m.bRefreshingOrders = True
        m.bRefreshingFills = True
        m.bRefreshingPositions = True
        
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for order information...", , , m.strBrokerName & " Refresh", True
        End If
        SendPfgMessage eGDPfgMessageType_GetOrders, strAccount
    Else
        DumpDebug "Order Refresh Requested but already in Progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.GetOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSingleOrder
'' Description: Ask the PFG servers for the status on a single order
'' Inputs:      Order Number
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetSingleOrder(ByVal strOrderNumber As String)
On Error GoTo ErrSection:

    If Not Refreshing Then
        ' PFG does not want the account number portion of the order number when
        ' requesting a single order refresh, so we need to take only the order
        ' number portion (06/27/2008 DAJ)...
        If InStr(strOrderNumber, "-") <> 0 Then
            strOrderNumber = Parse(strOrderNumber, "-", 2)
        End If
        
        DumpDebug "Getting Single Order: " & strOrderNumber
        SendPfgMessage eGDPfgMessageType_GetSingleOrder, strOrderNumber
    Else
        DumpDebug "Single order refresh asked for, but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.GetSingleOrder"
    
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
        m.bRefreshingPositions = True
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for position information...", , , m.strBrokerName & " Refresh", True
        End If
        If Len(strAccount) = 0 Then
            DumpDebug "Getting Positions"
            For lIndex = 0 To m.astrAccounts.Size - 1
                SendPfgMessage eGDPfgMessageType_GetPositions, m.astrAccounts(lIndex)
            Next lIndex
        Else
            DumpDebug "Getting Positions for " & strAccount
            SendPfgMessage eGDPfgMessageType_GetPositions, strAccount
        End If
    Else
        DumpDebug "Position Refresh Requested but already in Progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBlocks
'' Description: Ask the server for block information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetBlocks()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.GetBlocks"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccountInfo
'' Description: Ask the server for account information
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAccountInfo(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    If Len(strAccount) = 0 Then
        DumpDebug "Getting Account Information"
        For lIndex = 0 To m.astrAccounts.Size - 1
            SendPfgMessage eGDPfgMessageType_GetAccountInfo, m.astrAccounts(lIndex)
        Next lIndex
    Else
        DumpDebug "Getting Account Information for " & strAccount
        SendPfgMessage eGDPfgMessageType_GetAccountInfo, strAccount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.GetAccountInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LinkOrders
'' Description: Link two orders together in an Order-Cancel-Order
'' Inputs:      Order One, Order Two
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LinkOrders(ByVal strOrderNumberOne As String, ByVal strOrderNumberTwo As String)
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message to send to PFG

    DumpDebug "Linking " & strOrderNumberOne & " and " & strOrderNumberTwo
    strMessage = Parse(strOrderNumberOne, "|", 2) & vbTab & Parse(strOrderNumberTwo, "|", 2)
    SendPfgMessage eGDPfgMessageType_LinkOrders, strMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.LinkOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnlinkOrders
'' Description: Unlink two orders that were in an Order-Cancel-Order
'' Inputs:      Order One, Order Two
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UnlinkOrders(ByVal strOrderNumberOne As String, ByVal strOrderNumberTwo As String)
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message to send to PFG

    DumpDebug "Unlinking " & strOrderNumberOne & " and " & strOrderNumberTwo
    strMessage = Parse(strOrderNumberOne, "|", 2) & vbTab & Parse(strOrderNumberTwo, "|", 2)
    SendPfgMessage eGDPfgMessageType_UnlinkOrders, strMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.UnlinkOrders"
    
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
    RaiseError "cPFG.Refresh"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateSpread
'' Description: Create a spread with the given information
'' Inputs:      Spread Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateSpread(ByVal strSpreadInformation As String)
On Error GoTo ErrSection:

    Dim astrLegs As New cGdArray        ' Array of information for the spread legs
    Dim astrLeg As New cGdArray         ' Array of information for a single leg
    Dim lIndex As Long                  ' Index into a for loop
    Dim strMessage As String            ' Message to send to the PFG servers
    Dim strPfgBase As String            ' PFG Base symbol for the leg symbol
    Dim strPfgSymbol As String          ' PFG Symbol for the Genesis leg symbol
    
    If ConnectionStatus = eGDConnectionStatus_Connected Then
        astrLegs.SplitFields strSpreadInformation, ";"
        For lIndex = 0 To astrLegs.Size - 1
            astrLeg.SplitFields astrLegs(lIndex), ","
            
            strPfgSymbol = PfgSymbol(astrLeg(2), , , strPfgBase)
                        
            astrLeg(2) = strPfgBase
            astrLeg(3) = strPfgSymbol
            
            astrLegs(lIndex) = astrLeg.JoinFields(",")
        Next lIndex
        
        strMessage = m.astrAccounts(0) & vbTab & NextGenesisID(m.astrAccounts(0)) & vbTab & astrLegs.JoinFields(vbTab)
        DumpDebug "Requesting Creation of a Spread: '" & strMessage & "'"
        SendPfgMessage eGDPfgMessageType_CreateSpread, strMessage
    Else
        frmTest2.AddList "Not connected to " & m.strBrokerName
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.CreateSpread"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RequestSpreadQuote
'' Description: Request a spread for the given symbol
'' Inputs:      Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RequestSpreadQuote(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message to send to the PFG servers

    If ConnectionStatus = eGDConnectionStatus_Connected Then
        strMessage = m.astrAccounts(0) & vbTab & NextGenesisID(m.astrAccounts(0)) & vbTab & strSymbol
        DumpDebug "Requesting Quote for Spread Symbol: '" & strMessage & "'"
        SendPfgMessage eGDPfgMessageType_RequestForQuote, strMessage
    Else
        frmTest2.AddList "Not connected to " & m.strBrokerName
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.RequestSpreadQuote"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RequestSpreadCodes
'' Description: Request spread codes for the given information
'' Inputs:      Symbols, Type, Ratios, Sides
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RequestSpreadCodes(ByVal strSymbols As String, ByVal strType As String, ByVal strRatios As String, ByVal strSides As String)
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message to send to the PFG servers
    Dim astrSymbols As New cGdArray     ' Array of symbols
    Dim astrSpread As New cGdArray      ' Spread
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop

    If ConnectionStatus = eGDConnectionStatus_Connected Then
        astrSymbols.SplitFields strSymbols, ","
        For lIndex = 0 To astrSymbols.Size - 1
            astrSpread.SplitFields astrSymbols(lIndex), ";"
            For lIndex2 = 0 To astrSpread.Size - 1
                astrSpread(lIndex2) = PfgSymbol(astrSpread(lIndex2))
            Next lIndex2
            
            astrSymbols(lIndex) = NextGenesisID(m.astrAccounts(0)) & ";" & astrSpread.JoinFields(";")
        Next lIndex
        
        strMessage = astrSymbols.JoinFields(";") & vbTab & strType & vbTab & strRatios & vbTab & strSides & vbTab & NextGenesisID(m.astrAccounts(0))
        DumpDebug "Requesting Spread Codes: '" & strMessage & "'"
        SendPfgMessage eGDPfgMessageType_RequestSpreadCodes, strMessage
    Else
        frmTest2.AddList "Not connected to " & m.strBrokerName
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.RequestSpreadCodes"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RequestSpreadLegs
'' Description: Request spread legs for the given symbol
'' Inputs:      Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RequestSpreadLegs(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message to send to the PFG servers

    If ConnectionStatus = eGDConnectionStatus_Connected Then
        strMessage = strSymbol & vbTab & NextGenesisID(m.astrAccounts(0))
        DumpDebug "Requesting Legs for Spread Symbol: '" & strMessage & "'"
        SendPfgMessage eGDPfgMessageType_RequestSpreadLegs, strMessage
    Else
        frmTest2.AddList "Not connected to " & m.strBrokerName
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.RequestSpreadLegs"
    
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
    RaiseError "cPFG.ConnectionStatusForAccount"
    
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
    
    NextGenesisID = "TN" & Format(lCustomerID, "000000") & Format(lDate, "00000") & Format(lCounter, "00000")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.NextGenesisID"
    
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
    
    StartingGenesisIdForOptNav = "TN" & Format(lCustomerID, "000000") & Format(Date, "00000") & "90000"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.StartingGenesisIdForOptNav"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendPfgMessage
'' Description: Send an App Mail message to the PFG program
'' Inputs:      Type of Message, Message, Send Now?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendPfgMessage(ByVal nType As eGDPfgMessageTypes, ByVal strMessage As String, Optional ByVal bSendNow As Boolean = True)
On Error GoTo ErrSection:

    Dim astrMessage As New cGdArray     ' Message split out into an array

    If nType = eGDPfgMessageType_Connect Then
        astrMessage.SplitFields strMessage, vbTab
        astrMessage(1) = SetStrToChar(astrMessage(1), "*")
        DumpDebug vbTab & "Sending " & m.strBrokerName & " Message (" & MessageTypeToString(nType) & "): " & astrMessage.JoinFields(vbTab)
    Else
        DumpDebug vbTab & "Sending " & m.strBrokerName & " Message (" & MessageTypeToString(nType) & "): " & strMessage
    End If
    
    frmOnlineBroker.gdBroker.CreateMessage m.strControlID, nType, strMessage, , bSendNow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.SendPfgMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming App Mail message from the PFG program
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDPfgMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim frm As frmAlertPopup            ' Alert popup

    If (nType <> eGDPfgMessageType_ConnectionInfo) And (nType <> eGDPfgMessageType_Heartbeat) Then
        DumpDebug vbTab & m.strBrokerName & " Message Received (" & MessageTypeToString(nType) & "): " & strMessage
    End If
    
    Select Case nType
        Case eGDPfgMessageType_ConnectionInfo
            HandleConnectionInfo strMessage
        
        Case eGDPfgMessageType_AppLoaded
            m.bAppLoaded = True
            m.bManualDisconnect = False
            DumpDebug "ManualDisconnect = " & Str(m.bManualDisconnect) & " (AppLoaded)"
            
            If LiveTradingAllowed(m.nBroker) Then
                DumpDebug "Connecting to " & m.strBrokerName & " as " & m.strUserName & " (" & m.strHostIP & ":" & m.strHostPort & ")"
                SendPfgMessage eGDPfgMessageType_Connect, m.strUserName & vbTab & m.strPassword & vbTab & m.strHostIP & vbTab & m.strHostPort & vbTab & "GBT" & vbTab & "TS" & vbTab & FormatVersion & "." & App.Revision
            Else
                DumpDebug "Unloading the application since live trading is not allowed"
                SendPfgMessage eGDPfgMessageType_UnloadApp, ""
            End If
        
        Case eGDPfgMessageType_AppUnloaded
            m.bAppLoaded = False
            If m.bManualDisconnect Then
                DumpDebug "Heartbeat Cleared -- AppUnloaded"
                m.dHeartBeat = 0#
            Else
                m.dHeartBeat = gdTickCount - 110000
            End If
            m.bManualDisconnect = False
            DumpDebug "ManualDisconnect = " & Str(m.bManualDisconnect) & " (AppUnloaded)"
            ConnectionStatus = eGDConnectionStatus_Disconnected
        
        Case eGDPfgMessageType_Heartbeat
            m.dHeartBeat = gdTickCount
            If m.dLastFillTime > 0 Then
                If gdTickCount > m.dLastFillTime + 10000 Then
                    GetPositions
                    m.dLastFillTime = 0#
                End If
            End If
        
        Case eGDPfgMessageType_AccountR
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
                HandleAccount strMessage
                m.BrokerInfo.AddAccount strMessage
                m.BrokerInfo.SetConnectionStatusForAccount eGDConnectionStatus_Connected, Parse(strMessage, vbTab, 1)
            End If
        
        Case eGDPfgMessageType_Order
            m.dHeartBeat = gdTickCount
            'UpdateBrokerInfoOrder strMessage, False
            HandleOrderInfo strMessage, False, False
        
        Case eGDPfgMessageType_OrderR
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingOrders = True
                m.BrokerInfo.ClearOrders
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " order information...", , , m.strBrokerName & " Refresh", True
                End If
                SendBeginToOptionNav eGDOptNav_Order, m.nBroker
            ElseIf UCase(strMessage) = "END" Then
                m.BrokerInfo.RefreshingOrders = False
                m.bRefreshingOrders = False
                SendEndToOptionNav eGDOptNav_Order, m.nBroker
            Else
                'UpdateBrokerInfoOrder strMessage, True
                HandleOrderInfo strMessage, True, False
            End If
            
        Case eGDPfgMessageType_FillR
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingFills = True
                m.BrokerInfo.ClearFills
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " fill information...", , , m.strBrokerName & " Refresh", True
                End If
                SendBeginToOptionNav eGDOptNav_Fill, m.nBroker
            ElseIf UCase(strMessage) = "END" Then
                SendEndToOptionNav eGDOptNav_Fill, m.nBroker
                
                m.BrokerInfo.RefreshingFills = False
                m.bRefreshingFills = False
            Else
                'AddFillToBrokerInfo strMessage, True
                HandleOrderInfo strMessage, True, True
            End If
        
        Case eGDPfgMessageType_PositionR
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingPositions = True
                m.BrokerInfo.ClearPositions
                m.astrBrokerPos.Clear
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " position information...", , , m.strBrokerName & " Refresh", True
                End If
            ElseIf UCase(strMessage) = "END" Then
                m.BrokerInfo.RefreshingPositions = False
                If m.bVerboseRefresh Then
                    InfBox "Synchronizing " & m.strBrokerName & " position information...", , , m.strBrokerName & " Refresh", True
                End If
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
            
        Case eGDPfgMessageType_BlockR
            m.dHeartBeat = gdTickCount
        
        Case eGDPfgMessageType_AccountInfo
            m.dHeartBeat = gdTickCount
            HandleAccountInfo strMessage
            
        Case eGDPfgMessageType_OrdersLinked
            m.dHeartBeat = gdTickCount
            HandleLinkedOrders strMessage
        
        Case eGDPfgMessageType_OrdersUnlinked
            m.dHeartBeat = gdTickCount
            HandleUnlinkedOrders strMessage
            
        Case eGDPfgMessageType_SpreadCreated
            m.dHeartBeat = gdTickCount
            frmTest2.AddList m.strBrokerName & " Spread Create Response: '" & strMessage & "'"
        
        Case eGDPfgMessageType_SpreadQuote
            m.dHeartBeat = gdTickCount
            frmTest2.AddList m.strBrokerName & " Spread Quote Response: '" & strMessage & "'"
            
        Case eGDPfgMessageType_SpreadCodes
            m.dHeartBeat = gdTickCount
            frmTest2.AddList m.strBrokerName & " Spread Codes Response: '" & strMessage & "'"
            
        Case eGDPfgMessageType_SpreadLegs
            m.dHeartBeat = gdTickCount
            frmTest2.AddList m.strBrokerName & " Spread Legs Response: '" & strMessage & "'"
            
    End Select
    
    If Refreshing And (m.dRefreshingStart = 0#) Then
        m.dRefreshingStart = gdTickCount
    ElseIf Refreshing And (gdTickCount > m.dRefreshingStart + 30000#) Then
        If m.BrokerInfo.ConnectedAccounts.Size = 0 Then
            Set frm = New frmAlertPopup
            frm.ShowMessageBox "No response was received from the " & m.strBrokerName & " servers for a refresh request.  Trade Navigator will try to reconnect to the " & m.strBrokerName & " servers...", m.strBrokerName & " Communications Timeout", vbCenter
            m.dHeartBeat = gdTickCount - 110000
            Disconnect True, "Reconnecting because of lack of response"
        End If
        
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
    RaiseError "cPFG.HandleMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckHeartbeat
'' Description: Check the heartbeat to make sure that we are still connected
''              to the PFG servers
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
    RaiseError "cPFG.CheckHeartbeat"
    
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
    Dim strAmPm As String               ' AM/PM porition of the string
    Dim dReturn As Double               ' Return value of the function
    Dim iMonth As Integer               ' Month
    Dim strLongDate As String           ' Long date (YYYYMMDD)
    
    dReturn = 0#
    If Len(strDateTime) > 0 Then
        strDate = Parse(strDateTime, " ", 1)
        If InStr(strDate, "-") <> 0 Then
            strTime = Parse(strDateTime, " ", 3)
            strAmPm = Parse(strDateTime, " ", 4)
            
            If Len(Parse(strDate, "-", 3)) = 2 Then
                strLongDate = "20" & Parse(strDate, "-", 3) & Format(MonthNum(Parse(strDate, "-", 2)), "00") & Format(Parse(strDate, "-", 1), "00")
            Else
                strLongDate = Parse(strDate, "-", 3) & Format(MonthNum(Parse(strDate, "-", 2)), "00") & Format(Parse(strDate, "-", 1), "00")
            End If
        Else
            If InStr(strDateTime, "@") <> 0 Then
                strTime = Parse(strDateTime, " ", 3)
                strAmPm = Parse(strDateTime, " ", 4)
            Else
                strTime = Parse(strDateTime, " ", 2)
                strAmPm = Parse(strDateTime, " ", 3)
            End If
            
            If Len(Parse(strDate, "/", 3)) = 2 Then
                strLongDate = "20" & Parse(strDate, "/", 3) & Format(Parse(strDate, "/", 1), "00") & Format(Parse(strDate, "/", 2), "00")
            Else
                strLongDate = Parse(strDate, "/", 3) & Format(Parse(strDate, "/", 1), "00") & Format(Parse(strDate, "/", 2), "00")
            End If
        End If
        
        ' 08/06/2009 DAJ: The CDate call below crashes if time is blank, so default it to midnight for now...
        If Len(strTime) = 0 Then
            strTime = "12:00:00"
            strAmPm = "AM"
            
        ' 09/24/2010 DAJ: The CDate call below crashes if milliseconds are included, so get rid of them for now...
        ElseIf InStr(strTime, ".") > 0 Then
            strTime = Parse(strTime, ".", 1)
        End If
        
        dReturn = DateOf(strLongDate) + CDate(strTime & " " & strAmPm)
    End If
    
    DateToDouble = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.DateToDouble"
    
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
    RaiseError "cPFG.DumpDebug"
    
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

    If (ConnectionStatus = eGDConnectionStatus_Connected) And (m.astrAccounts.Size > 0) And (m.bSyncInProgress = False) Then
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
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.VerifyPositions"
    
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

    IsTradeableSymbol = (Len(PfgSymbol(strGenesisSymbol)) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.IsTradeableSymbol"
    
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

    IsEnabledSymbol = (Len(PfgSymbol(strGenesisSymbol, , , strBrokerBase)) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.IsEnabledSymbol"
    
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
        If IsForex(strGenesisSymbol) Then
            strBaseSymbol = strGenesisSymbol
        ElseIf UCase(SecurityType(strGenesisSymbol)) = "S" Then
            strBaseSymbol = "!"
        Else
            strBaseSymbol = Parse(strGenesisSymbol, "-", 1)
        End If
        
        For lIndex = 0 To m.astrSymbols.Size - 1
            If Parse(m.astrSymbols(lIndex), vbTab, 2) = strBaseSymbol Then
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
    RaiseError "cPFG.OrderTypeAllowed"
    
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
    
    If IsForex(strGenesisSymbol) Then
        strBaseSymbol = strGenesisSymbol
    ElseIf UCase(SecurityType(strGenesisSymbol)) = "S" Then
        strBaseSymbol = "!"
    Else
        strBaseSymbol = Parse(strGenesisSymbol, "-", 1)
    End If
    
    For lIndex = 0 To m.astrSymbols.Size - 1
        If Parse(m.astrSymbols(lIndex), vbTab, 2) = strBaseSymbol Then
            strMask = Parse(m.astrSymbols(lIndex), vbTab, 6)
            If Len(strMask) = 0 Then strMask = "110"
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
    RaiseError "cPFG.TimeInForceAllowed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TickValuePerMove
'' Description: Determine the tick value per move for the given symbol
'' Inputs:      Symbol, Bars
'' Returns:     Tick Value per Move for given symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TickValuePerMove(ByVal vSymbolOrSymbolID As Variant, Optional Bars As cGdBars = Nothing) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function

    If Bars Is Nothing Then
        Set Bars = New cGdBars
        SetBarProperties Bars, vSymbolOrSymbolID
    End If
    
    If IsForex(GetSymbol(vSymbolOrSymbolID)) Then
        dReturn = GetIniFileProperty("Forex", 10000#, "PipValue", m.strConnectIni)
    Else
        dReturn = Bars.TickValuePerMove
    End If
    
    TickValuePerMove = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.TickValuePerMove"
    
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
    Static bInProgress As Boolean       ' Are we already in progress?
    
    If bInProgress = False Then
        bInProgress = True
        
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
        
        bInProgress = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "cPFG.FixPositions"
    
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
        
        ' If we have a carried position mismatch...
        If lBrokerCarried <> FillSummary.CurrentPosition Then
            DumpDebug "Carried Position Mismatch for " & strSymbol & " in " & strAccount & ": " & m.strBrokerName & " = " & Str(lBrokerCarried) & " ; Trade Navigator = " & Str(FillSummary.CurrentPosition)
            
            ' If the broker says that the user did not carry a position, but we do...
            If (lBrokerCarried = 0&) And (FillSummary.CurrentPosition <> 0&) Then
                ' Create a "fake fill" for yesterday's session at yesterday's closing price to flatten our position...
                CreateFillAtPreviousClose vAccountNumberOrID, vSymbolOrSymbolID, FillSummary
                g.Broker.RebuildFillSummaryForSymbol FillSummary.AccountID, FillSummary.SymbolOrSymbolID, 0&, True
                
            ' If the broker says that the user did carry a position, but we do not...
            ElseIf (lBrokerCarried <> 0&) And (FillSummary.CurrentPosition = 0&) Then
                ' Create a fake fill for yesterday's session with the carried position info from the broker...
                CreateFillsForCarriedInfo vAccountNumberOrID, vSymbolOrSymbolID, FillSummary
                g.Broker.RebuildFillSummaryForSymbol FillSummary.AccountID, FillSummary.SymbolOrSymbolID, 0&, True
            
            ' If we both say that the user carried a position, but we don't agree...
            Else
                ' Create a "fake fill" for yesterday's session at yesterday's fill price to flatten our position...
                CreateFillAtPreviousClose vAccountNumberOrID, vSymbolOrSymbolID, FillSummary
                
                ' Create a fake fill for yesterday's session with the carried position info from the broker...
                CreateFillsForCarriedInfo vAccountNumberOrID, vSymbolOrSymbolID, FillSummary
                g.Broker.RebuildFillSummaryForSymbol FillSummary.AccountID, FillSummary.SymbolOrSymbolID, 0&, True
            End If
        
            m.BrokerInfo.CarriedMatch(strAccount, FillSummary.Symbol) = bReturn
        End If
        
#If 0 Then
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
                        g.TradingItems.DisableTradeItemsForSymbol FillSummary.AccountID, FillSummary.SymbolOrSymbolID, "Position mismatch"
                    End If
                    
                    m.astrToFix(lPos) = strAccount & vbTab & strSymbol & vbTab & "1"
                End If
            End If
        End If
#End If
    End If
    
    FixPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.FixPosition"
        
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
    Dim strReturn As String             ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim bPut As Boolean                 ' Is this a put or a call?
    Dim dStrike As Double               ' Strike price
    Dim strFuture As String             ' Future symbol
    Dim OptSym As New cOptionSymbol     ' Option symbol object
    
    strReturn = ""
    If Len(strGenesisSymbol) > 0 Then
        Select Case UCase(SecurityType(strGenesisSymbol, True))
            Case "I"
                If IsForex(strGenesisSymbol) Then
                    For lIndex = 0 To m.astrSymbols.Size - 1
                        If Parse(m.astrSymbols(lIndex), vbTab, 2) = strGenesisSymbol Then
                            strReturn = Parse(m.astrSymbols(lIndex), vbTab, 5) & vbTab & Parse(m.astrSymbols(lIndex), vbTab, 6)
                            Exit For
                        End If
                    Next lIndex
                End If
            
            Case "FO"
                OptSym.FromGenesis strGenesisSymbol
                If Len(OptSym.BaseSymbol) > 0 Then
                    For lIndex = 0 To m.astrSymbols.Size - 1
                        If Parse(m.astrSymbols(lIndex), vbTab, 2) = "O:" & Replace(OptSym.BaseSymbol, "-", "") Then
                            strReturn = Parse(m.astrSymbols(lIndex), vbTab, 5) & vbTab & Parse(m.astrSymbols(lIndex), vbTab, 6)
                            Exit For
                        End If
                    Next lIndex
                End If
                
            Case "F"
                strBase = Parse(strGenesisSymbol, "-", 1)
                
                For lIndex = 0 To m.astrSymbols.Size - 1
                    If Parse(m.astrSymbols(lIndex), vbTab, 2) = strBase Then
                        strReturn = Parse(m.astrSymbols(lIndex), vbTab, 5) & vbTab & Parse(m.astrSymbols(lIndex), vbTab, 6)
                        Exit For
                    End If
                Next lIndex
                
            Case "S"
                strBase = "!"
                
                For lIndex = 0 To m.astrSymbols.Size - 1
                    If Parse(m.astrSymbols(lIndex), vbTab, 2) = strBase Then
                        strReturn = Parse(m.astrSymbols(lIndex), vbTab, 5) & vbTab & Parse(m.astrSymbols(lIndex), vbTab, 6)
                        Exit For
                    End If
                Next lIndex
                
            Case "SO"
                strBase = "@"
                
                For lIndex = 0 To m.astrSymbols.Size - 1
                    If Parse(m.astrSymbols(lIndex), vbTab, 2) = strBase Then
                        strReturn = Parse(m.astrSymbols(lIndex), vbTab, 5) & vbTab & Parse(m.astrSymbols(lIndex), vbTab, 6)
                        Exit For
                    End If
                Next lIndex
                
        End Select
    End If
    
    SymbolInformation = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.SymbolInformation"
    
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
                        DumpDebug "Heartbeat Cleared -- HandleConnectionInfo (1)"
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
                    SendPfgMessage eGDPfgMessageType_UnloadApp, ""
                End If
            
            Case eGDConnectionStatus_Disconnecting
                m.BrokerInfo.SetConnectionStatusForUser eGDConnectionStatus_Disconnecting, strLastUser
            
            Case eGDConnectionStatus_Connecting
                m.BrokerInfo.SetConnectionStatusForUser eGDConnectionStatus_Connecting, m.strUserName
            
            Case eGDConnectionStatus_Connected
                m.bWasConnected = True
        
        End Select
        
        If m.strUserName <> strLastUser Then
            strLastUser = m.strUserName
        End If
        
        If (nStatus <> ConnectionStatus) Or (Len(strError) > 0) Then
            nPrev = ConnectionStatus
            ConnectionStatus = nStatus
                        
            g.Broker.ConnectionStatusChanged m.nBroker, nStatus
        
            If (nPrev <> eGDConnectionStatus_Connected) And (ConnectionStatus = eGDConnectionStatus_Connected) Then
                ' DAJ 09/16/2009: On a successful connection, store the access key information
                ' in the registry...
                RI_SetPfgAccessKey m.strUserName, m.strPassword
                
                m.lLastConnectDate = Date
                GetAccounts
            End If
            DumpDebug vbTab & m.strBrokerName & " Message Received (" & MessageTypeToString(eGDPfgMessageType_ConnectionInfo) & "): " & strConnectionInfo
        End If
        
        If (Len(strError) > 0) And (strError <> m.strLastMessage) Then
            If InStr(UCase(strError), "INVALID LOGIN") <> 0 Then
                m.bManualDisconnect = True
                DumpDebug "ManualDisconnect = " & Str(m.bManualDisconnect) & " (HandleConnectionInfo)"
                DumpDebug "Heartbeat Cleared -- HandleConnectionInfo (2)"
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
    RaiseError "cPFG.HandleConnectionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleAccount
'' Description: Do the appropriate thing with the given account information
'' Inputs:      Account Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleAccount(ByVal strAccountInfo As String)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strAccountNumber As String      ' Account Number from the callback
    Dim lAccountID As Long              ' Account ID
    Dim lPos As Long                    ' Position of account in the array
    Dim astrFlags As New cGdArray       ' Account flags
    Dim Acct As cPtAccount              ' Account object

    If Len(Trim(strAccountInfo)) > 0 Then
        strAccountNumber = Parse(strAccountInfo, vbTab, 1)
        astrFlags.SplitFields Parse(strAccountInfo, vbTab, 2), ","
        
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
            rs!SecTypeMask = 31
            rs!FillRT = False
            rs!ClosedProfit = 0#
            rs!HoldOcoAtBroker = GetIniFileProperty(strAccountNumber, False, "BrokerOCO", g.strIniFile)
            lAccountID = rs!AccountID
            rs.Update
        Else
            lAccountID = rs!AccountID
            Set Acct = New cPtAccount
            
            If Acct.Load(lAccountID) Then
                If Acct.AccountType <> m.nBroker Then
                    If g.Broker.IsPfgBroker(Acct.AccountType) Then
                        Acct.AccountType = m.nBroker
                        Acct.Save
                        g.Broker.UpdateAccountCache Acct
                    End If
                End If
            End If
        End If
        
        ' DAJ 08/13/2010: Only update the last connected time for an account that does not stipulate
        ' delayed data.  Otherwise, if this account is set for delayed streaming and all they have
        ' are demo accounts in their database, clear the last connected flag so that they will be
        ' delayed the next time they connect to streaming...
        If Len(astrFlags(3)) = 0 Then
            SetIniFileProperty "Last", EncryptToHex(Str(m.lLastConnectDate)), "Connect", m.strIniFile
        ElseIf HasNonDemoAccount = False Then
            SetIniFileProperty "Last", "", "Connect", m.strIniFile
        End If
        
        If m.astrAccounts.BinarySearch(strAccountNumber, lPos) = False Then
            m.astrAccounts.Add strAccountNumber, lPos
        End If
        
        RefreshAccount lAccountID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.HandleAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleOrderInfo
'' Description: Do the appropriate thing with the given order information
'' Inputs:      Order Information, Refresh, From Fill?
'' Returns:     None
''
'' RecordFrom:  Genesis ID, Order ID, Fill ID, Account, Base Symbol, Symbol,
''              Order Type, B/S, Quantity, Limit, Stop, Strike, Expiration,
''              Exchange, Instrument, Time In Force, Maturity, Reserved,
''              Quantity Filled, Remaining Quantity, Quantity Cum, Fill Price,
''              Fill Time, Event, Status, Message, Position ID, Open/Close,
''              Linked Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleOrderInfo(ByVal strOrderInfo As String, ByVal bRefresh As Boolean, ByVal bFromFillR As Boolean)
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
    Dim bDupFill As Boolean             ' Is this a duplicate fill?
    Dim frm As frmAlertPopup            ' Alert popup form
    Dim strPrevBrokerID As String       ' Previous broker ID
    Dim bNewOrder As Boolean            ' Is this a new order?
    Dim strFillID As String             ' Fill ID
    Dim nOrderType As eTT_OrderType     ' Order Type
    Dim iWeekNum As Integer             ' Week number
    Dim strCurrent As String            ' Order number with current week number
    Dim strPrev As String               ' Order number with previous week number
    Dim bExpire As Boolean              ' Is this an expiration notice?
    Dim bUnsolicitedCancel As Boolean   ' Is this an unsolicited cancel?
    
    If Len(strOrderInfo) > 0 Then
        astrOrder.SplitFields strOrderInfo, vbTab
        
        bDupFill = DuplicateFill(astrOrder(2))
        If bDupFill Then
            DumpDebug "Duplicate fill flag -- requesting a refresh on this order"
            GetSingleOrder astrOrder(1)
        Else
            strGenesisSymbol = GenesisSymbol(astrOrder(5), dMult, , False)
            If Len(astrOrder(3)) > 0 Then
                strAccountNumber = astrOrder(3)
            Else
                strAccountNumber = Parse(astrOrder(1), "-", 1)
            End If
            
            ' Need a Genesis symbol and need there to be an order/fill date in the record...
            If (Len(strGenesisSymbol) > 0) And (Len(astrOrder(22)) > 0) Then
                bExpire = (InStr(UCase(astrOrder(25)), "EXPIRED") <> 0)
                bUnsolicitedCancel = (UCase(astrOrder(23)) = "CANCEL") And (UCase(astrOrder(24)) = "CANCELED") And (bRefresh = False)
                SetBarProperties Bars, strGenesisSymbol
                
                ' The time that comes back from the PFG servers has the session date instead of the logical
                ' date, so we need to adjust the date if it occurs after 5 minutes after the End Time of
                ' the session...
                ' 10/15/2009 DAJ: It appears that expiration notices come after the close of
                ' the session, but have the calendar date instead of the session date, so we don't
                ' want to subtract one off of the date...
                ' 02/11/2011 DAJ: Fills that come in from a refresh are also coming in with the
                ' calendar date instead of the session date, so we don't want to subtract one
                ' off of the date...
                ' 05/04/2011 DAJ: It also seems that orders that come in a "35=P" order confirmation have
                ' the calendar date instead of the session date, so we don't want to subtract one
                ' off of the date...
                If bExpire Or bFromFillR Or (UCase(astrOrder(24)) = "QUEUED") Then
                    dFillTime = DateToDouble(astrOrder(22))
                Else
                    dFillTime = AdjustDateTime(DateToDouble(astrOrder(22)), strGenesisSymbol)
                End If
                
                ' PFG's Order ID is the account number followed by a six digit number that wraps around, so
                ' a user could get a duplicate order ID over time when the six digit number wraps around (I
                ' have actually seen it happen).  To get around this, we will prepend the date...
                ' Example: 20080825:A12345-123456
                'astrOrder(1) = Format(dFillTime, "YYYYMMDD") & ":" & astrOrder(1)
                
                ' The previous fix was alright, but the calendar date portion was goofing things up if
                ' the user placed an order after the session reopened at night and then the order filled
                ' the next day.  Tim's suggestion here is to use the week number instead.  As long as the
                ' user doesn't carry a GTC order over the weekend, this will work fine...
                ' Example: 05694|A12345-123456
                iWeekNum = WkNum(Int(dFillTime))
                strCurrent = Format(iWeekNum, "00000") & "|" & astrOrder(1)
                astrOrder(1) = strCurrent
                
                Set Order = LoadOrder(strAccountNumber, astrOrder(0), astrOrder(1), lOldOrderID)
                
                ' If the order wasn't found and it was a GTC order, see if the order was stored with
                ' the previous week number -- go as far as 5 weeks back (04/07/2009 DAJ)...
                If (Order Is Nothing) And (UCase(astrOrder(15)) = "GTC") Then
                    For lIndex = 1 To 5
                        strPrev = Format(iWeekNum - lIndex, "00000") & "|" & Parse(astrOrder(1), "|", 2)
                        DumpDebug "Order " & astrOrder(1) & ";" & astrOrder(0) & " not found -- trying " & strPrev
                        astrOrder(1) = strPrev
                        Set Order = LoadOrder(strAccountNumber, astrOrder(0), astrOrder(1), lOldOrderID)
                        If Not Order Is Nothing Then
                            Exit For
                        End If
                    Next lIndex
                End If
                
                If Order Is Nothing Then
                    astrOrder(1) = strCurrent
                    DumpDebug "Order " & astrOrder(1) & ";" & astrOrder(0) & " not found -- creating new order"
                    Set Order = New cPtOrder
                    bNewOrder = True
                Else
                    DumpDebug "Order " & astrOrder(1) & ";" & astrOrder(0) & " found"
                    bNewOrder = False
                End If
                
                bExpire = bExpire And IsOpenOrder(Order.Status)
                
                ' Only process the order if the status date in the record is greater than or equal to the
                ' status date of the order OR if the order is currently pending...
                ' 06/23/2008 DAJ: OR, if the order is not marked as a snapshot, then we must have moved
                ' it to history at some point and now it has come back to being a snapshot order...
                ' 10/15/2009 DAJ: OR this message is expiring an open order...
                ' 06/25/2012 DAJ: OR this is an unsolicited cancel.  We are running into cases where expired
                ' orders are coming back with an unsolicited cancel, but the status date is set to the
                ' original order date...
                If (dFillTime >= Order.StatusDate) Or (OrderIsPending(Order) = True) Or (Len(astrOrder(2)) > 0) Or (Order.IsSnapshot = False) Or (bExpire = True) Or (bUnsolicitedCancel = True) Then
                    ' If we have a rejection or error message, strip off the unnecessary strings that
                    ' correspond to the BEGIN...END block of the message (10/29/2007 DAJ)...
                    If Len(astrOrder(25)) > 0 Then
                        astrOrder(25) = Replace(astrOrder(25), "@Reject", "")
                        astrOrder(25) = Replace(astrOrder(25), "!END", "")
                    End If
                    
                    With Order
                        nPrevStatus = .Status
                        
                        If .OrderID = 0 Then .GenesisOrderID = astrOrder(0)
                        .BrokerID = astrOrder(1)
                        .AccountID = g.Broker.AccountIDForNumber(strAccountNumber)
                        
                        .Buy = (astrOrder(7) = "1") Or (UCase(astrOrder(7)) = "BUY")
                        If CLng(Val(astrOrder(8))) <> 0& Then
                            .Quantity = CLng(Val(astrOrder(8)))
                        End If
                        .SymbolOrSymbolID = strGenesisSymbol
                        
                        nOrderType = GenesisOrderType(astrOrder(6))
                        If Not ((nOrderType = eTT_OrderType_Market) And (.OrderType = eTT_OrderType_MIT)) Then
                            .OrderType = nOrderType
                            .LimitPrice = GenesisPrice(astrOrder(9), dMult)
                            .StopPrice = GenesisPrice(astrOrder(10), dMult)
                        End If
                        If .OrderDate = 0# Then .OrderDate = dFillTime
                        .Message = astrOrder(25)
                        If Len(astrOrder(28)) > 0 Then
                            .BrokerCancelOrderID = m.BrokerInfo.GetDatabaseIdForBrokerId(astrOrder(28))
                            DumpDebug vbTab & "Broker Cancel Order ID Updated: " & astrOrder(28) & " --> " & Str(.BrokerCancelOrderID)
                        End If
                        
                        .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.OrderDate, m.nBroker, .Symbol, False))
                        .IsSnapshot = True
                        
                        If (astrOrder(15) = "0") Or (UCase(astrOrder(15)) = "DAY") Then
                            .Expiration = -1&
                        Else
                            .Expiration = 0&
                        End If
                        
                        ' If there is no order ID yet, then we need to save to that we have an order ID
                        ' to send to the order ID changed function (if in an amend situation)...
                        If .OrderID = 0& Then
                            .Save
                        End If
                        
                        strPrevBrokerID = ""
                        If lOldOrderID <> 0& Then
                            Set OldOrder = New cPtOrder
                            If OldOrder.Load(lOldOrderID) Then
                                ' Copy information from the old order to the new order...
                                strPrevBrokerID = OldOrder.BrokerID
                                .AutoTradeItemID = OldOrder.AutoTradeItemID
                                .BrokerCancelOrderID = OldOrder.BrokerCancelOrderID
                                .CancelOrderID = OldOrder.CancelOrderID
                                .TriggerOrderID = OldOrder.TriggerOrderID
                                .TriggerOptions = OldOrder.TriggerOptions
                                .ConditionOptions = OldOrder.ConditionOptions
                                .TrailAmount = OldOrder.TrailAmount
                                .TrailOptions = OldOrder.TrailOptions
                                .ExitPos = OldOrder.ExitPos
                                .IsAutoExit = OldOrder.IsAutoExit
                                .GroupID = OldOrder.GroupID
                                .GroupName = OldOrder.GroupName
                            
                                ' Clean out information for the old order and refresh...
                                OldOrder.GenesisOrderID = ""
                                OldOrder.BrokerCancelOrderID = 0&
                                OldOrder.CancelOrderID = 0&
                                OldOrder.Save
                                
                                .PreviousOrder = OldOrder
                                
                                m.BrokerInfo.AddOrder OldOrder, bRefresh
                                OrderCallback OldOrder
                            End If
                            
                            DumpDebug "Order ID for " & .GenesisOrderID & " changed from " & Str(lOldOrderID) & " to " & Str(.OrderID)
                            OrderIDChanged lOldOrderID, .OrderID
                            SendOrderIdChangeToOptNav .AccountID, Str(lOldOrderID), Str(.OrderID)
                        End If
                        
                        If Len(astrOrder(2)) > 0 Then
                            bFillReceived = True
                            bFillExisted = False
                            bFillChanged = True
                            
                            ' PFG's Fill ID is usually FILLTN followed by a seven digit number that wraps around,
                            ' so since a user could get the same fill ID for multiple orders over time, we will
                            ' prepend the order ID (which is now Date-Account-OrderID) to keep it unique...
                            ' Example: 20080825:A12345-123456-FILLTN0000001
                            strFillID = Order.BrokerID & "-" & astrOrder(2)
                            
                            dFillPrice = GenesisPrice(astrOrder(21), dMult)
                            lFillQty = CLng(Val(astrOrder(18)))
                            
                            For lIndex = 1 To .Fills.Count
                                If .Fills(lIndex).BrokerID = strFillID Then
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
                                .BrokerID = strFillID
                                .BrokerOrderID = Order.BrokerID
                                .Buy = Order.Buy
                                If .FillDate = 0# Then
                                    .FillDate = dFillTime
                                End If
                                .IsManual = False
                                .IsSnapshot = True
                                .OrderID = Order.OrderID
                                .Price = dFillPrice
                                .Quantity = lFillQty
                                .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.FillDate, m.nBroker, Order.Symbol, False))
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
                        
                        .Status = GenesisOrderStatus(astrOrder(24), astrOrder(23), astrOrder(25))
                        If dFillTime > .StatusDate Then
                            .StatusDate = dFillTime
                        End If
                        If (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_ParkPending) Then
                            .Status = eTT_OrderStatus_Parked
                        ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_Parked) And (bNewOrder = False) Then
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
                        ElseIf (IsForex(strGenesisSymbol) = True) And (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_AmendPending) Then
                            ' For Forex symbols on an amend, we get a Cancelled status message followed by a
                            ' Replaced order status.  This causes the auto exits not to work correctly.  So,
                            ' for forex only, if the order status was Amend Pending and we are now getting a
                            ' Cancelled status, change it to Amended instead...
                            .Status = eTT_OrderStatus_Amended
                        ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_AmendPending) Then
                            ' All of a sudden for futures, we are getting a Cancelled message when attempting
                            ' an amend which in turn screws up the auto exits.  At least for now, if the order
                            ' was amend pending and is now being called Cancelled, lets call it Amended instead
                            ' so that the auto exits will work properly (DAJ 04/24/2008)...
                            .Status = eTT_OrderStatus_Amended
                        End If
                        
                        .Save
                        
                        If bFillReceived Then
                            m.BrokerInfo.AddOrder Order, bRefresh
                            If Len(strPrevBrokerID) > 0 Then
                                .PreviousBrokerID = strPrevBrokerID
                                m.BrokerInfo.AddOrder Order, bRefresh
                            End If
                            
                            m.BrokerInfo.AddFill Fill, bRefresh
                        
                            FillCallback Order, Fill, bFillExisted, True, bFillChanged
                            If UCase(astrOrder(23)) <> "REWIND" Then
                                g.ActivityLogs.AddFillToActivityLog m.nBroker, Fill, False
                                m.dLastFillTime = gdTickCount
                            End If
                            
                            SendFillToOptionNav Fill, bRefresh
                        Else
                            m.BrokerInfo.AddOrder Order, bRefresh
                            If Len(strPrevBrokerID) > 0 Then
                                .PreviousBrokerID = strPrevBrokerID
                                m.BrokerInfo.AddOrder Order, bRefresh
                            End If
                        
                            OrderCallback Order
                            If UCase(astrOrder(23)) <> "REWIND" Then
                                g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                            End If
                            
                            SendOrderToOptionNav Order, bRefresh
                        End If
                        
                        If (Len(astrOrder(25)) > 0) And (UCase(astrOrder(23)) <> "REWIND") Then
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox Order.OrderText & " REJECTED||" & astrOrder(25), m.strBrokerName & " Order " & Order.BrokerID & " Rejected", vbLeftJustify
                        End If
                        
                        If (.Status = eTT_OrderStatus_OverFilled) And (UCase(astrOrder(23)) <> "REWIND") Then
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox "Trade Navigator has received duplicate fill|information for order " & Order.BrokerID & ".||This will cause this order to appear over-filled and will likely result in a position mismatch.||PLEASE CALL YOUR BROKER AND VERIFY YOUR POSITIONS.", "Over Filled Order on " & m.strBrokerName, vbCenter
                        End If
                        
                        DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                    End With
                Else
                    DumpDebug "Order not processed because record date (" & DateFormat(dFillTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ") is older than order status date (" & DateFormat(Order.StatusDate, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ")"
                    m.BrokerInfo.AddOrder Order, bRefresh
                End If ' (dFillTime >= Order.StatusDate) Or (OrderIsPending(Order) = True) Or (Len(astrOrder(2)) > 0)
            Else
                DumpDebug "Order not processed because symbol could not be converted or record date was missing"
            End If ' (Len(strGenesisSymbol) > 0) And (Len(astrOrder(22)) > 0)
        End If ' bDupFill
        
        If bRefresh = False Then
            SendNowToOptionNav
        End If
    End If ' Len(strOrderInfo) > 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.HandleOrderInfo"
    
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
''              Carried Avg Entry, Instrument Type, ID, Carried Info
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandlePositionInfo(ByVal strPositionInfo As String)
On Error GoTo ErrSection:

    Dim astrPosition As New cGdArray    ' Array of position information
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim lPos As Long                    ' Position in the array
    
    If Len(strPositionInfo) > 0 Then
        astrPosition.SplitFields strPositionInfo, vbTab
        
        strGenesisSymbol = GenesisSymbol(astrPosition(2), , , True)
        If Len(strGenesisSymbol) > 0 Then
            If m.astrBrokerPos.BinarySearch(astrPosition(0) & vbTab & strGenesisSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                m.astrBrokerPos(lPos) = astrPosition(0) & vbTab & strGenesisSymbol & vbTab & astrPosition(3) & vbTab & astrPosition(7) & vbTab & astrPosition(13)
            Else
                m.astrBrokerPos.Add astrPosition(0) & vbTab & strGenesisSymbol & vbTab & astrPosition(3) & vbTab & astrPosition(7) & vbTab & astrPosition(13), lPos
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.HandlePositionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SynchronizePositions
'' Description: Synchronize the positions between Genesis and PFG
'' Inputs:      None
'' Returns:     None
''
'' Position:    Account, Symbol, Position, Average Entry, Overnight Position
'' FillSummary: Account, Symbol, At ID, Buys, Sells, Net, Total, PriceSum, Entries,
''              ClosedProfit, AvgEntry, Initial Fill Price, Initial Fill Date,
''              Session Date, Last Traded, Overnight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SynchronizePositions()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPosition As Long               ' Position from the broker
    Dim strPosition As String           ' Position information
    Dim FillSumms As cAccountPositions  ' Collection of fill summaries
    Dim strAccount As String            ' Account number
    Dim lBrokerPos As Long              ' Broker position
    Dim lBrokerCarried As Long          ' Broker carried position
    Dim lPos As Long                    ' Position in an array
    Static strFirstTime As String       ' String of first time mismatches
    
    ' Set the variable to notify that we are currently synchronizing positions...
    m.bSyncInProgress = True
    
    frmPositionConfirm.ShowMe m.nBroker, "", "", False, False
    
    SendBeginToOptionNav eGDOptNav_Position, m.nBroker
    Set FillSumms = m.BrokerInfo.FillSummaries(True)
    If Not FillSumms Is Nothing Then
        For lIndex = 1 To FillSumms.Count
            With FillSumms(lIndex)
                If .AutoTradeItemID = -1& Then
                    g.OrderStrategies.RefreshPosition .AccountID, .SymbolOrSymbolID
                    
                    strAccount = g.Broker.AccountNumberForID(.AccountID)
                    
                    ' Retrieve the broker position and broker carried position from the array...
                    lBrokerPos = CurrentPosition(strAccount, .Symbol)
                    lBrokerCarried = CarriedPosition(strAccount, .Symbol)
                    lPosition = lBrokerCarried + .NumBuysSnapshot - .NumSellsSnapshot
                    
                    ' Verify first that broker information is consistent with itself...
                    DumpDebug "Position Verification(" & strAccount & ", " & .Symbol & "): " & Str(lBrokerCarried) & " + " & Str(.NumBuysSnapshot) & " - " & Str(.NumSellsSnapshot) & " = " & Str(lBrokerPos)
                    If lPosition <> lBrokerPos Then
                        m.BrokerInfo.ConsistentBroker(strAccount, .Symbol) = False
                        
                        If InStr(strFirstTime, "," & strAccount & ";" & .Symbol & ",") <> 0 Then
                            DumpDebug "Position Mismatch (" & strAccount & ", " & .Symbol & "): Second Time"
                            strFirstTime = Replace(strFirstTime, "," & strAccount & ";" & .Symbol & ",", "")
                        
                            HandleSecondTimePositionMismatch m.strBrokerName, lBrokerPos, lPosition, FillSumms(lIndex)
                        Else
                            DumpDebug "Position Mismatch (" & strAccount & ", " & .Symbol & "): First Time"
                            strFirstTime = strFirstTime & "," & strAccount & ";" & .Symbol & ","
                        End If
                    Else
                        m.BrokerInfo.ConsistentBroker(strAccount, .Symbol) = True
                        
                        If InStr(strFirstTime, "," & strAccount & ";" & .Symbol & ",") <> 0 Then
                            strFirstTime = Replace(strFirstTime, "," & strAccount & ";" & .Symbol & ",", "")
                        End If
                        
                        ' Now verify that the carried position information matches...
                        If lBrokerPos <> .CurrentPositionSnapshot Then
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
                End If
            End With
        Next lIndex
    End If
    
    SendEndToOptionNav eGDOptNav_Position, m.nBroker, True
    
    If Len(strFirstTime) > 0 Then
        Refresh
    ElseIf m.astrToFix.Size > 0 Then
        FixPositions
    End If
    
    ' Set the variable to notify that we are no longer synchronizing positions...
    m.bSyncInProgress = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.SynchronizePositions"
    
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
    RaiseError "cPFG.LoadOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleAccountInfo
'' Description: Handle account information back from the PFG server
'' Inputs:      Account Info
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
Private Sub HandleAccountInfo(ByVal strAccountInfo As String)
On Error GoTo ErrSection:

    
    Dim lIndex As Long                  ' Index into a for loop
    Dim AccountInfo As cPfgAccountInfo  ' Account information object
    
    Set AccountInfo = New cPfgAccountInfo
    AccountInfo.FromString strAccountInfo
    
    For lIndex = 0 To Forms.Count - 1
        If TypeOf Forms(lIndex) Is frmAccountInfo Then
            If Forms(lIndex).Account = AccountInfo.AccountNumber Then
                Forms(lIndex).PfgAccountInfo = AccountInfo
                Exit For
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.HandleAccountInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleLinkedOrders
'' Description: Handle linked order information back from the PFG server
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      Order ID 1, Order ID 2, Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleLinkedOrders(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strOrderOne As String           ' Order Number for the first order
    Dim strOrderTwo As String           ' Order Number for the second order
    Dim strStatus As String             ' Status of the order link
    Dim Order1 As cPtOrder              ' Order object for the first order
    Dim Order2 As cPtOrder              ' Order object for the second order
    Dim strAccountNumber As String      ' Account number
    Dim strAltOrderID1 As String        ' Alternate broker ID
    Dim strAltOrderID2 As String        ' Alternate broker ID
    Dim bFailed As Boolean              ' Did the link fail?
    
    strOrderOne = Parse(strMessage, vbTab, 1)
    strOrderTwo = Parse(strMessage, vbTab, 2)
    strStatus = Parse(strMessage, vbTab, 3)
    bFailed = (UCase(strStatus) = "FAILED")
    
    If (Len(strOrderOne) > 0) And (Len(strOrderTwo) > 0) Then
        strAccountNumber = Parse(strOrderOne, "-", 1)
        strAltOrderID1 = m.BrokerInfo.GetAltBrokerID(strOrderOne)
        strAltOrderID2 = m.BrokerInfo.GetAltBrokerID(strOrderTwo)
        
        If (Len(strAltOrderID1) > 0) And (Len(strAccountNumber) > 0) Then
            Set Order1 = LoadOrder(strAccountNumber, "", strAltOrderID1)
            If Not Order1 Is Nothing Then
                strAccountNumber = Parse(strOrderTwo, "-", 1)
        
                If (Len(strAltOrderID2) > 0) And (Len(strAccountNumber) > 0) Then
                    Set Order2 = LoadOrder(strAccountNumber, "", strAltOrderID2)
                    If Not Order2 Is Nothing Then
                        If bFailed Then
                            Order1.BrokerCancelOrderID = 0
                        Else
                            Order1.BrokerCancelOrderID = Order2.OrderID
                        End If
                        
                        Order1.Save
                        OrderCallback Order1
        
                        g.OrderLinks.LinkConfirm strAltOrderID1, strAltOrderID2, bFailed
                        
                        If bFailed Then
                            InfBox "The " & m.strBrokerName & " OCO link between '" & strOrderOne & "' and '" & strOrderTwo & "' failed", "!", , m.strBrokerName & " OCO Link Failed"
                        End If
                    End If
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.HandleLinkedOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleUnlinkedOrders
'' Description: Handle unlinked order information back from the PFG server
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      Order ID 1, Order ID 2
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleUnlinkedOrders(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strOrderOne As String           ' Order Number for the first order
    Dim strOrderTwo As String           ' Order Number for the second order
    Dim Order1 As cPtOrder              ' Order object for the first order
    Dim strAccountNumber As String      ' Account number
    Dim strAltOrderID1 As String         ' Alternate broker ID
    Dim strAltOrderID2 As String         ' Alternate broker ID
    
    strOrderOne = Parse(strMessage, vbTab, 1)
    strOrderTwo = Parse(strMessage, vbTab, 2)
    
    If (Len(strOrderOne) > 0) And (Len(strOrderTwo) > 0) Then
        strAccountNumber = Parse(strOrderOne, "-", 1)
        strAltOrderID1 = m.BrokerInfo.GetAltBrokerID(strOrderOne)
        strAltOrderID2 = m.BrokerInfo.GetAltBrokerID(strOrderTwo)
        
        If (Len(strAltOrderID1) > 0) And (Len(strAccountNumber) > 0) Then
            Set Order1 = LoadOrder(strAccountNumber, "", strAltOrderID1)
            If Not Order1 Is Nothing Then
                Order1.BrokerCancelOrderID = 0&
                Order1.Save
                OrderCallback Order1
                
                g.OrderLinks.UnlinkConfirm strAltOrderID1, strAltOrderID2
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.HandleUnlinkedOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PfgSymbol
'' Description: Convert a Genesis symbol to a PFG symbol
'' Inputs:      Genesis symbol, Multiplier, Format
'' Returns:     PFG Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PfgSymbol(ByVal strGenesisSymbol As String, Optional dMult As Double, Optional strFormat As String, Optional strBaseSymbol As String) As String
On Error GoTo ErrSection:

    Dim strBase As String               ' Base symbol
    Dim strContract As String           ' Contract
    Dim strReturn As String             ' Return from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim strYear As String               ' Year of the contract
    Dim lMonth As Long                  ' Month of the contract
    Dim bPut As Boolean                 ' Is this a put or a call?
    Dim dStrike As Double               ' Strike price
    Dim dStrikeMult As Double           ' Strike multiplier
    Dim strStrike As String             ' String version of strike
    Dim strFuture As String             ' Future symbol
    Dim OptSym As New cOptionSymbol     ' Option symbol object

    strReturn = ""
    If Len(strGenesisSymbol) > 0 Then
        Select Case UCase(SecurityType(strGenesisSymbol, True))
            Case "I"
                If IsForex(strGenesisSymbol) Then
                    For lIndex = 0 To m.astrSymbols.Size - 1
                        If Parse(m.astrSymbols(lIndex), vbTab, 2) = strGenesisSymbol Then
                            strReturn = Parse(m.astrSymbols(lIndex), vbTab, 1)
                            dMult = Val(Parse(m.astrSymbols(lIndex), vbTab, 3))
                            strFormat = Parse(m.astrSymbols(lIndex), vbTab, 4)
                            
                            Exit For
                        End If
                    Next lIndex
                End If
            
            Case "FO"
                OptSym.FromGenesis strGenesisSymbol
                If Len(OptSym.BaseSymbol) > 0 Then
                    For lIndex = 0 To m.astrSymbols.Size - 1
                        If Parse(m.astrSymbols(lIndex), vbTab, 2) = "O:" & OptSym.BaseSymbol Then
                            dMult = Val(Parse(m.astrSymbols(lIndex), vbTab, 3))
                            strFormat = Parse(m.astrSymbols(lIndex), vbTab, 4)
                            dStrikeMult = Parse(m.astrSymbols(lIndex), vbTab, 8)
                            strBaseSymbol = Mid(Parse(m.astrSymbols(lIndex), vbTab, 1), 3)
                            strReturn = OptSym.ToPfg(strBaseSymbol, dStrikeMult)
                            
                            Exit For
                        End If
                    Next lIndex
                End If
            
            Case "F"
                strBase = Parse(strGenesisSymbol, "-", 1)
                strContract = Parse(strGenesisSymbol, "-", 2)
                
                For lIndex = 0 To m.astrSymbols.Size - 1
                    If Parse(m.astrSymbols(lIndex), vbTab, 2) = strBase Then
                        strYear = Left(strContract, 4)
                        lMonth = CLng(Val(Right(strContract, 2)))
                        
                        strBaseSymbol = Parse(m.astrSymbols(lIndex), vbTab, 1)
                        strReturn = strBaseSymbol & MonthToCode(lMonth) & Right(strYear, 1)
                        dMult = Val(Parse(m.astrSymbols(lIndex), vbTab, 3))
                        strFormat = Parse(m.astrSymbols(lIndex), vbTab, 4)
                        
                        Exit For
                    End If
                Next lIndex
                
            Case "S"
                If InStr(strGenesisSymbol, "_") = 0 Then
                    For lIndex = 0 To m.astrSymbols.Size - 1
                        If Parse(m.astrSymbols(lIndex), vbTab, 2) = "!" Then
                            strReturn = "EQ-" & strGenesisSymbol
                            dMult = Val(Parse(m.astrSymbols(lIndex), vbTab, 3))
                            strFormat = Parse(m.astrSymbols(lIndex), vbTab, 4)
                            
                            Exit For
                        End If
                    Next lIndex
                End If
                
                
            Case "SO"
                For lIndex = 0 To m.astrSymbols.Size - 1
                    If Parse(m.astrSymbols(lIndex), vbTab, 2) = "@" Then
                        OptSym.FromGenesis strGenesisSymbol
                        
                        strReturn = OptSym.ToPfg
                        dMult = Val(Parse(m.astrSymbols(lIndex), vbTab, 3))
                        strFormat = Parse(m.astrSymbols(lIndex), vbTab, 4)
                        
                        Exit For
                    End If
                Next lIndex
        End Select
    End If

    PfgSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.PfgSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Convert a PFG symbol to a Genesis symbol
'' Inputs:      PFG symbol, Multiplier, Format, For Position?
'' Returns:     Genesis Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbol(ByVal strPfgSymbol As String, Optional dMult As Double, Optional strFormat As String, Optional ByVal bPosition As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strBase As String               ' Base symbol
    Dim lPos As Long                    ' Position of base symbol in the array
    Dim strReturn As String             ' Return from the function
    Dim lYear As Long                   ' Year of the contract
    Dim lMonth As Long                  ' Month of the contract
    Dim bPut As Boolean                 ' Is this a put or a call?
    Dim dStrike As Double               ' Strike price
    Dim dStrikeMult As Double           ' Strike multiplier
    Dim strStrike As String             ' String version of strike
    Dim OptSym As New cOptionSymbol     ' Option symbol object
            
    strReturn = ""
    If Len(strPfgSymbol) > 0 Then
        ' Forex symbol...
        If InStr(strPfgSymbol, "/") <> 0 Then
            If FindGenesisBaseSymbol(strPfgSymbol, bPosition, lPos) Then
                strReturn = Parse(m.astrSymbols(lPos), vbTab, 2)
                dMult = Val(Parse(m.astrSymbols(lPos), vbTab, 3))
                strFormat = Parse(m.astrSymbols(lPos), vbTab, 4)
            End If
        
        ' Stock option symbol...
        ElseIf InStr(strPfgSymbol, "+") Then
            If FindGenesisBaseSymbol("@", bPosition, lPos) Then
                OptSym.FromPfg strPfgSymbol
                strReturn = OptSym.ToGenesis
                dMult = Val(Parse(m.astrSymbols(lPos), vbTab, 3))
                strFormat = Parse(m.astrSymbols(lPos), vbTab, 4)
            End If
            
        ' Future Option symbol...
        ElseIf UCase(Right(strPfgSymbol, 1)) = "O" Then
            OptSym.FromPfg strPfgSymbol
            If Len(OptSym.BrokerBase) > 0 Then
                If FindGenesisBaseSymbol("O:" & OptSym.BrokerBase, bPosition, lPos) Then
                    OptSym.ConvertBrokerInfo Parse(m.astrSymbols(lPos), vbTab, 2), Val(Parse(m.astrSymbols(lPos), vbTab, 8))
                    strReturn = OptSym.ToGenesis
                    dMult = Val(Parse(m.astrSymbols(lPos), vbTab, 3))
                    strFormat = Parse(m.astrSymbols(lPos), vbTab, 4)
                End If
            End If

        ' Future symbol...
        ElseIf IsDigit(Right(strPfgSymbol, 1)) Then
            If Len(strPfgSymbol) > 2 Then
                strBase = Left(strPfgSymbol, Len(strPfgSymbol) - 2)
                
                If FindGenesisBaseSymbol(strBase, bPosition, lPos) Then
                    lYear = CLng(Val(Right(strPfgSymbol, 1))) + 2000
                    lMonth = CodeToMonth(Mid(strPfgSymbol, Len(strPfgSymbol) - 1, 1))
                    Do While (lYear * 100 + lMonth) < (Year(Date) * 100 + Month(Date))
                        lYear = lYear + 10
                    Loop
                
                    strReturn = Parse(m.astrSymbols(lPos), vbTab, 2) & "-" & Str(lYear) & Format(lMonth, "00")
                    dMult = Val(Parse(m.astrSymbols(lPos), vbTab, 3))
                    strFormat = Parse(m.astrSymbols(lPos), vbTab, 4)
                End If
            End If
            
        ' Stock symbol...
        Else
            If FindGenesisBaseSymbol("!", bPosition, lPos) Then
                strReturn = Replace(strPfgSymbol, "EQ-", "")
                dMult = Val(Parse(m.astrSymbols(lPos), vbTab, 3))
                strFormat = Parse(m.astrSymbols(lPos), vbTab, 4)
            End If
            
        End If
    End If

    GenesisSymbol = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.GenesisSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FindGenesisBaseSymbol
'' Description: Find the Genesis Base symbol for the PFG one passed in
'' Inputs:      PFG Base Symbol, For Position?, Position, Genesis Base
'' Returns:     True if Found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindGenesisBaseSymbol(ByVal strPfgBaseSymbol As String, Optional ByVal bPosition As Boolean = False, Optional lPos As Long, Optional strGenesisBase As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Did we find the symbol?
    
    bFound = False
    For lIndex = 0 To m.astrSymbols.Size - 1
        If (bPosition = True) And (Parse(m.astrSymbols(lIndex), vbTab, 9) = strPfgBaseSymbol) Then
            bFound = True
        ElseIf Parse(m.astrSymbols(lIndex), vbTab, 1) = strPfgBaseSymbol Then
            bFound = True
        End If
        
        If bFound = True Then
            lPos = lIndex
            strGenesisBase = Parse(m.astrSymbols(lIndex), vbTab, 2)
            Exit For
        End If
    Next lIndex
    
    FindGenesisBaseSymbol = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.FindGenesisBaseSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PfgOrderType
'' Description: Convert a Genesis order type to a PFG order type
'' Inputs:      Genesis order type
'' Returns:     PFG Order Type (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PfgOrderType(ByVal nOrderType As eTT_OrderType) As String
On Error GoTo ErrSection:

    Select Case nOrderType
        Case eTT_OrderType_Market
            PfgOrderType = "1"
        Case eTT_OrderType_Stop
            PfgOrderType = "3"
        Case eTT_OrderType_Limit
            PfgOrderType = "2"
        Case eTT_OrderType_StopWithLimit
            PfgOrderType = "4"
        Case Else
            PfgOrderType = ""
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.PfgOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderType
'' Description: Convert a PFG order type to a Genesis order type
'' Inputs:      PFG order type
'' Returns:     Genesis Order Type (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderType(ByVal strPfgOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Select Case UCase(strPfgOrderType)
        Case "1", "MARKET"
            GenesisOrderType = eTT_OrderType_Market
        Case "2", "LIMIT"
            GenesisOrderType = eTT_OrderType_Limit
        Case "3", "STOP"
            GenesisOrderType = eTT_OrderType_Stop
        Case "4", "STOPLIMIT", "STOP / LIMIT"
            GenesisOrderType = eTT_OrderType_StopWithLimit
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.GenesisOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PfgPrice
'' Description: Convert a Genesis price to a PFG price
'' Inputs:      Genesis Price, Multiplier, Format
'' Returns:     PFG Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PfgPrice(ByVal dGenesisPrice As Double, ByVal dMult As Double, ByVal strFormat As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from the function
    Dim dValue As Double                ' Value of converted price
    
    If dMult <> 0 Then
        Select Case dMult
            Case 8
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 8, "0")
            
            Case 32
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 32, "00")
            
            Case 32.25
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 3200, "0000")
                strReturn = Left(strReturn, Len(strReturn) - 1)
            
            Case 32.5
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 320, "000")
                
            Case 64
                dValue = CDbl(Int(dGenesisPrice)) + (((dGenesisPrice - Int(dGenesisPrice)) * 64) / 100)
                strReturn = Replace(Format(dValue, strFormat), ",", ".")
            
            Case 64.5
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 640, "000")
                
            Case Else
                strReturn = Replace(Format(dGenesisPrice / dMult, strFormat), ",", ".")

        End Select
    End If
    
    PfgPrice = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.PfgPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisPrice
'' Description: Convert a PFG price to a Genesis price
'' Inputs:      PFG Price, Multiplier
'' Returns:     Genesis Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisPrice(ByVal strPfgPrice As String, ByVal dMult As Double) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value from the function
    Dim strWhole As String              ' Whole portion of the price
    Dim strFraction As String           ' Fractional portion of the price
    
    If (Left(strPfgPrice, 3) = "FIX") Or (strPfgPrice = "0") Then
        dReturn = 0#
    ElseIf (dMult <> 0) And (Len(strPfgPrice) > 0) Then
        Select Case dMult
            ' In the case of 8ths, PFG will send a price with the number of eighths after
            ' the decimal point, so 100.2=100 2/8, 100.4=100 4/8, etc.
            Case 8
                strWhole = Parse(strPfgPrice, ".", 1)
                strFraction = Parse(strPfgPrice, ".", 2)
                
                If Len(strFraction) = 0 Then
                    strFraction = "0"
                End If
                
                dReturn = Val(strWhole) + (Val(strFraction) / 8)
            
            ' In the case of 32nds, PFG will send a price with the number of 32nds after
            ' the decimal point, so 100.20=100 20/32, 100.16=100 16/32, etc.
            Case 32
                strWhole = Parse(strPfgPrice, ".", 1)
                strFraction = Parse(strPfgPrice, ".", 2)
                
                If Len(strFraction) = 0 Then
                    strFraction = "00"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "0"
                End If
                
                dReturn = Val(strWhole) + (Val(strFraction) / 32)
            
            ' In the case of quarter 32nds, PFG will send a price with the number of 32nds plus the
            ' fractional part of 32nds after the decimal point (with an implied 5 on the end if the
            ' PFG price ends in a 2 or a 7), so 100.200=100 20.0/32, 100.162=100 16.25/32,
            ' 100.315=100 31.5/32, 100.307=100 30.75/32, etc.
            Case 32.25
                strWhole = Parse(strPfgPrice, ".", 1)
                strFraction = Parse(strPfgPrice, ".", 2)
                
                If Len(strFraction) = 0 Then
                    strFraction = "000"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "00"
                ElseIf Len(strFraction) = 2 Then
                    strFraction = strFraction & "0"
                End If
                
                If Right(strFraction, 1) = "2" Or Right(strFraction, 1) = "7" Then
                    strFraction = strFraction & "5"
                Else
                    strFraction = strFraction & "0"
                End If
                
                dReturn = Val(strWhole) + ((Val(strFraction) / 100) / 32)
            
            ' In the case of half 32nds, PFG will send a price with the number of 32nds plus the
            ' fractional part of 32nds after the decimal point, so 100.200=100 20.0/32,
            ' 100.165=100 16.5/32, etc.
            Case 32.5
                strWhole = Parse(strPfgPrice, ".", 1)
                strFraction = Parse(strPfgPrice, ".", 2)
                
                If Len(strFraction) = 0 Then
                    strFraction = "000"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "00"
                ElseIf Len(strFraction) = 2 Then
                    strFraction = strFraction & "0"
                End If
                
                dReturn = Val(strWhole) + ((Val(strFraction) / 10) / 32)
            
            ' In the case of 64ths, PFG will send a price with the number of 64ths after
            ' the decimal point, so 100.20=100 20/64, 100.16=100 16/64, etc.
            Case 64
                strWhole = Parse(strPfgPrice, ".", 1)
                strFraction = Parse(strPfgPrice, ".", 2)
                
                ' If we don't have a fractional part, default to zero...
                If Len(strFraction) = 0 Then
                    strFraction = "00"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "0"
                
                ' Otherwise, if we get three digits, put a decimal point between the
                ' second and third digits...
                ElseIf Len(strFraction) = 3 Then
                    strFraction = Left(strFraction, 2) & "." & Right(strFraction, 1)
                End If
                
                dReturn = Val(strWhole) + (Val(strFraction) / 64)
            
            ' In the case of half 64ths, PFG will send a price with the number of 64ths plus the
            ' fractional part of 64ths after the decimal point, so 100.200=100 20.0/64,
            ' 100.165=100 16.5/64, etc.
            Case 64.5
                strWhole = Parse(strPfgPrice, ".", 1)
                strFraction = Parse(strPfgPrice, ".", 2)
                
                If Len(strFraction) = 0 Then
                    strFraction = "000"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "00"
                ElseIf Len(strFraction) = 2 Then
                    strFraction = strFraction & "0"
                End If
                
                dReturn = Val(strWhole) + ((Val(strFraction) / 10) / 64)
            
            Case Else
                dReturn = Val(strPfgPrice) * dMult
                
        End Select
    End If
    
    GenesisPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.GenesisPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderStatus
'' Description: Convert a PFG order status to a Genesis order status
'' Inputs:      PFG order status, Event for Callback, Message
'' Returns:     Genesis order status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderStatus(ByVal strPfgOrderStatus As String, ByVal strPfgEvent As String, Optional ByVal strMessage As String = "") As eTT_OrderStatus
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderStatus      ' Return value from the function

    Select Case UCase(strPfgOrderStatus)
        Case "QUEUED"
            nReturn = eTT_OrderStatus_Queued
        Case "PENDING"
            nReturn = eTT_OrderStatus_Working
        Case "CANCELED"
            If UCase(strPfgEvent) = "REPLACE" Then
                nReturn = eTT_OrderStatus_Amended
            Else
                nReturn = eTT_OrderStatus_Cancelled
            End If
        Case "FILLED", "UNKNOWN 'X'", "EXECUTED"
            nReturn = eTT_OrderStatus_Filled
        Case "REJECTED"
            If InStr(strMessage, "EXPIRED") <> 0 Then
                nReturn = eTT_OrderStatus_Expired
            Else
                nReturn = eTT_OrderStatus_Rejected
            End If
    End Select
    
    GenesisOrderStatus = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.GenesisOrderStatus"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MonthNum
'' Description: Number of the month for the given abbreviation
'' Inputs:      Month Abbreviation
'' Returns:     Number of the month
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MonthNum(ByVal strMonth As String) As Integer
On Error GoTo ErrSection:

    Select Case UCase(strMonth)
        Case "JAN", "JANUARY"
            MonthNum = 1
        Case "FEB", "FEBRUARY"
            MonthNum = 2
        Case "MAR", "MARCH"
            MonthNum = 3
        Case "APR", "APRIL"
            MonthNum = 4
        Case "MAY"
            MonthNum = 5
        Case "JUN", "JUNE"
            MonthNum = 6
        Case "JUL", "JULY"
            MonthNum = 7
        Case "AUG", "AUGUST"
            MonthNum = 8
        Case "SEP", "SEPTEMBER"
            MonthNum = 9
        Case "OCT", "OCTOBER"
            MonthNum = 10
        Case "NOV", "NOVEMBER"
            MonthNum = 11
        Case "DEC", "DECEMBER"
            MonthNum = 12
        Case Else
            MonthNum = 0
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.MonthNum"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MessageTypeToString
'' Description: Convert a message type enumeration to a string
'' Inputs:      Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MessageTypeToString(ByVal nType As eGDPfgMessageTypes) As String
On Error GoTo ErrSection:

    Select Case nType
        Case eGDPfgMessageType_Connect
            MessageTypeToString = "eGDPfgMessageType_Connect"
        Case eGDPfgMessageType_Disconnect
            MessageTypeToString = "eGDPfgMessageType_Disconnect"
        Case eGDPfgMessageType_AddOrder
            MessageTypeToString = "eGDPfgMessageType_AddOrder"
        Case eGDPfgMessageType_AmendOrder
            MessageTypeToString = "eGDPfgMessageType_AmendOrder"
        Case eGDPfgMessageType_CancelOrder
            MessageTypeToString = "eGDPfgMessageType_CancelOrder"
        Case eGDPfgMessageType_UnloadApp
            MessageTypeToString = "eGDPfgMessageType_UnloadApp"
        Case eGDPfgMessageType_GetAccounts
            MessageTypeToString = "eGDPfgMessageType_GetAccounts"
        Case eGDPfgMessageType_GetOrders
            MessageTypeToString = "eGDPfgMessageType_GetOrders"
        Case eGDPfgMessageType_GetFills
            MessageTypeToString = "eGDPfgMessageType_GetFills"
        Case eGDPfgMessageType_GetPositions
            MessageTypeToString = "eGDPfgMessageType_GetPositions"
        Case eGDPfgMessageType_GetSingleOrder
            MessageTypeToString = "eGDPfgMessageType_GetSingleOrder"
        
        Case eGDPfgMessageType_ConnectionInfo
            MessageTypeToString = "eGDPfgMessageType_ConnectionInfo"
        Case eGDPfgMessageType_AppLoaded
            MessageTypeToString = "eGDPfgMessageType_AppLoaded"
        Case eGDPfgMessageType_AppUnloaded
            MessageTypeToString = "eGDPfgMessageType_AppUnloaded"
        Case eGDPfgMessageType_Heartbeat
            MessageTypeToString = "eGDPfgMessageType_Heartbeat"
        Case eGDPfgMessageType_Order
            MessageTypeToString = "eGDPfgMessageType_Order"
        Case eGDPfgMessageType_AccountR
            MessageTypeToString = "eGDPfgMessageType_AccountR"
        Case eGDPfgMessageType_OrderR
            MessageTypeToString = "eGDPfgMessageType_OrderR"
        Case eGDPfgMessageType_FillR
            MessageTypeToString = "eGDPfgMessageType_FillR"
        Case eGDPfgMessageType_PositionR
            MessageTypeToString = "eGDPfgMessageType_PositionR"
            
        Case eGDPfgMessageType_GetBlocks
            MessageTypeToString = "eGDPfgMessageType_GetBlocks"
        Case eGDPfgMessageType_BlockR
            MessageTypeToString = "eGDPfgMessageType_BlockR"
        Case eGDPfgMessageType_GetAccountInfo
            MessageTypeToString = "eGDPfgMessageType_GetAccountInfo"
        Case eGDPfgMessageType_AccountInfo
            MessageTypeToString = "eGDPfgMessageType_AccountInfo"
        Case eGDPfgMessageType_LinkOrders
            MessageTypeToString = "eGDPfgMessageType_LinkOrders"
        Case eGDPfgMessageType_OrdersLinked
            MessageTypeToString = "eGDPfgMessageType_OrdersLinked"
        Case eGDPfgMessageType_UnlinkOrders
            MessageTypeToString = "eGDPfgMessageType_UnlinkOrders"
        Case eGDPfgMessageType_OrdersUnlinked
            MessageTypeToString = "eGDPfgMessageType_OrdersUnlinked"
        Case eGDPfgMessageType_CreateSpread
            MessageTypeToString = "eGDPfgMessageType_CreateSpread"
        Case eGDPfgMessageType_SpreadCreated
            MessageTypeToString = "eGDPfgMessageType_SpreadCreated"
        Case eGDPfgMessageType_RequestForQuote
            MessageTypeToString = "eGDPfgMessageType_RequestForQuote"
        Case eGDPfgMessageType_SpreadQuote
            MessageTypeToString = "eGDPfgMessageType_SpreadQuote"
        Case eGDPfgMessageType_RequestSpreadCodes
            MessageTypeToString = "eGDPfgMessageType_RequestSpreadCodes"
        Case eGDPfgMessageType_SpreadCodes
            MessageTypeToString = "eGDPfgMessageType_SpreadCodes"
        Case eGDPfgMessageType_RequestSpreadLegs
            MessageTypeToString = "eGDPfgMessageType_RequestSpreadLegs"
        Case eGDPfgMessageType_SpreadLegs
            MessageTypeToString = "eGDPfgMessageType_SpreadLegs"
            
        Case Else
            MessageTypeToString = Str(nType)
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.MessageTypeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StringToMessageType
'' Description: Convert a string to a message type enumeration
'' Inputs:      String
'' Returns:     Message Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StringToMessageType(ByVal strMessageType As String) As eGDPfgMessageTypes
On Error GoTo ErrSection:
    
    Dim nReturn As eGDPfgMessageTypes   ' Return value for the function
    
    Select Case strMessageType
        Case "eGDPfgMessageType_Connect"
            nReturn = eGDPfgMessageType_Connect
        Case "eGDPfgMessageType_Disconnect"
            nReturn = eGDPfgMessageType_Disconnect
        Case "eGDPfgMessageType_AddOrder"
            nReturn = eGDPfgMessageType_AddOrder
        Case "eGDPfgMessageType_AmendOrder"
            nReturn = eGDPfgMessageType_AmendOrder
        Case "eGDPfgMessageType_CancelOrder"
            nReturn = eGDPfgMessageType_CancelOrder
        Case "eGDPfgMessageType_UnloadApp"
            nReturn = eGDPfgMessageType_UnloadApp
        Case "eGDPfgMessageType_GetAccounts"
            nReturn = eGDPfgMessageType_GetAccounts
        Case "eGDPfgMessageType_GetOrders"
            nReturn = eGDPfgMessageType_GetOrders
        Case "eGDPfgMessageType_GetFills"
            nReturn = eGDPfgMessageType_GetFills
        Case "eGDPfgMessageType_GetPositions"
            nReturn = eGDPfgMessageType_GetPositions
        Case "eGDPfgMessageType_GetSingleOrder"
            nReturn = eGDPfgMessageType_GetSingleOrder
        
        Case "eGDPfgMessageType_ConnectionInfo"
            nReturn = eGDPfgMessageType_ConnectionInfo
        Case "eGDPfgMessageType_AppLoaded"
            nReturn = eGDPfgMessageType_AppLoaded
        Case "eGDPfgMessageType_AppUnloaded"
            nReturn = eGDPfgMessageType_AppUnloaded
        Case "eGDPfgMessageType_Heartbeat"
            nReturn = eGDPfgMessageType_Heartbeat
        Case "eGDPfgMessageType_Order"
            nReturn = eGDPfgMessageType_Order
        Case "eGDPfgMessageType_AccountR"
            nReturn = eGDPfgMessageType_AccountR
        Case "eGDPfgMessageType_OrderR"
            nReturn = eGDPfgMessageType_OrderR
        Case "eGDPfgMessageType_FillR"
            nReturn = eGDPfgMessageType_FillR
        Case "eGDPfgMessageType_PositionR"
            nReturn = eGDPfgMessageType_PositionR
        Case "eGDPfgMessageType_ContractR"
            nReturn = eGDPfgMessageType_ContractR
        
        Case "eGDPfgMessageType_GetBlocks"
            nReturn = eGDPfgMessageType_GetBlocks
        Case "eGDPfgMessageType_BlockR"
            nReturn = eGDPfgMessageType_BlockR
        Case "eGDPfgMessageType_GetAccountInfo"
            nReturn = eGDPfgMessageType_GetAccountInfo
        Case "eGDPfgMessageType_AccountInfo"
            nReturn = eGDPfgMessageType_AccountInfo
        Case "eGDPfgMessageType_LinkOrders"
            nReturn = eGDPfgMessageType_LinkOrders
        Case "eGDPfgMessageType_OrdersLinked"
            nReturn = eGDPfgMessageType_OrdersLinked
        Case "eGDPfgMessageType_UnlinkOrders"
            nReturn = eGDPfgMessageType_UnlinkOrders
        Case "eGDPfgMessageType_OrdersUnlinked"
            nReturn = eGDPfgMessageType_OrdersUnlinked
        Case "eGDPfgMessageType_CreateSpread"
            nReturn = eGDPfgMessageType_CreateSpread
        Case "eGDPfgMessageType_SpreadCreated"
            nReturn = eGDPfgMessageType_SpreadCreated
        Case "eGDPfgMessageType_RequestForQuote"
            nReturn = eGDPfgMessageType_RequestForQuote
        Case "eGDPfgMessageType_SpreadQuote"
            nReturn = eGDPfgMessageType_SpreadQuote
        Case "eGDPfgMessageType_RequestSpreadCodes"
            nReturn = eGDPfgMessageType_RequestSpreadCodes
        Case "eGDPfgMessageType_SpreadCodes"
            nReturn = eGDPfgMessageType_SpreadCodes
        Case "eGDPfgMessageType_RequestSpreadLegs"
            nReturn = eGDPfgMessageType_RequestSpreadLegs
        Case "eGDPfgMessageType_SpreadLegs"
            nReturn = eGDPfgMessageType_SpreadLegs
        
        Case Else
            nReturn = 0
    End Select
    
    StringToMessageType = nReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.StringToMessageType"
    
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
    Open AddSlash(m.strLogPath) & "TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ") - " & strMessage
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
    If Len(astrOrder(3)) > 0 Then
        astrToAdd(3) = astrOrder(3)
    Else
        astrToAdd(3) = Parse(astrOrder(1), "-", 1)
    End If
    astrToAdd(4) = GenesisSymbol(astrOrder(5), dMult, , False)
    astrToAdd(5) = Str(GenesisOrderType(astrOrder(6)))
    If (UCase(astrOrder(7)) = "BUY") Or (astrOrder(7) = "1") Then
        astrToAdd(6) = Str(CLng(True))
    Else
        astrToAdd(6) = Str(CLng(False))
    End If
    astrToAdd(7) = astrOrder(8)
    astrToAdd(8) = Str(GenesisPrice(astrOrder(9), dMult))
    astrToAdd(9) = Str(GenesisPrice(astrOrder(10), dMult))
    If astrOrder(15) = "0" Then astrToAdd(10) = "-1" Else astrToAdd(10) = "0"
    astrToAdd(11) = Str(GenesisOrderStatus(astrOrder(24), astrOrder(23), astrOrder(25)))
    astrToAdd(12) = astrOrder(19)
    astrToAdd(13) = Str(AdjustDateTime(DateToDouble(astrOrder(22)), astrToAdd(4)))
    
    m.BrokerInfo.AddOrder astrToAdd.JoinFields(vbTab), bRefresh
    
    If Len(astrOrder(2)) > 0 Then
        AddFillToBrokerInfo strOrderString, bRefresh
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.UpdateBrokerInfoOrder"
    
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
    RaiseError "cPFG.GenesisOrderToBrokerInfo"
    
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
    Dim strPfgSymbol As String          ' Pfg Symbol
    Dim dMult As Double                 ' Multiplier for converting prices
    Dim strFormat As String             ' Price format

    If Len(strOrder) > 0 Then
        Set astrOrder = New cGdArray
        astrOrder.Create eGDARRAY_Strings
        Set astrToCancel = New cGdArray
        astrToCancel.Create eGDARRAY_Strings
    
        astrOrder.SplitFields strOrder, vbTab
        
        strPfgSymbol = PfgSymbol(astrOrder(4), dMult, strFormat)
        If (Len(strPfgSymbol) > 0) And (astrOrder(5) <> Str(eTT_OrderType_Market)) Then
            astrToCancel(0) = astrOrder(1)
            astrToCancel(1) = astrOrder(0)
            astrToCancel(2) = astrOrder(3)
            astrToCancel(3) = Left(strPfgSymbol, Len(strPfgSymbol) - 2)
            astrToCancel(4) = strPfgSymbol
            astrToCancel(5) = PfgOrderType(CLng(Val(astrOrder(5))))
            If astrOrder(6) = "0" Then astrToCancel(6) = "2" Else astrToCancel(6) = "1"
            astrToCancel(7) = astrOrder(7)
            astrToCancel(8) = PfgPrice(Val(astrOrder(8)), dMult, strFormat)
            astrToCancel(9) = PfgPrice(Val(astrOrder(9)), dMult, strFormat)
            astrToCancel(10) = ""
            astrToCancel(11) = ""
            astrToCancel(12) = "CME"
            astrToCancel(13) = ""
            If astrOrder(10) = "-1" Then astrToCancel(14) = "0" Else astrToCancel(14) = "1"
            astrToCancel(15) = ""
            astrToCancel(16) = ""
            
            SendPfgMessage eGDPfgMessageType_CancelOrder, astrToCancel.JoinFields(vbTab)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.CancelBrokerInfoOrder"
    
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
    If DuplicateFill(astrFill(2)) = False Then
        astrToAdd(0) = astrFill(1)
        astrToAdd(1) = astrFill(0)
        ' TLB 12/21/2006: to be unique, our FillID must combine their OrderID and FillID
        astrToAdd(2) = astrFill(1) & "-" & astrFill(2)
        If Len(astrFill(3)) > 0 Then
            astrToAdd(3) = astrFill(3)
        Else
            astrToAdd(3) = Parse(astrFill(1), "-", 1)
        End If
        astrToAdd(4) = GenesisSymbol(astrFill(5), dMult, , False)
        astrToAdd(5) = Str(AdjustDateTime(DateToDouble(astrFill(22)), astrToAdd(4)))
        If (UCase(astrFill(7)) = "BUY") Or (astrFill(7) = "1") Then
            astrToAdd(6) = Str(CLng(True))
        Else
            astrToAdd(6) = Str(CLng(False))
        End If
        astrToAdd(7) = astrFill(18)
        astrToAdd(8) = Str(GenesisPrice(astrFill(21), dMult))
        
        m.BrokerInfo.AddFill astrToAdd.JoinFields(vbTab), bRefresh
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.AddFillToBrokerInfo"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddPositionToBrokerInfo
'' Description: Add the position string to the broker info class
'' Inputs:      Position String, Refresh?
'' Returns:     None
''
'' RecordFrom:  Account, Symbol, Symbol Long, Position, Realized, Unrealized,
''              Avg Entry, Carried Position, Carried Realized, Carried Unrealized,
''              Carried Avg Entry, Instrument Type
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
    astrToAdd(1) = GenesisSymbol(astrPosition(2), dMult, , True)
    astrToAdd(2) = astrPosition(3)
    astrToAdd(3) = GenesisPrice(astrPosition(6), dMult)
    astrToAdd(4) = astrPosition(7)
    
    m.BrokerInfo.AddPosition astrToAdd.JoinFields(vbTab), bRefresh

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.AddPositionToBrokerInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AdjustDateTime
'' Description: Adjust the date time to the calendar date, not the session date
'' Inputs:      Date/Time, Symbol
'' Returns:     Adjusted Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AdjustDateTime(ByVal dDateTime As Double, ByVal strSymbol As String) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value from the function
    Dim dExchTime As Double             ' Date/Time converted to exchange time
    Dim dCrossOver As Double            ' Crossover time
    Dim Bars As New cGdBars             ' Bars structure

    dReturn = dDateTime
    If (Len(strSymbol) > 0) And (dDateTime > 0) Then
        SetBarProperties Bars, strSymbol

        ' Convert the broker time passed in to exchange time...
        dExchTime = ConvertBrokerDate(dDateTime, m.nBroker, strSymbol, False)
        
        ' Determine the time 5 minutes after the ending time of the session in exchange time...
        dCrossOver = Int(dExchTime) + (Bars.Prop(eBARS_DefaultEndTime) + 5#) / 1440#
        
        ' Convert that crossover time into broker time...
        dCrossOver = ConvertToBrokerDate(dCrossOver, m.nBroker, strSymbol, False)
        
        ' Add the date portion of the date/time passed in to the time portion of the crossover
        ' to get the crossover time for the day that was passed in...
        dCrossOver = Int(dDateTime) + (dCrossOver - Int(dCrossOver))
        
        ' If the date/time passed in is past the crossover time, subtract one from the date
        ' because it was given in session date and we need calendar date...
        If dDateTime > dCrossOver Then
            dReturn = dReturn - 1#
        End If
    End If
    
    AdjustDateTime = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.AdjustDateTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DuplicateFill
'' Description: Is the given Fill ID the duplicate Fill ID flag?
'' Inputs:      Fill ID
'' Returns:     True if Duplicate Fill Flag, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DuplicateFill(ByVal strFillID As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function

    bReturn = False
    If Len(strFillID) >= 4 Then
        If Left(strFillID, 4) = "0000" Then
            bReturn = True
        End If
    End If
    
    DuplicateFill = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.DuplicateFill"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateFillAtPreviousClose
'' Description: Create fake fills at previous close
'' Inputs:      Account, Symbol, Fill Summary
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateFillAtPreviousClose(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, Optional ByVal FillSummary As cAccountPosition = Nothing)
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' Fill object
    Dim dPreviousCloseTime As Double    ' Previous close time
    Dim Bars As cGdBars                 ' Bars object
        
    If FillSummary Is Nothing Then
        Set FillSummary = g.Broker.FillSummary(vAccountNumberOrID, vSymbolOrSymbolID, -1&)
    End If
    
    If Not FillSummary Is Nothing Then
        Set Bars = New cGdBars
        SetBarProperties Bars, FillSummary.SymbolOrSymbolID
        
        Set Fill = New cPtFill
        With Fill
            .AccountID = FillSummary.AccountID
            .AutoTradingItemID = 0&
            .SymbolOrSymbolID = FillSummary.SymbolOrSymbolID
            .Quantity = Abs(FillSummary.CurrentPosition)
            .Price = PreviousCloseForSymbol(FillSummary.SymbolOrSymbolID, dPreviousCloseTime)
            .FillDate = ConvertToBrokerDate(dPreviousCloseTime, m.nBroker, FillSummary.Symbol, False)
            .SessionDate = Bars.SessionDateForTradeTime(dPreviousCloseTime)
            .BrokerID = "CarPosFix " & .Symbol & " " & DateFormat(CurrentTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
            .BrokerOrderID = .BrokerID
            .Buy = (FillSummary.CurrentPosition < 0&)
            .IsManual = True
            .IsSnapshot = False
            .Save
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.CreateFillAtPreviousClose"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateFillsForCarriedInfo
'' Description: Create fake fills based on the carried position information
'' Inputs:      Account, Symbol, Fill Summary
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateFillsForCarriedInfo(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, Optional ByVal FillSummary As cAccountPosition = Nothing)
On Error GoTo ErrSection:

    Dim astrInfo As cGdArray            ' Information split out into an array
    Dim astrPos As cGdArray             ' Information split out into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object
    Dim lPos As Long                    ' Position in an array
    Dim strAccount As String            ' Account number
    Dim dPreviousCloseTime As Double    ' Previous close time
    Dim Bars As cGdBars                 ' Bars object
    Dim strPfgSymbol As String          ' PFG Symbol
    Dim dMult As Double                 ' Price multiplier
    
    If FillSummary Is Nothing Then
        Set FillSummary = g.Broker.FillSummary(vAccountNumberOrID, vSymbolOrSymbolID, -1&)
    End If
    
    If Not FillSummary Is Nothing Then
        Set astrInfo = New cGdArray
        Set Bars = New cGdBars
        SetBarProperties Bars, FillSummary.SymbolOrSymbolID
        
        strAccount = g.Broker.GetAccountNumber(vAccountNumberOrID)
        If m.astrBrokerPos.BinarySearch(strAccount & vbTab & FillSummary.Symbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
            astrInfo.SplitFields Parse(m.astrBrokerPos(lPos), vbTab, 5), ","
        End If
        
        For lIndex = 0 To astrInfo.Size - 1
            Set astrPos = New cGdArray
            astrPos.SplitFields astrInfo(lIndex), ";"
            
            strPfgSymbol = PfgSymbol(FillSummary.Symbol, dMult)
        
            Set Fill = New cPtFill
            With Fill
                .AccountID = FillSummary.AccountID
                .AutoTradingItemID = 0&
                .SymbolOrSymbolID = FillSummary.SymbolOrSymbolID
                .Quantity = Abs(CLng(Val(astrPos(0))))
                .Price = GenesisPrice(astrPos(1), dMult) ' Val(astrPos(1))
                PreviousCloseForSymbol FillSummary.SymbolOrSymbolID, dPreviousCloseTime
                .FillDate = ConvertToBrokerDate(dPreviousCloseTime, m.nBroker, FillSummary.Symbol, False)
                .SessionDate = Bars.SessionDateForTradeTime(dPreviousCloseTime)
                .BrokerID = "CarPosFix " & astrPos(2)
                .BrokerOrderID = .BrokerID
                .Buy = (CLng(Val(astrPos(0))) > 0&)
                .IsManual = True
                .IsSnapshot = False
                .Save
            End With
        Next lIndex
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.CreateFillsForCarriedInfo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasNonDemoAccount
'' Description: Determine if the user has non-demo accounts in the database
'' Inputs:      None
'' Returns:     True if non demo accounts exist, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HasNonDemoAccount() As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountType]=" & Str(m.nBroker) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        If Left(rs!AccountNumber, 1) <> "D" Then
            bReturn = True
            Exit Do
        End If
        
        rs.MoveNext
    Loop
    
    HasNonDemoAccount = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPFG.HasNonDemoAccount"
    
End Function

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

    Set m.astrToFix = New cGdArray
    m.astrToFix.Create eGDARRAY_Strings
    
    m.bWasConnected = False
    m.bVerifyPositions = True
    m.bSyncInProgress = False
    m.bFixCalledOnce = False
    m.lLastConnectDate = 0&
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPFG.Class_Initialize"
    
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
    RaiseError "cPFG.Class_Terminate"
    
End Sub
