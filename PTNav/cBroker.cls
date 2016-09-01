VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBroker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBroker.cls
'' Description: Base class functionality for all brokers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/30/2011   DAJ         Return value in StartingGenesisIdForOptNav
'' 04/05/2011   DAJ         Added SymbolInformation call
'' 05/06/2011   DAJ         If order doesn't exist for a fill, create fill anyway
'' 06/08/2011   DAJ         Change account type if logging into different Rithmic broker
'' 06/21/2011   DAJ         Separate out Simulated trading types, Last connection status change
'' 06/22/2011   DAJ         Clear VerifyPositions flag if simulated and no accounts
'' 06/24/2011   DAJ         Send unsolicited position refresh to Option Nav
'' 08/25/2011   DAJ         Utilized standard message enums, added more price conversions
'' 09/07/2011   DAJ         Add an additional "invalid login" message
'' 09/20/2011   DAJ         Don't log price change messages from broker
'' 09/22/2011   DAJ         Make sure that order for a snapshot fill is called snapshot
'' 09/26/2011   DAJ         Send Begin/Ends to broker object, added phrase to invalid login check
'' 10/05/2011   DAJ         Fix order status after a fill if necessary
'' 10/07/2011   DAJ         Handle a busted fill situation (a fill with zero quantity)
'' 10/24/2011   DAJ         Drop order/fill/position with unrecognized account
'' 11/02/2011   DAJ         Added the IsDotNet flag
'' 11/09/2011   DAJ         Added another login error message to the "do not reconnect" list
'' 12/13/2011   DAJ         Change account type if logging into different RJ O'Brien broker
'' 01/18/2012   DAJ         Enhanced logging for automated trading
'' 01/30/2012   DAJ         Allow Reconnect flag on Connection Status message
'' 01/31/2012   DAJ         Handle unsolicted cancels in TradeSense order groups
'' 02/14/2012   DAJ         New status alerts for position mismatch / auto trade disabled
'' 02/14/2012   DAJ         Added multi-leg order support
'' 03/21/2012   DAJ         Fix for amending a partially filled order
'' 04/05/2012   DAJ         Added broker mode code
'' 04/17/2012   DAJ         Handle broker that doesn't give us carried position information
'' 04/20/2012   DAJ         Send carried fills to broker view form if broker mode
'' 05/31/2012   DAJ         Turnkey implementation
'' 06/04/2012   DAJ         Removed Turnkey argument for stand-alone
'' 06/05/2012   DAJ         Don't call ChangeOrderStatus in SendOrder for SimStream order
'' 06/21/2012   DAJ         Changed BrokerOrderID on Carried Position Fixes
'' 07/19/2012   DAJ         Change account type if logging into different RCG (PATS) broker
'' 07/31/2012   DAJ         Don't reconnect to CQG if just connected and "already logged in",
''                          clear broker positions collection when BEGIN on positions
'' 08/10/2012   DAJ         Created AllowReconnect function, added 'Permission Denied'
'' 08/21/2012   DAJ         Added FcmAccountNumber for accounts
'' 08/28/2012   DAJ         CarriedFillFromInfo
'' 09/11/2012   DAJ         Handle PreviousOrderID after OrderFromInfo
'' 09/18/2012   DAJ         Don't attempt a reconnect after a "Unknown Application" error from PATS
'' 09/25/2012   DAJ         Fixed GenesisPrice for incomplete prices from broker
'' 10/05/2012   DAJ         Added Account threshold for CQG brokers
'' 10/09/2012   DAJ         If PreviousOrderID set in order, use it to result in OrderIdChanged event
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 12/11/2012   DAJ         Use the flatten queue for position reversals
'' 12/11/2012   DAJ         Handle broker status in automated trading item
'' 12/11/2012   DAJ         Capture error on copy of stand-alone and handle better
'' 01/16/2013   DAJ         Don't set status to Disconnecting for SimStream
'' 01/18/2013   DAJ         Broker held OCO for Interactive Brokers
'' 01/18/2013   DAJ         CQG symbols enabled for market data
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 03/07/2013   DAJ         Added additional string into allow reconnect function
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
'' 03/11/2013   DAJ         Last known symbols for CQG
'' 03/14/2013   DAJ         Change account type if logging into different CQG broker
'' 04/03/2013   DAJ         Add order to CQG amend orders array if amended outside Trade Navigator
'' 04/08/2013   DAJ         Don't assign automated trading item ID to an existing fill
'' 04/09/2013   DAJ         Fake a position update after an unsolicited fill update for CQG
'' 04/15/2013   DAJ         Clear refresh flags upon connection for Rithmic brokers
'' 04/17/2013   DAJ         Added direction filter to CancelWorkingOrders call
'' 04/30/2013   DAJ         If CQG Broker, ask for symbols when go connected
'' 05/21/2013   DAJ         Retain Entry flag for a single leg order when coming back from broker
'' 05/23/2013   DAJ         Added new strings to the AllowReconnect routine
'' 06/06/2013   DAJ         In SendOrder, call OrderStatusChange when Error status
'' 07/16/2013   DAJ         Send Begin/End for accounts to broker object
'' 08/02/2013   DAJ         Make sure we only show order reject message to the user once
'' 10/16/2013   DAJ         Added another check for invalid user name/password
'' 10/24/2013   DAJ         Contract size override for forex
'' 12/03/2013   DAJ         Added "SoftwareNotPermitted" to reconnect check
'' 12/06/2013   DAJ         Added "Expired" to reconnect check; Timeout waiting for app loaded
'' 12/18/2013   DAJ         Implemented Cancel/Replace for OEC Forex
'' 02/28/2014   DAJ         Change the "We will cancel your market order..." message to alert popup
'' 04/24/2014   DAJ         Confirm flatten order when account associated to lot
'' 05/13/2014   DAJ         Fix for timeout trying to start stand alone process
'' 05/19/2014   DAJ         Log session date with order
'' 05/20/2014   DAJ         Log when PositionVerify and ConnectionStatus changes
'' 07/22/2014   DAJ         Enhancements for amending orders with PATS
'' 08/22/2014   DAJ         Moved Start and Stop Standalone code into functions
'' 08/25/2014   DAJ         Don't show connection messages or reset PositionVerify flag over the weekend
'' 10/09/2014   DAJ         Allow for converting accounts from one OEC broker to another
'' 10/24/2014   DAJ         Fill Display
'' 10/29/2014   DAJ         Remove old synthetic order/MIT code
'' 11/06/2014   DAJ         Don't handle an unsolicited position message where the position hasn't changed
'' 11/10/2014   DAJ         Always save the fill in HandleFillMessage when order not found ( Jim Thorsen )
'' 11/17/2014   DAJ         Set g.Broker.PositionsToVerify when VerifyPositions changes not ConnectionStatus
'' 12/04/2014   DAJ         Remove enabled symbol check for trading
'' 12/23/2014   DAJ         Remove symbol from ToFix list after doing a fix position
'' 01/07/2015   DAJ         When clearing refresh flags, clear the flags in BrokerInfo as well
'' 01/29/2015   DAJ         Added GetContractInfo call
'' 02/18/2015   DAJ         For Rithmic brokers, cancel flatten if cancel order failed
'' 05/12/2015   DAJ         Switched the two blocks of code in the CheckHeartbeat routine
'' 09/02/2015   DAJ         Don't change a Stop order to a Limit order for Rithmic brokers
'' 09/18/2015   DAJ         Set order to CancelPending in CancelWorkingOrders if cancel sent to the broker
'' 10/06/2015   DAJ         Utilize new ClearingDate on position if there when creating carried position fills
'' 03/11/2016   DAJ         Only show rejected order popup for manual trades
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    BrokerObj As Object                 ' Calling broker object
    
    strIniFile As String                ' Ini file for the broker information
    nBroker As eTT_AccountType          ' Broker type
    strBrokerName As String             ' Broker name
    strProcessName As String            ' Name of the process to check
    strProcessPath As String            ' Full path of the process to start
    strControlID As String              ' Control ID
    strLogPath As String                ' Path for the log files
    strConnectIni As String             ' Name and path of the ini file with connection information

    nStatus As eGDConnectionStatus      ' Connection status to the broker servers
    bWasConnected As Boolean            ' This broker was has been connected
    dHeartBeat As Double                ' Tick count as of last message receieved
    bReconnecting As Boolean            ' Are we currently reconnecting?
    lReconnectAttempts As Long          ' Reconnect attempts that have been made
    bAppLoaded As Boolean               ' Is the stand-alone application loaded?
    bManualDisconnect As Boolean        ' Is the user manually disconnecting?
    bAskedPassword As Boolean           ' Have we asked for the password yet?
    dLastFillTime As Double             ' Tick count of when we received the last fill
    strLastMessage As String            ' Last message from the connection status message
    bVerifyPositions As Boolean         ' Do we want to verify the positions with the user?
    bVerboseRefresh As Boolean          ' Do we want to tell the user about the refresh?
    bSyncInProgress As Boolean          ' Are we in the process of synchronizing positions?
    bFixCalledOnce As Boolean           ' Has the fix positions call been done at least once?
    bIsDotNet As Boolean                ' Is the stand-alone application written in dot Net?
    bIsBroker As Boolean                ' Are we running in broker mode?
    dLastConnection As Double           ' Last connection time/date
    lAccountThreshold As Long           ' Number of accounts allowed for non-broker user
    bAllowBrokerOco As Boolean          ' Are we allowing a broker held OCO for this broker?
    dStartedProcess As Double           ' Tick count when we started the broker stand-alone
        
    astrAccounts As cGdArray            ' List of accounts currently connected to
    astrBrokerPos As cGdArray           ' List of current positions
    BrokerPositions As cGdTree          ' Broker positions
    astrToFix As cGdArray               ' List of account/symbol pairs that need their position fixed
    abShowMenuItem As cGdArray          ' Array of menu items with show flags
    CarriedFills As cGdTree             ' Fills that make up the carried position
    RejectShown As cGdTree              ' Collection of reject messages shown keyed by order ID
    
    WaitReplace As cGdTree              ' Collection of orders waiting for replace
    Originals As cGdTree                ' Collection of original orders being cancel/replaced
    GidMap As cGdTree                   ' Map of new Genesis ID to old Genesis ID

    dRefreshingStart As Double          ' Start of the refreshing time
    bRefreshingAccounts As Boolean      ' Are we refreshing accounts?
    bRefreshingOrders As Boolean        ' Are we refreshing orders?
    bRefreshingFills As Boolean         ' Are we refreshing fills?
    bRefreshingPositions As Boolean     ' Are we refreshing positions?
    
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

Public Property Get HeartBeat() As Double
    HeartBeat = m.dHeartBeat
End Property
Public Property Let HeartBeat(ByVal dHeartBeat As Double)
    m.dHeartBeat = dHeartBeat
End Property

Public Property Get ConnectionStatus() As eGDConnectionStatus
    ConnectionStatus = m.nStatus
End Property
Public Property Let ConnectionStatus(ByVal nNewStatus As eGDConnectionStatus)
    If nNewStatus <> m.nStatus Then
        DumpDebug vbTab & "ConnectionStatus changed from '" & g.BrokerEnums.ConnectionStatusToString(m.nStatus) & "' to '" & g.BrokerEnums.ConnectionStatusToString(nNewStatus) & "'"
        m.nStatus = nNewStatus
        
        If Not g.Broker Is Nothing Then
            g.Broker.LastStatusChange = Now
        End If
        'If nNewStatus = eGDConnectionStatus_Connected Then
        '    g.Broker.PositionsToVerify = True
        'End If
    End If
End Property

Public Property Get BrokerInfo() As cBrokerInfo
    Set BrokerInfo = m.BrokerInfo
End Property

Public Property Get Refreshing() As Boolean
    Refreshing = RefreshingAccounts Or RefreshingPositions Or RefreshingOrders Or RefreshingFills
End Property

Public Property Get PositionVerify() As Boolean
    PositionVerify = m.bVerifyPositions
End Property
Public Property Let PositionVerify(ByVal bVerifyPositions As Boolean)
    If bVerifyPositions <> m.bVerifyPositions Then
        DumpDebug vbTab & "PositionVerify changed from " & Str(m.bVerifyPositions) & " to " & Str(bVerifyPositions)
        m.bVerifyPositions = bVerifyPositions
        
        If m.bVerifyPositions = True Then
            g.Broker.PositionsToVerify = True
        End If
    End If
End Property

Public Property Get AskedPassword() As Boolean
    AskedPassword = m.bAskedPassword
End Property
Public Property Let AskedPassword(ByVal bAskedPassword As Boolean)
    m.bAskedPassword = bAskedPassword
End Property

Public Property Get AppLoaded() As Boolean
    AppLoaded = m.bAppLoaded
End Property
Public Property Let AppLoaded(ByVal bAppLoaded As Boolean)
On Error GoTo ErrSection:

    m.bAppLoaded = bAppLoaded
    
    If bAppLoaded = True Then
        StartedProcess = 0#
        
        If LiveTradingAllowed(m.nBroker) Then
            m.BrokerObj.SendConnect
        Else
            DumpDebug "Unloading the application since live trading is not allowed"
            m.BrokerObj.UnloadApp
        End If
    Else
        If m.bManualDisconnect Then
            HeartBeat = 0#
        Else
            HeartBeat = gdTickCount - 110000
        End If
        m.bManualDisconnect = False
        ConnectionStatus = eGDConnectionStatus_Disconnected
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBroker.AppLoaded.Let"
    
End Property

Public Property Get ShowTradeMenuItem(ByVal nItem As eGDTradingMenu) As Boolean
    ShowTradeMenuItem = (m.abShowMenuItem(nItem) = 1)
End Property
Public Property Let ShowTradeMenuItem(ByVal nItem As eGDTradingMenu, ByVal bShowItem As Boolean)
    If bShowItem Then
        m.abShowMenuItem(nItem) = 1
    Else
        m.abShowMenuItem(nItem) = 0
    End If
End Property

Public Property Get Accounts() As cGdArray
    Set Accounts = m.astrAccounts
End Property

Public Property Get CarriedFills() As cGdTree
    Set CarriedFills = m.CarriedFills
End Property

Public Property Get IsBroker() As Boolean
    IsBroker = m.bIsBroker
End Property
Public Property Let IsBroker(ByVal bIsBroker As Boolean)
    m.bIsBroker = bIsBroker
End Property

Public Property Get AllowBrokerOco() As Boolean
    AllowBrokerOco = m.bAllowBrokerOco
End Property
Public Property Let AllowBrokerOco(ByVal bAllowBrokerOco As Boolean)
    m.bAllowBrokerOco = bAllowBrokerOco
End Property

Public Property Get LastDateConnected() As Long
    LastDateConnected = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", m.strIniFile))))
End Property
Private Property Let LastDateConnected(ByVal lLastDateConnected As Long)
    SetIniFileProperty "Last", EncryptToHex(Str(lLastDateConnected)), "Connect", m.strIniFile
End Property

Public Property Get ContractSize(ByVal strGenesisSymbol As String, ByVal strAccountNumber As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the property
    Dim strKey As String                ' Key into the collection
    
    lReturn = kNullData
    If g.Broker.IsOecBroker(m.nBroker) Then
        If Not m.BrokerObj Is Nothing Then
            lReturn = m.BrokerObj.ContractSize(strGenesisSymbol, strAccountNumber)
        End If
    End If
    
    ContractSize = lReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBroker.ContractSize.Get"
    
End Property

Private Property Get RefreshingAccounts() As Boolean
    RefreshingAccounts = m.bRefreshingAccounts
End Property
Private Property Let RefreshingAccounts(ByVal bRefreshingAccounts As Boolean)
    If m.bRefreshingAccounts <> bRefreshingAccounts Then
        m.bRefreshingAccounts = bRefreshingAccounts
        DumpDebug vbTab & vbTab & "RefreshingAccounts = " & Str(m.bRefreshingAccounts)
    End If
End Property

Private Property Get RefreshingOrders() As Boolean
    RefreshingOrders = m.bRefreshingOrders
End Property
Private Property Let RefreshingOrders(ByVal bRefreshingOrders As Boolean)
    If m.bRefreshingOrders <> bRefreshingOrders Then
        m.bRefreshingOrders = bRefreshingOrders
        DumpDebug vbTab & vbTab & "RefreshingOrders = " & Str(m.bRefreshingOrders)
    End If
End Property

Private Property Get RefreshingFills() As Boolean
    RefreshingFills = m.bRefreshingFills
End Property
Private Property Let RefreshingFills(ByVal bRefreshingFills As Boolean)
    If m.bRefreshingFills <> bRefreshingFills Then
        m.bRefreshingFills = bRefreshingFills
        DumpDebug vbTab & vbTab & "RefreshingFills = " & Str(m.bRefreshingFills)
    End If
End Property

Private Property Get RefreshingPositions() As Boolean
    RefreshingPositions = m.bRefreshingPositions
End Property
Private Property Let RefreshingPositions(ByVal bRefreshingPositions As Boolean)
    If m.bRefreshingPositions <> bRefreshingPositions Then
        m.bRefreshingPositions = bRefreshingPositions
        DumpDebug vbTab & vbTab & "RefreshingPositions = " & Str(m.bRefreshingPositions)
    End If
End Property

Private Property Get StartedProcess() As Double
    StartedProcess = m.dStartedProcess
End Property
Private Property Let StartedProcess(ByVal dStartedProcess As Double)
    If dStartedProcess <> m.dStartedProcess Then
        DumpDebug vbTab & vbTab & "Started Process changed from " & Str(m.dStartedProcess) & " to " & dStartedProcess
        m.dStartedProcess = dStartedProcess
    End If
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

    Dim lReturn As Long                 ' Return value for the function
    Dim strKey As String                ' Key into the collection
    
    lReturn = 0&
    strKey = strAccountNumber & vbTab & strSymbol
    If m.BrokerPositions.Exists(strKey) Then
        lReturn = m.BrokerPositions(strKey).CurrentPosition
    End If
    
    CurrentPosition = lReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBroker.CurrentPosition.Get"
    
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

    Dim lReturn As Long                 ' Return value for the function
    Dim strKey As String                ' Key into the collection
    
    lReturn = 0&
    strKey = strAccountNumber & vbTab & strSymbol
    If m.BrokerPositions.Exists(strKey) Then
        lReturn = m.BrokerPositions(strKey).CarriedPosition
    End If
    
    CarriedPosition = lReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBroker.CarriedPosition.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize this object
'' Inputs:      Calling Broker Object, Broker, Log Path, Dot Net?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(BrokerObj As Object, ByVal nBroker As eTT_AccountType, ByVal strLogPath As String, Optional ByVal bIsDotNet As Boolean = False)
On Error GoTo ErrSection:

    Set m.BrokerObj = BrokerObj
    m.nBroker = nBroker
    m.strLogPath = strLogPath
    m.bIsDotNet = bIsDotNet
    
    If Not DirExist(m.strLogPath) Then
        MkDir m.strLogPath
    End If
    KillFile AddSlash(m.strLogPath) & "*.LOG /o=-30"
    
    Set m.BrokerInfo = New cBrokerInfo
    m.BrokerInfo.Init m.nBroker, m.strLogPath
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.Init"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CopyProcess
'' Description: Attempt to copy the process
'' Inputs:      Source
'' Returns:     True if success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CopyProcess(ByVal strSource As String) As Boolean
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout variable
    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If StopStandalone = True Then
        DumpDebug "Killing Process '" & m.strProcessName & "'"
        
        lTimeOut = 0&
        Do While (StopStandalone = True) And (lTimeOut < 5&)
            Sleep 1
            lTimeOut = lTimeOut + 1&
        Loop
    End If
    
    If (StopStandalone = False) Then
        DumpDebug "Copying '" & strSource & "' to '" & m.strProcessPath & "'"
        bReturn = CopyStandalone(strSource, m.strProcessPath)
    End If
    
    CopyProcess = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.CopyProcess"
    
End Function

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

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Not m.BrokerObj Is Nothing Then
        bReturn = m.BrokerObj.ShowConfig
    End If
    
    ShowConfig = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.ShowConfig"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowBrokerConnectionInfo
'' Description: Allow the user to view their broker connection information
'' Inputs:      New?, User Name, Connect?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowBrokerConnectionInfo(Optional ByVal bNew As Boolean = False, Optional ByVal strUserName As String = "", Optional ByVal bConnect As Boolean = True)
On Error GoTo ErrSection:

    If Not m.BrokerObj Is Nothing Then
        If m.BrokerObj.ShowBrokerConnectionInfo(bNew, strUserName) Then
            If bConnect Then
                Connect
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.ShowBrokerConnectionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Connect
'' Description: Call the connect routine for the broker
'' Inputs:      New User ID, Are we switching?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Connect(Optional ByVal strNewUserID As String = "", Optional ByVal bSwitching As Boolean = False)
On Error GoTo ErrSection:

    Dim bConnect As Boolean             ' Should we still try to connect?
    Dim strProcess As String            ' Name of the process to start
    Dim lTimeOut As Long                ' Timeout variable
    Dim strArgs As String               ' Arguments to send to the standalone
    
    If (g.bUnloading = False) And (Not m.BrokerObj Is Nothing) Then
        If m.BrokerObj.ShowLogin(strNewUserID, bSwitching) Then
            If LiveTradingAllowed(m.nBroker) Then
                If m.nBroker <> eTT_AccountType_Etrade Then
                    If m.bIsDotNet Then
                        bConnect = HasDotNet
                    Else
                        bConnect = True
                    End If
                    
                    If bConnect = True Then
                        LoadAccountThreshold
                        ConnectionStatus = eGDConnectionStatus_Connecting
                        
                        DumpDebug "Trade Navigator version: " & FormatVersion & "." & Str(App.Revision) & " " & DateFormat(FileDate(App.Path & "\" & App.EXEName & ".EXE"), MM_DD_YYYY, HH_MM, AMPM_UPPER)
                        If Len(m.strProcessName) > 0 Then
                            If StartStandalone = False Then
                                m.BrokerObj.SendConnect
                            End If
                        Else
                            m.BrokerObj.SendConnect
                        End If
                    End If
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.Connect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disconnect
'' Description: Call the disconnect routine for the broker
'' Inputs:      Force Reconnect?, Reason for Disconnect
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Disconnect(Optional ByVal bForceReconnect As Boolean = False, Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    If (ConnectionStatus = eGDConnectionStatus_Connected) Then
        ' 01/16/2013 DAJ: Can't remember exactly why I added the Disconnecting status here ( probably
        ' just for completeness ).  Unfortunately, because of a timing issue, if we do it for Sim Trade
        ' with working orders, we get stuck in a loop trying to Cancel ( or Park ) the working order
        ' but we can't because we are not connected...
        If m.nBroker <> eTT_AccountType_SimStream Then
            ConnectionStatus = eGDConnectionStatus_Disconnecting
        End If
        
        m.bManualDisconnect = Not bForceReconnect
        If m.bManualDisconnect Then
            HeartBeat = 0#
            g.OrderStrategies.DeactivateExitsForBroker m.nBroker, "Disconnecting from " & m.strBrokerName, True
        End If
        
        If Len(strReason) = 0 Then
            DumpDebug "Disconnecting from " & m.strBrokerName
        Else
            DumpDebug "Disconnecting from " & m.strBrokerName & ": " & strReason
        End If
        
        If Not m.BrokerObj Is Nothing Then
            m.BrokerObj.SendDisconnect
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.Disconnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reconnect
'' Description: Attempt to reconnect to the broker servers
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
                        
        If (ConnectionStatus <> eGDConnectionStatus_Connected) And (ConnectionStatus <> eGDConnectionStatus_Connecting) Then
            If (m.bAppLoaded = False) Or (lTimeOut >= 30&) Then
                If StopStandalone = True Then
                    DumpDebug "Process '" & m.strProcessName & "' killed"
                    Sleep 3
                End If
                StartedProcess = 0#
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
    RaiseError "cBroker.Reconnect"
    
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

    Dim nNewStatus As eTT_OrderStatus   ' New order status

    If NotSent(Order.Status) Then
        If AddOrder(Order) = False Then
            nNewStatus = eTT_OrderStatus_Error
        Else
            nNewStatus = eTT_OrderStatus_Sent
        End If
    Else
        If AmendOrder(Order) = False Then
            nNewStatus = eTT_OrderStatus_Error
        Else
            nNewStatus = eTT_OrderStatus_AmendPending
        End If
    End If
                            
    ' Don't call this for SimBroker and Simulated Stream accounts because the AddOrder and
    ' AmendOrder calls for those brokers already call this...
    ' 06/06/2013 DAJ: Make sure to do this no matter what if the order status is Error, because
    ' the SimStream and SimBroker objects do NOT call it in that case...
    If (nNewStatus = eTT_OrderStatus_Error) Or ((m.nBroker <> eTT_AccountType_SimBroker) And (m.nBroker <> eTT_AccountType_SimStream)) Then
        Order.ChangeOrderStatus nNewStatus, m.BrokerInfo
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.SendOrder"
    
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

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Not m.BrokerObj Is Nothing Then
        If CanSendOrder(Order, "AddOrder") Then
            DumpDebug "AddOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
            m.BrokerObj.AddOrder Order
            bReturn = True
        End If
    End If
    
    AddOrder = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.AddOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AmendOrder
'' Description: Attempt to amend the order
'' Inputs:      Order to add
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AmendOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim OriginalOrder As cPtOrder       ' Original order
    
    bReturn = False
    If Not m.BrokerObj Is Nothing Then
        If CanSendOrder(Order, "AmendOrder") Then
            If UseAmend(Order) Then
                DumpDebug "AmendOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
                m.BrokerObj.AmendOrder Order
            Else
                Set OriginalOrder = New cPtOrder
                OriginalOrder.Load Order.OrderID
                
                DumpDebug "CancelReplaceOrder: " & Order.OrderText(True, True, True)
                If m.Originals.Exists(OriginalOrder.GenesisOrderID) Then
                    DumpDebug vbTab & "Order replaced in Originals: " & OriginalOrder.OrderText(True, True, True)
                    Set m.Originals(OriginalOrder.GenesisOrderID) = OriginalOrder
                Else
                    DumpDebug vbTab & "Order added to Originals: " & OriginalOrder.OrderText(True, True, True)
                    m.Originals.Add OriginalOrder, OriginalOrder.GenesisOrderID
                End If
                
                If m.WaitReplace.Exists(Order.BrokerID) Then
                    DumpDebug vbTab & "Order replaced in WaitReplace: " & Order.OrderText(True, True, True)
                    Set m.WaitReplace(Order.BrokerID) = Order
                Else
                    DumpDebug vbTab & "Order added to WaitReplace: " & Order.OrderText(True, True, True)
                    m.WaitReplace.Add Order, Order.BrokerID
                End If
                
                m.BrokerObj.CancelOrder OriginalOrder
            End If
            
            bReturn = True
        End If
    End If
    
    AmendOrder = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.AmendOrder"
    
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
    Dim strBrokerSymbol As String       ' Broker Symbol
    Dim strBrokerOrderType As String    ' Broker Order Type
    Dim bReturn As Boolean              ' Return value for the function
    Dim strBrokerBase As String         ' Broker base symbol
    Dim strBrokerExchange As String     ' Broker exchange
    Dim strMessage As String            ' Message to display to user
    
    bReturn = False
    If Not m.BrokerObj Is Nothing Then
        strBrokerSymbol = m.BrokerObj.BrokerSymbol(Order.Symbol)
        strBrokerOrderType = m.BrokerObj.BrokerOrderType(Order.OrderType)
        
        If NotSent(Order.Status) Then
            DumpDebug "CancelOrder: '" & Order.OrderText(True, True, True) & "'"
            Order.ChangeOrderStatus eTT_OrderStatus_Cancelled, m.BrokerInfo
        ElseIf Len(strBrokerSymbol) = 0 Then
            Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
            g.Broker.ShowUnknownSymbolError Order.Symbol, m.nBroker, "CancelOrder", True, Me
        ' 12/04/2014 DAJ: With the changes in exchange fees coming from the exchange, there is more of a
        ' possibility that traders could have data turned off, but trading turned on for symbols.  Because of
        ' this, we need to let them trade symbols that they don't get data for...
        'ElseIf IsEnabledSymbol(Order.Symbol, strBrokerBase, strBrokerExchange) = False Then
        '    Order.Message = Order.Symbol & " not enabled through your " & m.strBrokerName & " account"
        '    g.Broker.ShowNotEnabledForSymbolError Order.Symbol, m.nBroker, strBrokerBase, strBrokerExchange, "CancelOrder", True, Me
        ElseIf Len(strBrokerOrderType) = 0 Then
            strMessage = OrderType(Order.OrderType) & " is not a valid order type for " & m.strBrokerName
            Order.Message = OrderType(Order.OrderType) & " cannot be converted to a " & m.strBrokerName & " order type"
        
            DumpDebug "Cancel Order Error shown to user: '" & Trim(Replace(strMessage, "|", " ")) & "'"
            InfBox strMessage, "!", , m.strBrokerName & " Order Error"
        ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
            Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
            g.Broker.ShowNotConnectedError Order.AccountID, m.nBroker, "CancelOrder", True, Me
        ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
            strMessage = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
            Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
            
            DumpDebug "Cancel Order Error shown to user: '" & Trim(Replace(strMessage, "|", " ")) & "'"
            InfBox strMessage, "!", , m.strBrokerName & " Order Error"
        ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
            strMessage = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
            Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
            
            DumpDebug "Cancel Order Error shown to user: '" & Trim(Replace(strMessage, "|", " ")) & "'"
            InfBox strMessage, "!", , m.strBrokerName & " Order Error"
        Else
            If Not g.TsoGroups Is Nothing Then
                g.TsoGroups.CancelRequested Order
            End If
            
            DumpDebug "CancelOrder: " & Order.OrderText(True, True, True)
            m.BrokerObj.CancelOrder Order
                                    
            ' If the user is cancelling a Market order, warn them that they would not
            ' be able to do so with a live brokerage account...
            If (Order.OrderType = eTT_OrderType_Market) And (g.Broker.IsLiveAccount(Order.Broker) = False) Then
                'InfBox "We will cancel your Market order, but please be advised that you will not be able to cancel Market orders with a Live broker", "i", , "Cancelling Market Order"
                ShowAlertPopup "We will cancel your Market order, but please be|advised that you will not be able to cancel|Market orders with a Live broker", "Cancelling Market Order", vbCenter
            End If
            
            bReturn = True
        End If
    End If
    
    CancelOrder = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.CancelOrder"
    
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
    Dim strBrokerSymbol As String       ' Broker Symbol
    Dim strBrokerOrderType As String    ' Broker Order Type
    Dim bReturn As Boolean              ' Return value for the function
    Dim strBrokerBase As String         ' Broker base symbol
    Dim strBrokerExchange As String     ' Broker exchange
    Dim strMessage As String            ' Message to display to user
    
    bReturn = False
    If Not m.BrokerObj Is Nothing Then
        strBrokerSymbol = m.BrokerObj.BrokerSymbol(Order.Symbol)
        strBrokerOrderType = m.BrokerObj.BrokerOrderType(Order.OrderType)
        
        If Len(strBrokerSymbol) = 0 Then
            Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
            g.Broker.ShowUnknownSymbolError Order.Symbol, m.nBroker, "ParkOrder", True, Me
        ' 12/04/2014 DAJ: With the changes in exchange fees coming from the exchange, there is more of a
        ' possibility that traders could have data turned off, but trading turned on for symbols.  Because of
        ' this, we need to let them trade symbols that they don't get data for...
        'ElseIf IsEnabledSymbol(Order.Symbol, strBrokerBase, strBrokerExchange) = False Then
        '    Order.Message = Order.Symbol & " not enabled through your " & m.strBrokerName & " account"
        '    g.Broker.ShowNotEnabledForSymbolError Order.Symbol, m.nBroker, strBrokerBase, strBrokerExchange, "ParkOrder", True, Me
        ElseIf Len(strBrokerOrderType) = 0 Then
            strMessage = OrderType(Order.OrderType) & " is not a valid order type for " & m.strBrokerName
            Order.Message = OrderType(Order.OrderType) & " cannot be converted to a " & m.strBrokerName & " order type"
        
            DumpDebug "Park Order Error shown to user: '" & Trim(Replace(strMessage, "|", " ")) & "'"
            InfBox strMessage, "!", , m.strBrokerName & " Order Error"
        ElseIf NotSent(Order.Status) Then
            DumpDebug "ParkOrder: '" & Order.OrderText(True, True, True) & "'"
            Order.ChangeOrderStatus eTT_OrderStatus_Parked, m.BrokerInfo
        ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
            Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
            g.Broker.ShowNotConnectedError Order.AccountID, m.nBroker, "ParkOrder", True, Me
        ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
            strMessage = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
            Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
            
            DumpDebug "Park Order Error shown to user: '" & Trim(Replace(strMessage, "|", " ")) & "'"
            InfBox strMessage, "!", , m.strBrokerName & " Order Error"
        ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
            strMessage = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
            Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
            
            DumpDebug "Park Order Error shown to user: '" & Trim(Replace(strMessage, "|", " ")) & "'"
            InfBox strMessage, "!", , m.strBrokerName & " Order Error"
        Else
            DumpDebug "ParkOrder: " & Order.OrderText(True, True, True)
            m.BrokerObj.ParkOrder Order
            
            Order.Status = eTT_OrderStatus_ParkPending
            Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
            Order.Save
            m.BrokerInfo.AddOrder Order, False
            
            bReturn = True
        End If
    End If
    
    ParkOrder = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.ParkOrder"
    
End Function

#If 0 Then
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
    RaiseError "cBroker.CancelOrderFromBrokerID"
    
End Sub
#End If

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
                ' DAJ 09/18/2015: Customer had an issue with the flatten where he was sending cancel requests to Rithmic every 100ms
                ' since I changed the timer interval on the flatten.  By setting the order status to Cancel Pending here if the
                ' cancel was sent to the broker, we won't get back in here because of the OrderIsPending check above...
                If CancelOrder(Orders(lIndex)) = True Then
                    Orders(lIndex).ChangeOrderStatus eTT_OrderStatus_CancelPending, m.BrokerInfo
                End If
                
                lReturn = lReturn + 1&
            End If
        Next lIndex
    End If
    
    CancelWorkingOrders = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.CancelWorkingOrders"
    
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
            If bConfirm Then
                DumpDebug "User Informed: " & strMessage & " in account " & strAccount
            Else
                DumpDebug "User Not Informed: " & strMessage & " in account " & strAccount
            End If
            
            bReturn = ExitPositionForSymbol(strAccount, strSymbol, lAtID, lPosition, strGenesisOrderID)
        End If
    End If
    
    FlattenPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.FlattenPosition"
    
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
    Dim lQuantity As Long               ' Quantity for the order
    Dim strMessage As String            ' Message to display to the user
    Dim strReturn As String             ' Return from user question
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    DumpDebug "Reversing Position for: " & strAccount & ", " & strSymbol & ", " & Str(lAtID)
    lPosition = m.BrokerInfo.CurrentPosition(strAccount, strSymbol, lAtID)
    If lPosition <> 0& Then
        lQuantity = lPosition * 2
    
        If lPosition < 0 Then
            strMessage = "BUY " & Abs(lQuantity) & " " & strSymbol & " at the MARKET"
        Else
            strMessage = "SELL " & Abs(lQuantity) & " " & strSymbol & " at the MARKET"
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
    RaiseError "cBroker.ReversePosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrdersAsOco
'' Description: Submit the given orders as an Order-Cancels-Order
'' Inputs:      Order1, Order2
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SubmitOrdersAsOco(Order1 As cPtOrder, Order2 As cPtOrder) As Boolean
On Error GoTo ErrSection:
    
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If m.bAllowBrokerOco = True Then
        If Not m.BrokerObj Is Nothing Then
            If OrderTypeAllowedForOco(Order1.OrderType) Then
                If OrderTypeAllowedForOco(Order2.OrderType) Then
                    If CanSendOrder(Order1, "SubmitOCO") = True Then
                        If CanSendOrder(Order2, "SubmitOCO") = True Then
                            bReturn = m.BrokerObj.SubmitOrdersAsOco(Order1, Order2)
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    SubmitOrdersAsOco = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.SumbitOrdersAsOco"
    
End Function

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
            If Not m.BrokerObj Is Nothing Then
                m.BrokerObj.Refresh bVerbose
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.Refresh"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Request accounts from the broker server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAccounts()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    If Not Refreshing Then
        RefreshingAccounts = True
        
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for account information...", , , m.strBrokerName & " Refresh", True
        End If
        
        If Not m.BrokerObj Is Nothing Then
            m.BrokerObj.GetAccounts
        End If
    Else
        DumpDebug "Accounts requested but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.GetAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Request orders from the broker server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetOrders(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    If Not Refreshing Then
        RefreshingOrders = True
        
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for order information...", , , m.strBrokerName & " Refresh", True
        End If
        
        If Not m.BrokerObj Is Nothing Then
            m.BrokerObj.GetOrders strAccount
        End If
    Else
        DumpDebug "Orders requested but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.GetOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFills
'' Description: Request fills from the broker server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetFills(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    If Not Refreshing Then
        RefreshingFills = True
        
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for fill information...", , , m.strBrokerName & " Refresh", True
        End If
        
        If Not m.BrokerObj Is Nothing Then
            m.BrokerObj.GetFills strAccount
        End If
    Else
        DumpDebug "Fills requested but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.GetFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositions
'' Description: Request positions from the broker server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPositions(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    If Not Refreshing Then
        RefreshingPositions = True
        
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for position information...", , , m.strBrokerName & " Refresh", True
        End If
        
        If Not m.BrokerObj Is Nothing Then
            m.BrokerObj.GetPositions strAccount
        End If
    Else
        DumpDebug "Positions requested but a refresh is already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetContractInfo
'' Description: Request contract information from the server
'' Inputs:      Genesis Symbol, Show Message?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetContractInfo(ByVal strGenesisSymbol As String, Optional ByVal bShowMessage As Boolean = False)
On Error GoTo ErrSection:

    If g.Broker.IsIbBroker(m.nBroker) Or g.Broker.IsRithmicBroker(m.nBroker) Then
        m.BrokerObj.GetContractInfo strGenesisSymbol, bShowMessage
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.GetContractInfo"
    
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
    RaiseError "cBroker.ConnectionStatusForAccount"
    
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

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Not m.BrokerObj Is Nothing Then
        strReturn = m.BrokerObj.NextGenesisID(strAccountNumber)
    End If
    
    NextGenesisID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.NextGenesisID"
    
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

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Not m.BrokerObj Is Nothing Then
        strReturn = m.BrokerObj.StartingGenesisIdForOptNav
    End If
    
    StartingGenesisIdForOptNav = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.StartingGenesisIdForOptNav"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckHeartbeat
'' Description: Check the heartbeat to make sure that we are still connected
''              to the broker servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckHeartbeat()
On Error GoTo ErrSection:

    If FileExist(AddSlash(App.Path) & "SkipHeartbeat.FLG") = False Then
        If g.bUnloading = False Then
            ' 05/12/2015 DAJ: Switched these two blocks of code around because the heartbeat
            ' doesn't get reset on a non-manual disconnect, so we were always getting into the
            ' heartbeat block instead of this block...
            If (ConnectionStatus = eGDConnectionStatus_Connecting) And (StartedProcess > 0#) Then
                If StartedProcess < gdTickCount - 15000 Then
                    DumpDebug "Timed out waiting for application to load"
                    ConnectionStatus = eGDConnectionStatus_Disconnected
                
                    If StopStandalone = True Then
                        DumpDebug "Process '" & m.strProcessName & "' killed"
                    End If
                
                    StartedProcess = 0#
                End If
            
            ElseIf HeartBeat <> 0 Then
                If (HeartBeat < (gdTickCount - 120000)) Then
                    If m.lReconnectAttempts = 1& Then
                        InfBox "Communication has been lost with " & m.strBrokerName & ".  Attempting to reconnect...", , "+-OK", m.strBrokerName & " Connection", True
                        DumpDebug "Communication has been lost with " & m.strBrokerName & ".  Attempting to reconnect..."
                    End If
                    Reconnect
                ElseIf (ConnectionStatus = eGDConnectionStatus_Connected) Then
                    If m.lReconnectAttempts > 1& Then
                        InfBox "TradeNavigator has reconnected to " & m.strBrokerName, , "+-OK", m.strBrokerName & " Connection", True
                        DumpDebug "TradeNavigator has reconnected to " & m.strBrokerName
                    End If
                    m.lReconnectAttempts = 0&
                End If
        
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.CheckHeartbeat"
    
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

    If IsBroker = False Then
        If (ConnectionStatus = eGDConnectionStatus_Connected) And (m.astrAccounts.Size > 0) And (m.bSyncInProgress = False) Then
            If ((PositionVerify = True) Or (bForce = True)) And (m.bFixCalledOnce = True) Then
                ' Only do this if there is not a modal dialog up...
                If frmMain.Enabled Then
                    PositionVerify = False
                    
                    bReturn = frmPositionConfirm.ShowMe(m.nBroker, strAccount, strSymbol, True, True)
                    If (Len(strAccount) = 0) And (Len(strSymbol) = 0) Then
                        PositionVerify = Not bReturn
                    End If
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.VerifyPositions"
    
End Sub

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
            
            For lIndex = m.astrToFix.Size - 1 To 0 Step -1
                astrEntry.SplitFields m.astrToFix(lIndex), vbTab
                If (bOnlyFirstTime = False) Or (astrEntry(2) = "0") Then
                    If FixPosition(astrEntry(0), astrEntry(1)) = True Then
                        DumpDebug "Symbol removed from To Fix list ( FixPositions ) : " & astrEntry(0) & ", " & astrEntry(1)
                        m.astrToFix.Remove lIndex
                    End If
                End If
            Next lIndex
        End If
            
        If m.bFixCalledOnce = False Then m.bFixCalledOnce = True
        
        bInProgress = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.FixPositions"
    
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

    bReturn = False
    If FillSummary Is Nothing Then
        Set FillSummary = g.Broker.FillSummary(vAccountNumberOrID, vSymbolOrSymbolID, -1&)
    End If
    
    If Not FillSummary Is Nothing Then
        strAccount = g.Broker.GetAccountNumber(vAccountNumberOrID)
        strSymbol = GetSymbol(vSymbolOrSymbolID)
        lBrokerPos = CurrentPosition(strAccount, FillSummary.Symbol)
        
        lBrokerCarried = CarriedPosition(strAccount, FillSummary.Symbol)
        If lBrokerCarried = kNullData Then
            lBrokerCarried = lBrokerPos - FillSummary.NumBuysSnapshot + FillSummary.NumSellsSnapshot
        End If
        
        ' If we have a carried position mismatch...
        If lBrokerCarried <> FillSummary.CurrentPosition Then
            DumpDebug "Carried Position Mismatch for " & strSymbol & " in " & strAccount & ": " & m.strBrokerName & " = " & Str(lBrokerCarried) & " ; Trade Navigator = " & Str(FillSummary.CurrentPosition)
            
            ' If the broker says that the user did not carry a position, but we do...
            If (lBrokerCarried = 0&) And (FillSummary.CurrentPosition <> 0&) Then
                ' Create a "fake fill" for yesterday's session at yesterday's closing price to flatten our position...
                CreateFillAtPreviousClose vAccountNumberOrID, vSymbolOrSymbolID, FillSummary
                g.Broker.RebuildFillSummaryForSymbol FillSummary.AccountID, FillSummary.SymbolOrSymbolID, 0&, True
                bReturn = True
                
            ' If the broker says that the user did carry a position, but we do not...
            ElseIf (lBrokerCarried <> 0&) And (FillSummary.CurrentPosition = 0&) Then
                ' Create a fake fill for yesterday's session with the carried position info from the broker...
                CreateFillsForCarriedInfo vAccountNumberOrID, vSymbolOrSymbolID, FillSummary
                g.Broker.RebuildFillSummaryForSymbol FillSummary.AccountID, FillSummary.SymbolOrSymbolID, 0&, True
                bReturn = True
            
            ' If we both say that the user carried a position, but we don't agree...
            Else
                ' Create a "fake fill" for yesterday's session at yesterday's fill price to flatten our position...
                CreateFillAtPreviousClose vAccountNumberOrID, vSymbolOrSymbolID, FillSummary
                
                ' Create a fake fill for yesterday's session with the carried position info from the broker...
                CreateFillsForCarriedInfo vAccountNumberOrID, vSymbolOrSymbolID, FillSummary
                g.Broker.RebuildFillSummaryForSymbol FillSummary.AccountID, FillSummary.SymbolOrSymbolID, 0&, True
                bReturn = True
            End If
        
            m.BrokerInfo.CarriedMatch(strAccount, FillSummary.Symbol) = bReturn
        End If
    End If
    
    FixPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.FixPosition"
    
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

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Not m.BrokerObj Is Nothing Then
        bReturn = m.BrokerObj.IsTradeableSymbol(strGenesisSymbol)
    End If

    IsTradeableSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.IsTradeableSymbol"
    
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

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Not m.BrokerObj Is Nothing Then
        bReturn = m.BrokerObj.IsEnabledSymbol(strGenesisSymbol, strBrokerBase, strBrokerExchange)
    End If

    IsEnabledSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.IsEnabledSymbol"
    
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
    
    bReturn = False
    If Not m.BrokerObj Is Nothing Then
        bReturn = m.BrokerObj.OrderTypeAllowed(nOrderType, strGenesisSymbol)
    End If
    
    OrderTypeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.OrderTypeAllowed"
    
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
    
    bReturn = False
    If Not m.BrokerObj Is Nothing Then
        bReturn = m.BrokerObj.TimeInForceAllowed(nTimeInForce, strGenesisSymbol)
    End If
    
    TimeInForceAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.TimeInForceAllowed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendMessage
'' Description: Send a message to the stand-alone application
'' Inputs:      Message Type, Message, Send Now?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendMessage(ByVal lMessageType As Long, ByVal strMessage As String, Optional ByVal bSendNow As Boolean = True)
On Error GoTo ErrSection:

    frmOnlineBroker.gdBroker.CreateMessage m.strControlID, lMessageType, strMessage, , bSendNow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.SendMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleConnectionInfo
'' Description: Do the appropriate thing with the given connection information
'' Inputs:      Connection Information, Allow reconnect?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleConnectionInfo(ByVal nStatus As eGDConnectionStatus, ByVal strError As String, ByVal strUserName As String, Optional ByVal bAllowReconnect As Boolean = True)
On Error GoTo ErrSection:

    Dim nPrev As eGDConnectionStatus    ' Previous Connection status
    Dim frm As frmAlertPopup            ' Alert popup message
    Static strLastUser As String        ' Last user name
    
    Select Case nStatus
        Case eGDConnectionStatus_Disconnected
            m.BrokerInfo.SetConnectionStatus eGDConnectionStatus_Disconnected
            
            If (ConnectionStatus <> eGDConnectionStatus_Disconnected) Then
                ' DAJ 08/25/2014: As per Tim, don't show a connection error message or reset the
                ' position verify flag if we are between Friday night and Sunday morning...
                If (m.bManualDisconnect = True) Or ((Len(strError) > 0) And (IsTheWeekend = False)) Then
                    HeartBeat = 0#
                    PositionVerify = True
                End If
                
                If m.bWasConnected = True Then
                    m.bWasConnected = False
                    g.Alerts.CheckBrokerStatusAlert m.nBroker
                End If
                
                m.astrAccounts.Clear
                
                ClearRefreshFlags
                
                DumpDebug "Unloading Application since we have been disconnected"
                m.BrokerObj.UnloadApp
            End If
        
        Case eGDConnectionStatus_Disconnecting
            m.BrokerInfo.SetConnectionStatusForUser eGDConnectionStatus_Disconnecting, strLastUser
        
        Case eGDConnectionStatus_Connecting
            m.BrokerInfo.SetConnectionStatusForUser eGDConnectionStatus_Connecting, strUserName
        
        Case eGDConnectionStatus_Connected
            ' With the Rithmic connection, we actually have connections to three different
            ' servers.  If one or two go down, but not all three, we never get to a 'Disconnected'
            ' status, so the refresh flags don't get cleared...
            If ((ConnectionStatus <> eGDConnectionStatus_Connected) And (g.Broker.IsRithmicBroker(m.nBroker) = True)) Then
                ClearRefreshFlags
            End If
            
            m.bWasConnected = True
    
    End Select
    
    If strUserName <> strLastUser Then
        strLastUser = strUserName
    End If
    
    If (nStatus <> ConnectionStatus) Or (Len(strError) > 0) Then
        nPrev = ConnectionStatus
        ConnectionStatus = nStatus
        
        g.Broker.ConnectionStatusChanged m.nBroker, nStatus
        
        DumpDebug vbTab & m.strBrokerName & " Message Received (ConnectionInfo): " & vbTab & Str(nStatus) & vbTab & strError
        If (nPrev <> eGDConnectionStatus_Connected) And (ConnectionStatus = eGDConnectionStatus_Connected) Then
            LastDateConnected = Date
            m.dLastConnection = CurrentTime
            
            If Not m.BrokerObj Is Nothing Then
                If IsBroker Then
                    If g.Broker.IsCqgBroker(m.nBroker) Then
                        m.BrokerObj.GetSymbols
                    End If
                Else
                    m.BrokerObj.JustConnected
                End If
            End If
        End If
    End If
    
    ' DAJ 08/25/2014: As per Tim, don't show a connection error message or reset the
    ' position verify flag if we are between Friday night and Sunday morning...
    If (Len(strError) > 0) And (strError <> m.strLastMessage) And (IsTheWeekend = False) Then
        If (bAllowReconnect = False) Or (AllowReconnect(strError) = False) Then
            m.bManualDisconnect = True
            HeartBeat = 0#
            PositionVerify = True
            m.bAskedPassword = False
        End If
        
        Set frm = New frmAlertPopup
        frm.ShowMessageBox strError, m.strBrokerName & " Message", vbCenter
    End If
    m.strLastMessage = strError
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.HandleConnectionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StatusReceived
'' Description: Connection status message received from the broker
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub StatusReceived(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim nStatus As eGDConnectionStatus  ' Connection status
    Dim strError As String              ' Error message from server
    Dim strUserName As String           ' User name
    Dim bAllowReconnect As Boolean      ' Allow a reconnect?
    
    If Len(strMessage) > 0 Then
        m.dHeartBeat = gdTickCount
        
        If Not m.BrokerObj Is Nothing Then
            If m.BrokerObj.StatusFromInfo(strMessage, nStatus, strError, strUserName, bAllowReconnect) Then
                HandleConnectionInfo nStatus, strError, strUserName, bAllowReconnect
                If FormIsLoaded("frmBrokerView") Then
                    frmBrokerView.Broker_ConnectionStatus nStatus
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.StatusReceived"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountReceived
'' Description: Account message received from the broker
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AccountReceived(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim Acct As cPtAccount              ' Account object
    Dim rs As Recordset                 ' Recordset into the database
    Dim lPos As Long                    ' Position in an array
    Dim strFirstField As String         ' First field in the string
    Dim strFcmAccountNumber As String   ' FCM Account Number
    Static astrAccounts As cGdArray     ' Static array of accounts

    If astrAccounts Is Nothing Then
        Set astrAccounts = New cGdArray
        astrAccounts.Create eGDARRAY_Strings
    End If

    If Len(strMessage) > 0 Then
        m.dHeartBeat = gdTickCount
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If IsBroker Then
            If FormIsLoaded("frmBrokerView") Then
                frmBrokerView.Broker_Account strMessage
            End If
            If UCase(strFirstField) = "END" Then
                RefreshingAccounts = False
            End If
        Else
            If UCase(strFirstField) = "BEGIN" Then
                If Not m.BrokerObj Is Nothing Then
                    m.BrokerObj.AccountFromInfo strMessage, Nothing
                End If
                
                m.BrokerInfo.RefreshingAccounts = True
                m.BrokerInfo.ClearAccounts
                astrAccounts.Clear
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " account information...", , , m.strBrokerName & " Refresh", True
                End If
            ElseIf UCase(strFirstField) = "END" Then
                If Not m.BrokerObj Is Nothing Then
                    m.BrokerObj.AccountFromInfo strMessage, Nothing
                End If
                
                Set m.astrAccounts = astrAccounts.MakeCopy
                m.BrokerInfo.RefreshingAccounts = False
                RefreshingAccounts = False
                GetOrders
            Else
                Set Acct = New cPtAccount
                If Not m.BrokerObj Is Nothing Then
                    If m.BrokerObj.AccountFromInfo(strMessage, Acct) Then
                        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                                "WHERE [AccountNumber]='" & Acct.AccountNumber & "';", dbOpenDynaset)
                        If rs.BOF And rs.EOF Then
                            Acct.Save
                        Else
                            strFcmAccountNumber = Acct.FcmAccountNumber
                            
                            Acct.Load rs!AccountID
                            
                            ' DAJ 06/08/2011: If the account is found, but the account types don't match then if
                            ' both account types are Rithmic brokers, change the account type to the one logged
                            ' in right now...
                            ' DAJ 12/13/2011: If the account is found, but the account types don't match then if
                            ' both account types are RJ O'Brien brokers, change the account type to the one logged
                            ' in right now...
                            ' DAJ 07/19/2012: If the account is found, but the account types don't match then if
                            ' both account types are RCG (PATS) brokers, change the account type to the one logged
                            ' in right now...
                            ' DAJ 03/14/2013: If the account is found, but the account types don't match then if
                            ' both account types are CQG brokers, change the account type to the one logged
                            ' in right now...
                            If Acct.AccountType <> m.nBroker Then
                                If CanConvertBroker(Acct.AccountType) Then
                                    Acct.AccountType = m.nBroker
                                    Acct.Save
                                    g.Broker.UpdateAccountCache Acct
                                End If
                            End If
                            
                            ' DAJ 08/20/2012: If we received an FCM account number from the broker, but the
                            ' account in the database doesn't have it, set it now...
                            If (Len(Acct.FcmAccountNumber) = 0) And (Len(strFcmAccountNumber) > 0) Then
                                Acct.FcmAccountNumber = strFcmAccountNumber
                                Acct.Save
                            End If
                        End If
                                        
                        If astrAccounts.BinarySearch(Acct.AccountNumber, lPos) = False Then
                            astrAccounts.Add Acct.AccountNumber, lPos
                        End If
                        
                        RefreshAccount Acct.AccountID
                    
                        m.BrokerInfo.AddAccountObject Acct, eGDConnectionStatus_Connected
                        m.BrokerInfo.SetConnectionStatusForAccount eGDConnectionStatus_Connected, Acct.AccountNumber
                    End If
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.AccountReceived"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderReceived
'' Description: Order message received from the broker
'' Inputs:      Message, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderReceived(ByVal strMessage As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    
    If Len(strMessage) > 0 Then
        m.dHeartBeat = gdTickCount
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If IsBroker Then
            If UCase(strFirstField) = "END" Then
                RefreshingOrders = False
            End If
            If FormIsLoaded("frmBrokerView") Then
                frmBrokerView.Broker_Order strMessage, bRefresh
            End If
        Else
            If UCase(strFirstField) = "BEGIN" Then
                If Not m.BrokerObj Is Nothing Then
                    m.BrokerObj.OrderFromInfo "BEGIN", Nothing
                End If
                
                m.BrokerInfo.RefreshingOrders = True
                m.BrokerInfo.ClearOrders
                
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " order information...", , , m.strBrokerName & " Refresh", True
                End If
                SendBeginToOptionNav eGDOptNav_Order, m.nBroker
            ElseIf UCase(strFirstField) = "END" Then
                If Not m.BrokerObj Is Nothing Then
                    m.BrokerObj.OrderFromInfo "END", Nothing
                End If
                
                SendEndToOptionNav eGDOptNav_Order, m.nBroker
                
                m.BrokerInfo.RefreshingOrders = False
                RefreshingOrders = False
                GetFills
            Else
                HandleOrderMessage strMessage, bRefresh
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.OrderReceived"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillReceived
'' Description: Fill message received from the broker
'' Inputs:      Message, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FillReceived(ByVal strMessage As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    
    If Len(strMessage) > 0 Then
        m.dHeartBeat = gdTickCount
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If IsBroker Then
            If UCase(strFirstField) = "END" Then
                RefreshingFills = False
            End If
            If FormIsLoaded("frmBrokerView") Then
                frmBrokerView.Broker_Fill strMessage, bRefresh
            End If
        Else
            If UCase(strFirstField) = "BEGIN" Then
                If Not m.BrokerObj Is Nothing Then
                    m.BrokerObj.FillFromInfo "BEGIN", Nothing
                End If
                
                m.BrokerInfo.RefreshingFills = True
                m.BrokerInfo.ClearFills
    
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " fill information...", , , m.strBrokerName & " Refresh", True
                End If
                SendBeginToOptionNav eGDOptNav_Fill, m.nBroker
            ElseIf UCase(strFirstField) = "END" Then
                If Not m.BrokerObj Is Nothing Then
                    m.BrokerObj.FillFromInfo "END", Nothing
                End If
                
                SendEndToOptionNav eGDOptNav_Fill, m.nBroker
                
                m.BrokerInfo.RefreshingFills = False
                RefreshingFills = False
                GetPositions
            ElseIf Not m.BrokerObj Is Nothing Then
                HandleFillMessage strMessage, bRefresh
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.FillReceived"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CarriedFillReceived
'' Description: Carried fill message received from the broker
'' Inputs:      Message, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CarriedFillReceived(ByVal strMessage As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    
    If Len(strMessage) > 0 Then
        m.dHeartBeat = gdTickCount
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If IsBroker Then
            If FormIsLoaded("frmBrokerView") Then
                frmBrokerView.Broker_CarriedFill strMessage, bRefresh
            End If
        Else
            If UCase(strFirstField) = "BEGIN" Then
                If Not m.BrokerObj Is Nothing Then
                    m.BrokerObj.CarriedFillFromInfo "BEGIN", Nothing
                End If
                m.CarriedFills.Clear
    
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " fill information...", , , m.strBrokerName & " Refresh", True
                End If
            ElseIf UCase(strFirstField) = "END" Then
                If Not m.BrokerObj Is Nothing Then
                    m.BrokerObj.CarriedFillFromInfo "END", Nothing
                End If
            ElseIf Not m.BrokerObj Is Nothing Then
                HandleCarriedFillMessage strMessage, bRefresh
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.CarriedFillReceived"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionReceived
'' Description: Position message received from the broker
'' Inputs:      Message, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PositionReceived(ByVal strMessage As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    
    If Len(strMessage) > 0 Then
        m.dHeartBeat = gdTickCount
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If IsBroker Then
            If UCase(strFirstField) = "END" Then
                RefreshingPositions = False
            End If
            If FormIsLoaded("frmBrokerView") Then
                frmBrokerView.Broker_Position strMessage, bRefresh
            End If
        Else
            If UCase(strFirstField) = "BEGIN" Then
                If Not m.BrokerObj Is Nothing Then
                    m.BrokerObj.PositionFromInfo "BEGIN", Nothing
                End If
                
                m.BrokerInfo.RefreshingPositions = True
                m.BrokerInfo.ClearPositions
                
                ' DAJ 07/30/2012: We need to clear the broker positions collection here because if
                ' we didn't get anything back through a refresh for a symbol, we can end up using
                ' old information from this collection...
                m.BrokerPositions.Clear
                
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " position information...", , , m.strBrokerName & " Refresh", True
                End If
            ElseIf UCase(strFirstField) = "END" Then
                If Not m.BrokerObj Is Nothing Then
                    m.BrokerObj.PositionFromInfo "END", Nothing
                End If
                
                m.BrokerInfo.RefreshingPositions = False
                If m.bVerboseRefresh Then
                    InfBox "Synchronizing " & m.strBrokerName & " position information...", , , m.strBrokerName & " Refresh", True
                End If
                SynchronizePositions
                RefreshingPositions = False
                If m.bVerboseRefresh Then
                    InfBox ""
                    m.bVerboseRefresh = False
                End If
                If (m.astrAccounts.Size = 0) And (g.Broker.IsLiveAccount(m.nBroker) = False) Then
                    PositionVerify = False
                End If
            ElseIf Not m.BrokerObj Is Nothing Then
                'AddPositionToBrokerInfo strMessage, True
                HandlePositionMessage strMessage, bRefresh
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.PositionReceived"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrder
'' Description: Load the order for the given Genesis ID or Broker ID
'' Inputs:      Account Number or ID, Genesis Order ID, Broker ID, Old Order ID
'' Returns:     Order (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadOrder(ByVal vAccountNumberOrID As Variant, ByVal strGenesisOrderID As String, ByVal strBrokerID As String, Optional lOldOrderID As Long) As cPtOrder
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lAccountID As Long              ' Account ID for the given account number
    Dim Order As cPtOrder               ' Order to return
    Dim bFound As Boolean               ' Have we found the order?

    bFound = False
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    lOldOrderID = 0&
    
    ' First, try to look it up in the broker info in case it is already loaded...
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
                        "WHERE [BrokerOrderID]='" & strBrokerID & "' AND [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
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
                        "WHERE [GenesisOrderID]='" & strGenesisOrderID & "' AND [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
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
    RaiseError "cBroker.LoadOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerPrice
'' Description: Convert a Genesis price to the broker price
'' Inputs:      Genesis Price, Multiplier, Price Format
'' Returns:     Broker Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerPrice(ByVal dGenesisPrice As Double, ByVal dMult As Double, Optional ByVal strPriceFormat As String = "#") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from the function
    
    strReturn = ""
    If (dMult <> 0) And (dGenesisPrice <> 0) Then
        Select Case dMult
            Case 8
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 8, "0")
                
            Case 80
                strReturn = Str(Int(dGenesisPrice)) & Format((dGenesisPrice - Int(dGenesisPrice)) * 8, "0")
            
            Case 32
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 32, "00")
            
            Case 32.25
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 3200, "0000")
                strReturn = Left(strReturn, Len(strReturn) - 1)
            
            Case 32.5
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 320, "000")
                
            Case 325
                strReturn = Str(Int(dGenesisPrice)) & Format((dGenesisPrice - Int(dGenesisPrice)) * 320, "000")
                
            Case 64
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 64, "00")
                
            Case 6400
                strReturn = Str(Int(dGenesisPrice)) & Format((dGenesisPrice - Int(dGenesisPrice)) * 64, "00")
            
            Case 64.5
                strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 640, "000")
                
            Case 6450
                strReturn = Str(Int(dGenesisPrice)) & Format((dGenesisPrice - Int(dGenesisPrice)) * 640, "000")
                
            Case Else
                strReturn = Replace(Format(dGenesisPrice / dMult, strPriceFormat), ",", ".")

        End Select
    End If
    
    BrokerPrice = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.BrokerPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisPrice
'' Description: Convert a broker price to a Genesis price
'' Inputs:      Broker Price, Multiplier
'' Returns:     Genesis Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisPrice(ByVal strBrokerPrice As String, ByVal dMult As Double) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value from the function
    Dim strWhole As String              ' Whole portion of the price
    Dim strFraction As String           ' Fractional portion of the price
    
    If (Len(strBrokerPrice) = 0) Or (strBrokerPrice = "0") Then
        dReturn = 0#
    ElseIf (dMult <> 0) And (Len(strBrokerPrice) > 0) Then
        Select Case dMult
            ' In the case of 8ths, broker will send a price with the number of eighths after
            ' the decimal point, so 100.2=100 2/8, 100.4=100 4/8, etc.
            Case 8
                strWhole = Parse(strBrokerPrice, ".", 1)
                strFraction = Parse(strBrokerPrice, ".", 2)
                
                If Len(strFraction) = 0 Then
                    strFraction = "0"
                End If
                
                dReturn = Val(strWhole) + (Val(strFraction) / 8)
                
            ' In the case of 8ths, broker will send a price with the number of eighths as
            ' the last digit, so 1002=100 2/8, 1004=100 4/8, etc.
            Case 80
                strWhole = Left(strBrokerPrice, Len(strBrokerPrice) - 1)
                strFraction = Right(strBrokerPrice, 1)
                
                If Len(strFraction) = 0 Then
                    strFraction = "0"
                End If
                
                dReturn = Val(strWhole) + (Val(strFraction) / 8)
            
            ' In the case of 32nds, broker will send a price with the number of 32nds after
            ' the decimal point, so 100.20=100 20/32, 100.16=100 16/32, etc.
            Case 32
                strWhole = Parse(strBrokerPrice, ".", 1)
                strFraction = Parse(strBrokerPrice, ".", 2)
                
                If Len(strFraction) = 0 Then
                    strFraction = "00"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "0"
                End If
                
                dReturn = Val(strWhole) + (Val(strFraction) / 32)
            
            ' In the case of quarter 32nds, broker will send a price with the number of 32nds plus the
            ' fractional part of 32nds after the decimal point (with an implied 5 on the end if the
            ' broker price ends in a 2 or a 7), so 100.200=100 20.0/32, 100.162=100 16.25/32,
            ' 100.315=100 31.5/32, 100.307=100 30.75/32, etc.
            Case 32.25
                strWhole = Parse(strBrokerPrice, ".", 1)
                strFraction = Parse(strBrokerPrice, ".", 2)
                
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
            
            ' In the case of half 32nds, broker will send a price with the number of 32nds plus the
            ' fractional part of 32nds after the decimal point, so 100.200=100 20.0/32,
            ' 100.165=100 16.5/32, etc.
            Case 32.5
                strWhole = Parse(strBrokerPrice, ".", 1)
                strFraction = Parse(strBrokerPrice, ".", 2)
                
                If Len(strFraction) = 0 Then
                    strFraction = "000"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "00"
                ElseIf Len(strFraction) = 2 Then
                    strFraction = strFraction & "0"
                End If
                
                dReturn = Val(strWhole) + ((Val(strFraction) / 10) / 32)
                
            ' In the case of half 32nds, broker will send a price with the number of 32nds plus the
            ' fractional part of 32nds as the last three digits, so 100200=100 20.0/32,
            ' 100165=100 16.5/32, etc.
            Case 325
                strWhole = Left(strBrokerPrice, Len(strBrokerPrice) - 3)
                strFraction = Right(strBrokerPrice, 3)
                
                If Len(strFraction) = 0 Then
                    strFraction = "000"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "00"
                ElseIf Len(strFraction) = 2 Then
                    strFraction = strFraction & "0"
                End If
                
                dReturn = Val(strWhole) + ((Val(strFraction) / 10) / 32)
            
            ' In the case of 64ths, broker will send a price with the number of 64ths after
            ' the decimal point, so 100.20=100 20/64, 100.16=100 16/64, etc.
            Case 64
                strWhole = Parse(strBrokerPrice, ".", 1)
                strFraction = Parse(strBrokerPrice, ".", 2)
                
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
            
            ' In the case of 64ths, broker will send a price with the number of 64ths,
            ' so 10020=100 20/64, 10016=100 16/64, etc.
            Case 6400
                strWhole = Left(strBrokerPrice, Len(strBrokerPrice) - 2)
                strFraction = Right(strBrokerPrice, 2)
                
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
            
            ' In the case of half 64ths, broker will send a price with the number of 64ths plus the
            ' fractional part of 64ths after the decimal point, so 100.200=100 20.0/64,
            ' 100.165=100 16.5/64, etc.
            Case 64.5
                strWhole = Parse(strBrokerPrice, ".", 1)
                strFraction = Parse(strBrokerPrice, ".", 2)
                
                If Len(strFraction) = 0 Then
                    strFraction = "000"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "00"
                ElseIf Len(strFraction) = 2 Then
                    strFraction = strFraction & "0"
                End If
                
                dReturn = Val(strWhole) + ((Val(strFraction) / 10) / 64)
            
            ' In the case of half 64ths, broker will send a price with the number of 64ths,
            ' so 10020=100 20/64, 100165=100 16.5/64, etc.
            Case 6450
                strWhole = Left(strBrokerPrice, Len(strBrokerPrice) - 3)
                strFraction = Right(strBrokerPrice, 3)
                
                ' If we don't have a fractional part, default to zero...
                If Len(strFraction) = 0 Then
                    strFraction = "000"
                ElseIf Len(strFraction) = 1 Then
                    strFraction = strFraction & "00"
                ElseIf Len(strFraction) = 2 Then
                    strFraction = strFraction & "0"
                End If
                
                dReturn = Val(strWhole) + ((Val(strFraction) / 10) / 64)
            
            Case Else
                dReturn = Val(strBrokerPrice) * dMult
                
        End Select
    End If
    
    GenesisPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.GenesisPrice"
    
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

    Dim strReturn As String             ' Return value for the function
    
    If Not m.BrokerObj Is Nothing Then
        strReturn = m.BrokerObj.SymbolInformation(strGenesisSymbol)
    End If
    
    SymbolInformation = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.SymbolInformation"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateAccount
'' Description: Update the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateAccount(ByVal Acct As cPtAccount)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position of the account in the array

    If g.Broker.IsLiveAccount(m.nBroker) = False Then
        If m.astrAccounts.BinarySearch(Acct.AccountNumber, lPos) = False Then
            m.astrAccounts.Add Acct.AccountNumber, lPos
        End If
    End If

    m.BrokerInfo.AddAccountObject Acct, eGDConnectionStatus_Connected
    m.BrokerInfo.SetConnectionStatusForAccount eGDConnectionStatus_Connected, Acct.AccountNumber

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.UpdateAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartStandalone
'' Description: Start the stand-alone process for the broker
'' Inputs:      None
'' Returns:     True if started process, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StartStandalone() As Boolean
On Error GoTo ErrSection:
    
    Dim bReturn As Boolean              ' Return value for the function
    Dim bConnect As Boolean             ' Should we still try to connect?
    Dim strArgs As String               ' Arguments to send to the standalone
    
    bReturn = False
    If Len(m.strProcessName) > 0 Then
        If (frmOnlineBroker.gdBroker.FindControlID(m.strControlID) = 0) Or (StandaloneIsRunning = False) Then
            If m.bIsDotNet Then
                bConnect = HasDotNet
            Else
                bConnect = True
            End If
            
            If bConnect Then
                If StopStandalone = True Then
                    Sleep 3
                End If
                
                DumpDebug "Loading application '" & m.strProcessPath & "'"
                RunProcess m.strProcessPath, strArgs, , vbHide
                StartedProcess = gdTickCount
                
                bReturn = True
            End If
        End If
    End If
    
    StartStandalone = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.StartStandalone"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StopStandalone
'' Description: Stop the stand-alone process for the broker
'' Inputs:      None
'' Returns:     True if killed process, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StopStandalone() As Boolean
On Error GoTo ErrSection:

    StopStandalone = (KillProcess(m.strProcessName, False) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.StopStandalone"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StandaloneIsRunning
'' Description: Determine if the stand-alone process for the broker is running
'' Inputs:      None
'' Returns:     True if running, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StandaloneIsRunning() As Boolean
On Error GoTo ErrSection:

    StandaloneIsRunning = (KillProcess(m.strProcessName, True) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.StandaloneIsRunning"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFill
'' Description: Attempt to load the fill with the given information
'' Inputs:      Account, Broker ID
'' Returns:     Fill (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadFill(ByVal vAccountNumberOrID As Variant, ByVal strBrokerFillID As String) As cPtFill
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID
    Dim ReturnFill As cPtFill           ' Fill to return from the function
    Dim rs As Recordset                 ' Recordset into the database
    
    Set ReturnFill = m.BrokerInfo.Fills(strBrokerFillID)
    If ReturnFill Is Nothing Then
        lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
        
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                    "WHERE [BrokerFillID]='" & strBrokerFillID & "' AND [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            Set ReturnFill = New cPtFill
            If ReturnFill.Load(rs!FillID, rs) = False Then
                Set ReturnFill = Nothing
            End If
        End If
    End If
        
    Set LoadFill = ReturnFill
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.LoadFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SynchronizePositions
'' Description: Synchronize the positions between Genesis and the broker
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
    
    If IsBroker = False Then
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
                        If lBrokerCarried = kNullData Then
                            lBrokerCarried = lBrokerPos - .NumBuysSnapshot + .NumSellsSnapshot
                        End If
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
                                    DumpDebug "Symbol removed from To Fix list ( SynchronizePositions ): " & strAccount & ", " & .Symbol
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
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.SynchronizePositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddBrokerRtSymbols
'' Description: Add the list of symbols from the broker to the RT overrides
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddBrokerRtSymbols()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of symbols
    Dim lIndex As Long                  ' Index into a for loop
    Dim lLastDate As Long               ' Last successful connection to broker
    
    If g.Broker.IsBrokerUser(m.nBroker) Then
        lLastDate = LastDateConnected
        
        If g.Broker.IsBrokerSimUser(m.nBroker) Or (lLastDate >= Date - 30) Then
            If Not m.BrokerObj Is Nothing Then
                Set astrSymbols = m.BrokerObj.GenesisSymbolList
            End If
            
            DumpDebug "Symbols enabled for real-time: '" & astrSymbols.JoinFields(";") & "'"
            For lIndex = 0 To astrSymbols.Size - 1
                g.RealTime.AddBrokerRtSymbol astrSymbols(lIndex) & "-", m.strBrokerName & "/D"
            Next lIndex
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.AddBrokerRtSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnabledSymbols
'' Description: Get the list of symbols the user is enabled to trade
'' Inputs:      None
'' Returns:     Enabled Symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnabledSymbols() As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Return value for the function
    
    If m.BrokerObj Is Nothing Then
        Set astrReturn = Nothing
    Else
        Set astrReturn = m.BrokerObj.EnabledSymbols
    End If
    
    Set EnabledSymbols = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.EnabledSymbols"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimeZone
'' Description: Determine the time zone this broker sends time in for the symbol
'' Inputs:      Symbol
'' Returns:     Time Zone
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TimeZone(ByVal strSymbol As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Not m.BrokerObj Is Nothing Then
        strReturn = m.BrokerObj.TimeZone(strSymbol)
    End If
    
    TimeZone = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.TimeZone"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Send a string to the log file for the day
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DumpDebug(ByVal strMessage As String)
On Error Resume Next

#If 0 Then

    Dim fh As Integer                   ' File handle to open file with
    fh = FreeFile
    Open AddSlash(m.strLogPath) & "TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ") - " & strMessage
        Close #fh
    End If

#Else

    Static LogFile As cLogFile
    If LogFile Is Nothing Then
        Set LogFile = New cLogFile
        LogFile.OpenFile AddSlash(m.strLogPath) & "TN*.LOG"
    End If
    LogFile.WriteText strMessage
    
#End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendBrokerMessage
'' Description: Send an App Mail message to the stand-alone program
'' Inputs:      Type of Message, Message, Send Now?, Log String
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendBrokerMessage(ByVal nMessage As eGDBrokerMessageTypes, ByVal strMessage As String, Optional ByVal bSendNow As Boolean = True, Optional ByVal strLogString = "")
On Error GoTo ErrSection:

    Dim astrMessage As New cGdArray     ' Message split out into an array

    If Len(strLogString) > 0 Then
        DumpDebug vbTab & "Sending " & m.strBrokerName & " Message (" & BrokerMessageTypeToString(nMessage) & "): " & strLogString
    Else
        DumpDebug vbTab & "Sending " & m.strBrokerName & " Message (" & BrokerMessageTypeToString(nMessage) & "): " & strMessage
    End If
    
    SendMessage nMessage, strMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.SendBrokerMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming App Mail message from the stand-alone program
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDBrokerMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    If (nType <> eGDBrokerMessageType_ConnectionInfo) And (nType <> eGDBrokerMessageType_Heartbeat) And (nType <> eGDBrokerMessageType_PriceUpdate) Then
        DumpDebug vbTab & m.strBrokerName & " Message Received (" & BrokerMessageTypeToString(nType) & "): " & strMessage
    End If
    
    Select Case nType
        Case eGDBrokerMessageType_ConnectionInfo
            StatusReceived strMessage
            
        Case eGDBrokerMessageType_AppLoaded
            AppLoaded = True
            
        Case eGDBrokerMessageType_AppUnloaded
            AppLoaded = False
            
        Case eGDBrokerMessageType_Heartbeat
            m.dHeartBeat = gdTickCount
            
        Case eGDBrokerMessageType_AccountRefresh
            AccountReceived strMessage
            
        Case eGDBrokerMessageType_Order
            OrderReceived strMessage, False
        
        Case eGDBrokerMessageType_OrderRefresh
            OrderReceived strMessage, True
            
        Case eGDBrokerMessageType_Fill
            FillReceived strMessage, False
        
        Case eGDBrokerMessageType_FillRefresh
            FillReceived strMessage, True
            
        Case eGDBrokerMessageType_Position
            PositionReceived strMessage, False
        
        Case eGDBrokerMessageType_PositionRefresh
            PositionReceived strMessage, True
            
        Case eGDBrokerMessageType_CarriedFillRefresh
            CarriedFillReceived strMessage, True
            
        Case eGDBrokerMessageType_NumberOfAccounts
            NumberOfAccountsReceived strMessage
            
        Case eGDBrokerMessageType_SpreadFill
            HandleSpreadFillMessage strMessage, False
            
        Case Else
            If Not m.BrokerObj Is Nothing Then
                m.BrokerObj.HandleMessage nType, strMessage
            End If
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.HandleMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleOrderMessage
'' Description: Handle an order message received from the broker
'' Inputs:      Message, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleOrderMessage(ByVal strMessage As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim brokerOrder As cPtOrder         ' Order filled with information from the message
    Dim existingOrder As cPtOrder       ' Existing order if it exists
    Dim Order As cPtOrder               ' Order to fill in
    Dim PrevOrder As cPtOrder           ' Previous order
    Dim lOldOrderID As Long             ' Previous order ID for this order
    Dim bNewOrder As Boolean            ' Is this a new order?
    Dim nPrevStatus As eTT_OrderStatus  ' Previous order status
    Dim strPrevBrokerID As String       ' Previous broker ID
    Dim frm As frmAlertPopup            ' Alert pop-up form
    Dim OtherOrder As cPtOrder          ' Other order in a broker OCO situation
    Dim lPos As Long                    ' Position in the array
    Dim strRejectMessage As String      ' Message to display to the user
    Dim bShowMessage As Boolean         ' Show the message to the user?
    Dim strPrevGenesisID As String      ' Previous Genesis ID

    If Not m.BrokerObj Is Nothing Then
        If m.BrokerObj.OrderFromInfo(strMessage, brokerOrder) Then
            If brokerOrder.AccountID > 0 Then
                Set existingOrder = LoadOrder(brokerOrder.AccountID, brokerOrder.GenesisOrderID, brokerOrder.BrokerID, lOldOrderID)
                If existingOrder Is Nothing Then
                    DumpDebug "Order with BrokerID = '" & brokerOrder.BrokerID & "' and GenesisID = '" & brokerOrder.GenesisOrderID & "' does not exist"
                    Set Order = New cPtOrder
                    bNewOrder = True
                    
                    If (lOldOrderID <= 0&) And (brokerOrder.PreviousOrderID > 0&) Then
                        lOldOrderID = brokerOrder.PreviousOrderID
                        DumpDebug vbTab & "Old Order ID set to " & Str(lOldOrderID)
                    End If
                Else
                    DumpDebug "Order with BrokerID = '" & brokerOrder.BrokerID & "' and GenesisID = '" & brokerOrder.GenesisOrderID & "' exists"
                    Set Order = existingOrder.MakeCopy
                    bNewOrder = False
                    
                    If bRefresh = False Then
                        If brokerOrder.Status = eTT_OrderStatus_AmendPending Then
                            If g.Broker.IsCqgBroker(m.nBroker) = True Then
                                If m.BrokerObj.AmendOrders.Exists(Order.GenesisOrderID) = False Then
                                    DumpDebug vbTab & "Order added to amend orders array: '" & Order.OrderText(True, True, True) & "'"
                                    m.BrokerObj.AmendOrders.Add Order, Order.GenesisOrderID
                                End If
                            End If
                        End If
                    End If
                End If
                
                With Order
                    nPrevStatus = .Status
                    
                    If .OrderID = 0 Then
                        .GenesisOrderID = brokerOrder.GenesisOrderID
                    ElseIf (g.Broker.IsPatsBroker(m.nBroker) = True) And (Not brokerOrder.PreviousOrder Is Nothing) Then
                        If .GenesisOrderID = brokerOrder.PreviousOrder.GenesisOrderID Then
                            .GenesisOrderID = brokerOrder.GenesisOrderID
                        End If
                    End If
                    
                    .BrokerID = brokerOrder.BrokerID
                    .ExchangeID = brokerOrder.ExchangeID
                    .AccountID = brokerOrder.AccountID
                    '.Buy = brokerOrder.Buy
                    If brokerOrder.Quantity > 0 Then
                        .Quantity = brokerOrder.Quantity
                    End If
                    .SessionDate = brokerOrder.SessionDate
                    '.SymbolOrSymbolID = brokerOrder.SymbolOrSymbolID
                    
                    ' 05/20/2013 DAJ: If this is an existing single-leg order, just set the
                    ' buy and symbol on the leg here ( so we don't overwrite the enter/exit
                    ' flag ), otherwise just make a copy of the order legs for now...
                    If (bNewOrder = False) And (.NumberOfLegs = 1) And (brokerOrder.NumberOfLegs = 1) Then
                        .Buy = brokerOrder.Buy
                        .SymbolOrSymbolID = brokerOrder.SymbolOrSymbolID
                    Else
                        .OrderLegs = brokerOrder.OrderLegs.MakeCopy
                    End If
                    
                    ' 09/02/2015 DAJ: While all brokers turn a stop order into a limit order behind the scenes, Rithmic
                    ' seems to be the only one sending it back to us that way.  This is causing an issue with contingency
                    ' orders because the limit price can be far away from the market and we end up moving the contingency
                    ' orders based on that new order price, so we won't let that happen here...
                    If ((.OrderType = eTT_OrderType_Stop) And (brokerOrder.OrderType = eTT_OrderType_Limit)) And (g.Broker.IsRithmicBroker(m.nBroker) = True) Then
                        DumpDebug vbTab & "Order type not changed from a Stop to a Limit"
                    Else
                        .OrderType = brokerOrder.OrderType
                        .LimitPrice = brokerOrder.LimitPrice
                        .StopPrice = brokerOrder.StopPrice
                    End If
                    
                    If .OrderDate = 0# Then
                        .OrderDate = brokerOrder.OrderDate
                    End If
                    .SessionDate = brokerOrder.SessionDate
                    .IsSnapshot = True
                    .Expiration = brokerOrder.Expiration
                    
                    ' If there is no order ID yet, then we need to save so that we have an order ID
                    ' to send to the order ID changed function (if in an amend situation)...
                    If .OrderID = 0& Then
                        .Save
                    End If
                    
                    strPrevBrokerID = ""
                    Set PrevOrder = Nothing
                    If lOldOrderID <> 0& Then
                        Set PrevOrder = New cPtOrder
                        If PrevOrder.Load(lOldOrderID) = False Then
                            Set PrevOrder = Nothing
                        End If
                    ElseIf (brokerOrder.Status = eTT_OrderStatus_Working) And (m.GidMap.Exists(.GenesisOrderID)) Then
                        strPrevGenesisID = m.GidMap(.GenesisOrderID)
                        
                        If m.Originals.Exists(strPrevGenesisID) Then
                            lOldOrderID = m.Originals(strPrevGenesisID).OrderID
                            
                            Set PrevOrder = New cPtOrder
                            If PrevOrder.Load(lOldOrderID) = False Then
                                Set PrevOrder = Nothing
                            End If
                        End If
                    End If
                        
                    If Not PrevOrder Is Nothing Then
                        ' Copy information from the old order to the new order...
                        strPrevBrokerID = PrevOrder.BrokerID
                        .AutoTradeItemID = PrevOrder.AutoTradeItemID
                        .BrokerCancelOrderID = PrevOrder.BrokerCancelOrderID
                        .CancelOrderID = PrevOrder.CancelOrderID
                        .TriggerOrderID = PrevOrder.TriggerOrderID
                        .TriggerOptions = PrevOrder.TriggerOptions
                        .ConditionOptions = PrevOrder.ConditionOptions
                        .TrailAmount = PrevOrder.TrailAmount
                        .TrailOptions = PrevOrder.TrailOptions
                        .ExitPos = PrevOrder.ExitPos
                        .IsAutoExit = PrevOrder.IsAutoExit
                        .GroupID = PrevOrder.GroupID
                        .GroupName = PrevOrder.GroupName
                        .PreviousOrder = PrevOrder
                    
                        ' Clean out information for the old order and refresh...
                        PrevOrder.GenesisOrderID = ""
                        PrevOrder.BrokerCancelOrderID = 0&
                        PrevOrder.CancelOrderID = 0&
                        PrevOrder.Save
                        
                        m.BrokerInfo.AddOrder PrevOrder, bRefresh
                        OrderCallback PrevOrder
                            
                        DumpDebug "Order ID for '" & .GenesisOrderID & "' changed from " & Str(lOldOrderID) & " to " & Str(.OrderID)
                        OrderIDChanged lOldOrderID, .OrderID
                        SendOrderIdChangeToOptNav .AccountID, Str(lOldOrderID), Str(.OrderID)
                    End If
                    
                    ' 05/17/2013 DAJ: This line right here is resetting the Entry flag on
                    ' each of the legs.  In the case of automated trading, it is causing the
                    ' code to think that the order is now a Short Exit instead of a Long Entry
                    ' and vice versa ( see new block above )...
                    '.OrderLegs = brokerOrder.OrderLegs.MakeCopy
                    
                    .Status = brokerOrder.Status
                    .StatusDate = brokerOrder.StatusDate
                    
                    If (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_ParkPending) Then
                        .Status = eTT_OrderStatus_Parked
                    ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_Parked) And (bNewOrder = False) Then
                        .Status = eTT_OrderStatus_Parked
                    ElseIf (.Status = eTT_OrderStatus_CancelPending) And (nPrevStatus = eTT_OrderStatus_ParkPending) Then
                        .Status = eTT_OrderStatus_ParkPending
                    ElseIf (Order.FillQuantity > Order.Quantity) Then
                        .Status = eTT_OrderStatus_OverFilled
                    ElseIf m.WaitReplace.Exists(.BrokerID) Then
                        If (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_AmendPending) Then
                            .Status = eTT_OrderStatus_Amended
                        ElseIf (.Status = eTT_OrderStatus_CancelPending) And (nPrevStatus = eTT_OrderStatus_AmendPending) Then
                            .Status = eTT_OrderStatus_AmendPending
                        End If
                    ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_Amended) Then
                        .Status = eTT_OrderStatus_Amended
                    End If
                    
                    ' 02/18/2015 DAJ: For Rithmic accounts, if the order goes from 'Cancel Pending' to 'Working',
                    ' that probably means that the cancel failed.  In this case, we will want to stop the
                    ' flatten ( or cancel all ) process if that is what caused the cancel in the first place
                    ' because Rithmic doesn't want people continually trying to cancel an order...
                    If g.Broker.IsRithmicBroker(.Broker) Then
                        If (.Status = eTT_OrderStatus_Working) And (nPrevStatus = eTT_OrderStatus_CancelPending) Then
                            If Not g.FlattenQueue Is Nothing Then
                                g.FlattenQueue.Remove g.Broker.AccountNumberForID(.AccountID), .Symbol, .AutoTradeItemID, "an order could not be cancelled"
                            End If
                        End If
                    End If
                    
                    ' 01/14/2013 DAJ: If the order is part of an unconfirmed broker-held OCO AND it is in a
                    ' working status AND the other order is in a working status then confirm the broker-held
                    ' OCO...
                    If (.Status = eTT_OrderStatus_Working) Or (.Status = eTT_OrderStatus_Partial) Or (.Status = eTT_OrderStatus_PreSubmitted) Then
                        If .BrokerCancelOrderID < 0 Then
                            Set OtherOrder = New cPtOrder
                            If OtherOrder.Load(Abs(.BrokerCancelOrderID)) Then
                                If (OtherOrder.Status = eTT_OrderStatus_Working) Or (OtherOrder.Status = eTT_OrderStatus_Partial) Or (OtherOrder.Status = eTT_OrderStatus_PreSubmitted) Then
                                    .BrokerCancelOrderID = Abs(.BrokerCancelOrderID)
                                    
                                    If OtherOrder.BrokerCancelOrderID < 0 Then
                                        OtherOrder.BrokerCancelOrderID = Abs(OtherOrder.BrokerCancelOrderID)
                                        OtherOrder.Save
                                        
                                        OrderCallback OtherOrder
                                    End If
                                End If
                            End If
                        End If
                    
                    ' 01/14/2013 DAJ: If the order is part of a confirmed broker-held OCO AND it is no longer
                    ' in a working status AND the other order is in a working status then clear the broker-held
                    ' OCO...
                    ElseIf IsOpenOrder(.Status) = False Then
                        If .BrokerCancelOrderID > 0 Then
                            Set OtherOrder = New cPtOrder
                            If OtherOrder.Load(.BrokerCancelOrderID) Then
                                If (OtherOrder.Status = eTT_OrderStatus_Working) Or (OtherOrder.Status = eTT_OrderStatus_Partial) Or (OtherOrder.Status = eTT_OrderStatus_PreSubmitted) Then
                                    .BrokerCancelOrderID = 0
                                    
                                    If OtherOrder.BrokerCancelOrderID > 0 Then
                                        OtherOrder.BrokerCancelOrderID = 0
                                        OtherOrder.Save
                                        
                                        OrderCallback OtherOrder
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    .Message = brokerOrder.Message
                    If Len(strPrevBrokerID) > 0 Then
                        .PreviousBrokerID = strPrevBrokerID
                    ElseIf (Len(brokerOrder.PreviousBrokerID) > 0) And (brokerOrder.PreviousBrokerID <> brokerOrder.BrokerID) Then
                        .PreviousBrokerID = brokerOrder.PreviousBrokerID
                    End If
                    
                    .Save
                
                    m.BrokerInfo.AddOrder Order, bRefresh
                    OrderCallback Order
                    
                    If bRefresh = False Then
                        g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                    End If
                    
                    SendOrderToOptionNav Order, bRefresh
                    
                    ' 03/11/2016 DAJ: Only show rejected order popup messages if it is a manual order...
                    If (Len(brokerOrder.Message) > 0) And (bRefresh = False) And (Order.AutoTradeItemID = 0) Then
                        strRejectMessage = Order.OrderText & " REJECTED||" & brokerOrder.Message
                        
                        If m.RejectShown.Exists(Str(.OrderID)) Then
                            bShowMessage = strRejectMessage <> m.RejectShown(Str(.OrderID))
                        Else
                            bShowMessage = True
                        End If
                        
                        If bShowMessage = True Then
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox strRejectMessage, m.strBrokerName & " Order " & Order.BrokerID & " Rejected", vbLeftJustify
                            m.RejectShown.Add strRejectMessage, Str(.OrderID)
                            
                            DumpDebug "Message shown to user: '" & strRejectMessage & "'"
                        Else
                            DumpDebug "Message not shown to user because we already showed it: '" & strRejectMessage & "'"
                        End If
                    End If
                    
                    DumpDebug Order.OrderText(True, True, True) & ": Status = '" & OrderStatus(Order.Status) & "'; Session = " & DateFormat(Order.SessionDate, MM_DD_YYYY)
                    
                    If (.Status = eTT_OrderStatus_Amended) And (m.WaitReplace.Exists(.BrokerID) = True) Then
                        Set OtherOrder = m.WaitReplace(.BrokerID).MakeCopy
                        Set PrevOrder = OtherOrder.MakeCopy
                        
                        OtherOrder.BrokerID = ""
                        OtherOrder.PreviousBrokerID = .BrokerID
                        
                        strPrevGenesisID = OtherOrder.GenesisOrderID
                        OtherOrder.GenesisOrderID = NextGenesisID(g.Broker.AccountNumberForID(OtherOrder.AccountID))
                        If m.GidMap.Exists(OtherOrder.GenesisOrderID) Then
                            DumpDebug vbTab & "GidMap Entry changed for '" & OtherOrder.GenesisOrderID & "' from '" & m.GidMap(OtherOrder.GenesisOrderID) & "' to '" & strPrevGenesisID & "'"
                            m.GidMap(OtherOrder.GenesisOrderID) = strPrevGenesisID
                        Else
                            DumpDebug vbTab & "GidMap Entry added for '" & OtherOrder.GenesisOrderID & "': '" & strPrevGenesisID & "'"
                            m.GidMap.Add strPrevGenesisID, OtherOrder.GenesisOrderID
                        End If
                        
                        DumpDebug "AddOrder Because of Cancel/Replace: " & OtherOrder.OrderText(True, True, True)
                        AddOrder OtherOrder
                        
                        DumpDebug vbTab & "Order removed from WaitReplace: " & PrevOrder.OrderText(True, True, True)
                        m.WaitReplace.Remove .BrokerID
                    
                    ElseIf (.Status = eTT_OrderStatus_Working) And (m.GidMap.Exists(.GenesisOrderID)) Then
                        strPrevGenesisID = m.GidMap(.GenesisOrderID)
                        
                        If m.Originals.Exists(strPrevGenesisID) Then
                            DumpDebug vbTab & "Order removed from Originals: " & m.Originals(strPrevGenesisID).OrderText(True, True, True)
                            m.Originals.Remove strPrevGenesisID
                        End If
                        
                        DumpDebug vbTab & "GidMap Entry Removed for '" & .GenesisOrderID & "': '" & strPrevGenesisID & "'"
                        m.GidMap.Remove .GenesisOrderID
                    
                    ElseIf (.Status = eTT_OrderStatus_Rejected) And (m.GidMap.Exists(.GenesisOrderID)) Then
                        strPrevGenesisID = m.GidMap(.GenesisOrderID)
                        
                        If m.Originals.Exists(strPrevGenesisID) Then
                            Set OtherOrder = m.Originals(strPrevGenesisID).MakeCopy
                            
                            OtherOrder.BrokerID = ""
                            OtherOrder.PreviousBrokerID = ""
                            OtherOrder.GenesisOrderID = NextGenesisID(g.Broker.AccountNumberForID(OtherOrder.AccountID))
                        
                            DumpDebug "AddOrder Because of Cancel/Replace Reject: " & OtherOrder.OrderText(True, True, True)
                            AddOrder OtherOrder
                            
                            DumpDebug vbTab & "Order removed from Originals: " & m.Originals(strPrevGenesisID).OrderText(True, True, True)
                            m.Originals.Remove strPrevGenesisID
                        End If
                        
                        DumpDebug vbTab & "GidMap Entry Removed for '" & .GenesisOrderID & "': '" & strPrevGenesisID & "'"
                        m.GidMap.Remove .GenesisOrderID
                    End If
                End With
            Else
                DumpDebug "Order not processed because account could not be converted"
            End If
        Else
            DumpDebug "Order not processed because symbol could not be converted"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.HandleOrderMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleFillMessage
'' Description: Handle a fill message received from the broker
'' Inputs:      Message, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleFillMessage(ByVal strMessage As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim brokerFill As cPtFill           ' Fill object loaded with information from the message
    Dim strGenesisOrderID As String     ' Genesis ID for the order the fill belongs to
    Dim Fill As cPtFill                 ' Fill to set properties on and save
    Dim FillOrder As cPtOrder           ' Order that the fill belongs to
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFillExisted As Boolean         ' Did the fill already exist?
    Dim bFillChanged As Boolean         ' Did the fill change?
    Dim lFillQuantity As Long           ' Fill quantity
    
    If Not m.BrokerObj Is Nothing Then
        If m.BrokerObj.FillFromInfo(strMessage, brokerFill, strGenesisOrderID) Then
            If brokerFill.AccountID > 0 Then
                Set FillOrder = LoadOrder(brokerFill.AccountID, strGenesisOrderID, brokerFill.BrokerOrderID)
                If Not FillOrder Is Nothing Then
                    bFillExisted = False
                    bFillChanged = False
                    
                    ' When CQG sends fills back for a spread order in a refresh, they send back the
                    ' main 'ExecID' for the order ( which means they will be duplicate for each leg
                    ' of the spread ).  In order to make those unique, append the symbol for the leg
                    ' to the 'ExecID'.  We will also do this for unsolicited fills so as to be
                    ' consistent ( and so that we can find it later in a refresh )...
                    If (IsSpreadSymbol(FillOrder.Symbol) = True) And (g.Broker.IsCqgBroker(brokerFill.Broker) = True) Then
                        brokerFill.BrokerID = brokerFill.BrokerID & "_" & brokerFill.Symbol
                    End If
                    
                    For lIndex = 1 To FillOrder.Fills.Count
                        If (FillOrder.Fills(lIndex).BrokerID = brokerFill.BrokerID) Or ((Len(brokerFill.PreviousBrokerID) > 0) And (FillOrder.Fills(lIndex).BrokerID = brokerFill.PreviousBrokerID)) Then
                            Set Fill = FillOrder.Fills(lIndex)
                            
                            bFillExisted = True
                            bFillChanged = (Fill.FillDate <> brokerFill.FillDate) Or (Fill.Price <> brokerFill.Price) Or (Fill.Quantity <> brokerFill.Quantity)
                            
                            Exit For
                        End If
                    Next lIndex
                    
                    If (Not (Fill Is Nothing)) And (brokerFill.Quantity = 0) Then
                        Fill.Delete "Busted Fill"
                        If bRefresh = False Then
                            Refresh
                        End If
                    Else
                        If Fill Is Nothing Then
                            Set Fill = New cPtFill
                        End If
                        
                        With Fill
                            .AccountID = FillOrder.AccountID
                            ' 04/08/2013 DAJ: Only assign the automated trading item ID if
                            ' this is a new fill -- the user could have reassigned the fill
                            ' which does not reassign the order...
                            If .FillID = 0& Then
                                .AutoTradingItemID = FillOrder.AutoTradeItemID
                            End If
                            .BrokerID = brokerFill.BrokerID
                            .PreviousBrokerID = brokerFill.PreviousBrokerID
                            .BrokerOrderID = FillOrder.BrokerID
                            .Buy = brokerFill.Buy ' FillOrder.Buy
                            If .FillDate = 0# Then
                                .FillDate = brokerFill.FillDate
                            End If
                            .IsManual = False
                            .IsSnapshot = True
                            .IsSpread = False
                            .OrderID = FillOrder.OrderID
                            .Price = brokerFill.Price
                            .Quantity = brokerFill.Quantity
                            .SessionDate = brokerFill.SessionDate
                            If Len(brokerFill.Symbol) > 0 Then
                                .SymbolOrSymbolID = brokerFill.SymbolOrSymbolID
                            Else
                                .SymbolOrSymbolID = FillOrder.SymbolOrSymbolID
                            End If
                            
                            ' If we don't have a Fill ID yet (this is a new fill), we need to save
                            ' the fill to get a Fill ID for the key into the Fills collection of
                            ' the order...
                            If .FillID = 0& Then
                                .Save
                            End If
                        End With
                            
                        With FillOrder
                            Set .Fills(Str(Fill.FillID)) = Fill
                            .IsSnapshot = True
                            
                            lFillQuantity = .FillQuantity
                            
                            ' Fix the status on the order if necessary...
                            If (lFillQuantity = .Quantity) Then
                                .Status = eTT_OrderStatus_Filled
                            ElseIf (lFillQuantity > .Quantity) Then
                                .Status = eTT_OrderStatus_OverFilled
                            ElseIf ((.Status = eTT_OrderStatus_Cancelled) Or (.Status = eTT_OrderStatus_Expired)) And (lFillQuantity > 0) Then
                                .Status = eTT_OrderStatus_BalCancelled
                            ElseIf ((.Status = eTT_OrderStatus_Working) Or (.Status = eTT_OrderStatus_Filled)) And (lFillQuantity > 0) Then
                                .Status = eTT_OrderStatus_Partial
                            End If
                            
                            .Save
                        End With
                        
                        m.BrokerInfo.AddOrder FillOrder, bRefresh
                        m.BrokerInfo.AddFill Fill, bRefresh
                        
                        FillCallback FillOrder, Fill, bFillExisted, False
                        If bRefresh = False Then
                            g.ActivityLogs.AddFillToActivityLog m.nBroker, Fill, False
                            m.dLastFillTime = gdTickCount
                        End If
                        DumpDebug mTradeTracker.FillDisplay(Fill)
                        
                        SendFillToOptionNav Fill, bRefresh
                        
                        If bRefresh = False Then
                            If g.Broker.IsCqgBroker(m.nBroker) Then
                                ' If we get an unsolicited fill from CQG, generate an unsolicited
                                ' position message as well...
                                m.BrokerObj.GetPositionForFill strMessage
                            End If
                        End If
                    End If
                ElseIf brokerFill.Quantity > 0 Then
                    DumpDebug "Order with BrokerID = '" & brokerFill.BrokerOrderID & "' does not exist"
                    Set Fill = LoadFill(brokerFill.AccountID, brokerFill.BrokerID)
                    If (Not (Fill Is Nothing)) And (brokerFill.Quantity = 0) Then
                        Fill.Delete "Busted Fill"
                        If bRefresh = False Then
                            Refresh
                        End If
                    Else
                        If Fill Is Nothing Then
                            DumpDebug "Fill with BrokerID = '" & brokerFill.BrokerID & "' does not exist"
                            Set Fill = New cPtFill
                            bFillExisted = False
                        Else
                            DumpDebug "Fill with BrokerID = '" & brokerFill.BrokerID & "' exists"
                            bFillExisted = True
                        End If
                        
                        bFillChanged = (Fill.FillDate <> brokerFill.FillDate) Or (Fill.Price <> brokerFill.Price) Or (Fill.Quantity <> brokerFill.Quantity)
                        
                        With Fill
                            .AccountID = brokerFill.AccountID
                            ' 04/08/2013 DAJ: Only assign the automated trading item ID if
                            ' this is a new fill ( keep if existing fill )...
                            If .FillID = 0& Then
                                .AutoTradingItemID = 0
                            End If
                            .BrokerID = brokerFill.BrokerID
                            .BrokerOrderID = brokerFill.BrokerOrderID
                            .Buy = brokerFill.Buy
                            If .FillDate = 0# Then
                                .FillDate = brokerFill.FillDate
                            End If
                            .IsManual = False
                            .IsSnapshot = True
                            .IsSpread = False
                            .OrderID = 0
                            .Price = brokerFill.Price
                            .Quantity = brokerFill.Quantity
                            .SessionDate = brokerFill.SessionDate
                            .SymbolOrSymbolID = brokerFill.SymbolOrSymbolID
                            
                            ' If we don't have a Fill ID yet (this is a new fill), we need to save
                            ' the fill to get a Fill ID for the key into the Fills collection of
                            ' the order...
                            ' DAJ 11/10/2014: We need to save this no matter what here in case something
                            ' about the fill changed ( because nothing else saves it after this point ) ...
                            'If .FillID = 0& Then
                                .Save
                            'End If
                        End With
                            
                        m.BrokerInfo.AddFill Fill, bRefresh
                        
                        If bRefresh = False Then
                            g.ActivityLogs.AddFillToActivityLog m.nBroker, Fill, False
                            m.dLastFillTime = gdTickCount
                        End If
                        DumpDebug mTradeTracker.FillDisplay(Fill)
                        
                        SendFillToOptionNav Fill, bRefresh
                        
                        If bRefresh = False Then
                            If g.Broker.IsCqgBroker(m.nBroker) Then
                                ' If we get an unsolicited fill from CQG, generate an unsolicited
                                ' position message as well...
                                m.BrokerObj.GetPositionForFill strMessage
                            End If
                        End If
                    End If
                Else
                    DumpDebug "Fill not processed because of a zero quantity"
                    If bRefresh = False Then
                        Refresh
                    End If
                End If
            Else
                DumpDebug "Fill not processed because account could not be converted"
            End If
        Else
            DumpDebug "Fill not processed because symbol could not be converted"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.HandleFillMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleSpreadFillMessage
'' Description: Handle a spread fill message received from the broker
'' Inputs:      Message, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleSpreadFillMessage(ByVal strMessage As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim brokerFill As cPtFill           ' Fill object loaded with information from the message
    Dim strGenesisOrderID As String     ' Genesis ID for the order the fill belongs to
    Dim Fill As cPtFill                 ' Fill to set properties on and save
    Dim FillOrder As cPtOrder           ' Order that the fill belongs to
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFillExisted As Boolean         ' Did the fill already exist?
    Dim bFillChanged As Boolean         ' Did the fill change?
    Dim lFillQuantity As Long           ' Fill quantity
    
    If Not m.BrokerObj Is Nothing Then
        If m.BrokerObj.FillFromInfo(strMessage, brokerFill, strGenesisOrderID) Then
            If brokerFill.AccountID > 0 Then
                Set FillOrder = LoadOrder(brokerFill.AccountID, strGenesisOrderID, brokerFill.BrokerOrderID)
                If Not FillOrder Is Nothing Then
                    bFillExisted = False
                    bFillChanged = False
                    
                    For lIndex = 1 To FillOrder.SpreadFills.Count
                        If (FillOrder.SpreadFills(lIndex).BrokerID = brokerFill.BrokerID) Or ((Len(brokerFill.PreviousBrokerID) > 0) And (FillOrder.SpreadFills(lIndex).BrokerID = brokerFill.PreviousBrokerID)) Then
                            Set Fill = FillOrder.SpreadFills(lIndex)
                            
                            bFillExisted = True
                            bFillChanged = (Fill.FillDate <> brokerFill.FillDate) Or (Fill.Price <> brokerFill.Price) Or (Fill.Quantity <> brokerFill.Quantity)
                            
                            Exit For
                        End If
                    Next lIndex
                    
                    If (Not (Fill Is Nothing)) And (brokerFill.Quantity = 0) Then
                        Fill.Delete "Busted Fill"
                        If bRefresh = False Then
                            Refresh
                        End If
                    Else
                        If Fill Is Nothing Then
                            Set Fill = New cPtFill
                        End If
                        
                        With Fill
                            .AccountID = FillOrder.AccountID
                            .AutoTradingItemID = FillOrder.AutoTradeItemID
                            .BrokerID = brokerFill.BrokerID
                            .PreviousBrokerID = brokerFill.PreviousBrokerID
                            .BrokerOrderID = FillOrder.BrokerID
                            .Buy = brokerFill.Buy ' FillOrder.Buy
                            If .FillDate = 0# Then
                                .FillDate = brokerFill.FillDate
                            End If
                            .IsManual = False
                            .IsSnapshot = True
                            .IsSpread = True
                            .OrderID = FillOrder.OrderID
                            .Price = brokerFill.Price
                            .Quantity = brokerFill.Quantity
                            .SessionDate = brokerFill.SessionDate
                            If Len(brokerFill.Symbol) > 0 Then
                                .SymbolOrSymbolID = brokerFill.SymbolOrSymbolID
                            Else
                                .SymbolOrSymbolID = FillOrder.SymbolOrSymbolID
                            End If
                            
                            ' If we don't have a Fill ID yet (this is a new fill), we need to save
                            ' the fill to get a Fill ID for the key into the Fills collection of
                            ' the order...
                            If .FillID = 0& Then .Save
                        End With
                            
                        With FillOrder
                            Set .SpreadFills(Str(Fill.FillID)) = Fill
                            .IsSnapshot = True
                            
                            lFillQuantity = .FillQuantity
                            
                            ' Fix the status on the order if necessary...
                            If (lFillQuantity = .Quantity) Then
                                .Status = eTT_OrderStatus_Filled
                            ElseIf (lFillQuantity > .Quantity) Then
                                .Status = eTT_OrderStatus_OverFilled
                            ElseIf ((.Status = eTT_OrderStatus_Cancelled) Or (.Status = eTT_OrderStatus_Expired)) And (lFillQuantity > 0) Then
                                .Status = eTT_OrderStatus_BalCancelled
                            ElseIf ((.Status = eTT_OrderStatus_Working) Or (.Status = eTT_OrderStatus_Filled)) And (lFillQuantity > 0) Then
                                .Status = eTT_OrderStatus_Partial
                            End If
                            
                            .Save
                        End With
                        
                        m.BrokerInfo.AddOrder FillOrder, bRefresh
                        m.BrokerInfo.AddFill Fill, bRefresh
                        
                        FillCallback FillOrder, Fill, bFillExisted, False
                        If bRefresh = False Then
                            g.ActivityLogs.AddFillToActivityLog m.nBroker, Fill, False
                            m.dLastFillTime = gdTickCount
                        End If
                        DumpDebug mTradeTracker.FillDisplay(Fill)
                        
                        SendFillToOptionNav Fill, bRefresh
                    End If
                ElseIf brokerFill.Quantity > 0 Then
                    DumpDebug "Order with BrokerID = '" & brokerFill.BrokerOrderID & "' does not exist"
                    Set Fill = LoadFill(brokerFill.AccountID, brokerFill.BrokerID)
                    If (Not (Fill Is Nothing)) And (brokerFill.Quantity = 0) Then
                        Fill.Delete "Busted Fill"
                        If bRefresh = False Then
                            Refresh
                        End If
                    Else
                        If Fill Is Nothing Then
                            DumpDebug "Fill with BrokerID = '" & brokerFill.BrokerID & "' does not exist"
                            Set Fill = New cPtFill
                            bFillExisted = False
                        Else
                            DumpDebug "Fill with BrokerID = '" & brokerFill.BrokerID & "' exists"
                            bFillExisted = True
                        End If
                        
                        bFillChanged = (Fill.FillDate <> brokerFill.FillDate) Or (Fill.Price <> brokerFill.Price) Or (Fill.Quantity <> brokerFill.Quantity)
                        
                        With Fill
                            .AccountID = brokerFill.AccountID
                            .AutoTradingItemID = 0
                            .BrokerID = brokerFill.BrokerID
                            .BrokerOrderID = brokerFill.BrokerOrderID
                            .Buy = brokerFill.Buy
                            If .FillDate = 0# Then
                                .FillDate = brokerFill.FillDate
                            End If
                            .IsManual = False
                            .IsSnapshot = True
                            .IsSpread = True
                            .OrderID = 0
                            .Price = brokerFill.Price
                            .Quantity = brokerFill.Quantity
                            .SessionDate = brokerFill.SessionDate
                            .SymbolOrSymbolID = brokerFill.SymbolOrSymbolID
                            
                            ' If we don't have a Fill ID yet (this is a new fill), we need to save
                            ' the fill to get a Fill ID for the key into the Fills collection of
                            ' the order...
                            If .FillID = 0& Then .Save
                        End With
                            
                        m.BrokerInfo.AddFill Fill, bRefresh
                        
                        If bRefresh = False Then
                            g.ActivityLogs.AddFillToActivityLog m.nBroker, Fill, False
                            m.dLastFillTime = gdTickCount
                        End If
                        DumpDebug mTradeTracker.FillDisplay(Fill)
                        
                        SendFillToOptionNav Fill, bRefresh
                    End If
                Else
                    DumpDebug "Fill not processed because of a zero quantity"
                    If bRefresh = False Then
                        Refresh
                    End If
                End If
            Else
                DumpDebug "Fill not processed because account could not be converted"
            End If
        Else
            DumpDebug "Fill not processed because symbol could not be converted"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.HandleSpreadFillMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleCarriedFillMessage
'' Description: Handle a carried fill message received from the broker
'' Inputs:      Message, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleCarriedFillMessage(ByVal strMessage As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim brokerFill As cPtFill           ' Fill object loaded with information from the message
    Dim strGenesisOrderID As String     ' Genesis ID for the order the fill belongs to
    
    If Not m.BrokerObj Is Nothing Then
        If m.BrokerObj.CarriedFillFromInfo(strMessage, brokerFill, strGenesisOrderID) Then
            m.CarriedFills.Add brokerFill
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.HandleCarriedFillMessage"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandlePositionMessage
'' Description: Handle a position message received from the broker
'' Inputs:      Message, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandlePositionMessage(ByVal strMessage As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim BrokerPosition As cPtPosition   ' Position object to fill from the message
    Dim strKey As String                ' Key into the collection
    Dim FillSummary As cAccountPosition ' Account position object from the broker info
    Dim lIndex As Long                  ' Index into a for loop

    If Not m.BrokerObj Is Nothing Then
        Set BrokerPosition = New cPtPosition
        If m.BrokerObj.PositionFromInfo(strMessage, BrokerPosition) Then
            If BrokerPosition.AccountID > 0 Then
                ' 11/06/2014 DAJ: Don't process an unsolicited position where the position information
                ' hasn't changed.  Earlier this year, we discovered that Open E-Cry sends an unsolicited
                ' position update when the open equity changes.  I changed the stand-alone not to send
                ' those, but now we discovered that Rithmic is doing it as well...
                If (bRefresh = True) Or (CurrentPosition(BrokerPosition.AccountNumber, BrokerPosition.Symbol) <> BrokerPosition.CurrentPosition) Or (CarriedPosition(BrokerPosition.AccountNumber, BrokerPosition.Symbol) <> BrokerPosition.CarriedPosition) Then
                    m.BrokerInfo.AddPositionFromObject BrokerPosition, bRefresh
                    
                    strKey = BrokerPosition.AccountNumber & vbTab & BrokerPosition.Symbol
                    If m.BrokerPositions.Exists(strKey) Then
                        m.BrokerPositions(strKey) = BrokerPosition
                    Else
                        m.BrokerPositions.Add BrokerPosition, strKey
                    End If
                
                    ' If we get an unsolicited position refresh for only one symbol, make sure to
                    ' pass it along to Option Navigator as well...
                    If bRefresh = False Then
                        Set FillSummary = m.BrokerInfo.FillSummaryForSymbol(BrokerPosition.AccountNumber, BrokerPosition.Symbol, -1&)
                        If Not FillSummary Is Nothing Then
                            SendPositionToOptionNav FillSummary, bRefresh
                        End If
                    End If
                Else
                    DumpDebug vbTab & vbTab & "Position not processed because unsolicited position didn't change"
                End If
            Else
                DumpDebug vbTab & vbTab & "Position not processed because account could not be converted"
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.HandlePositionMessage"
    
End Sub

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
            .BrokerOrderID = Replace(.BrokerID, "CarPosFix", "CPF")
            .Buy = (FillSummary.CurrentPosition < 0&)
            .IsManual = True
            .IsSnapshot = False
            .Save
        End With
                
        DumpDebug vbTab & "Created fill at previous close: '" & mTradeTracker.FillDisplay(Fill) & "'"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.CreateFillAtPreviousClose"
    
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

    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object
    Dim strAccount As String            ' Account number
    Dim Bars As cGdBars                 ' Bars object
    Dim strKey As String                ' Key into the broker position collection
    Dim brokerPos As cPtPosition        ' Broker position object
    Dim carFill As cPtFill              ' Carried fill
    Dim lCarriedPosition As Long        ' Carried position
    Dim dPreviousClose As Double        ' Previous close for the symbol
    Dim dPreviousCloseTime As Double    ' Previous close time for the symbol
    Dim bUsedAvgEntry As Boolean        ' Did we use the average entry from the broker?
    
    If FillSummary Is Nothing Then
        Set FillSummary = g.Broker.FillSummary(vAccountNumberOrID, vSymbolOrSymbolID, -1&)
    End If
    
    If Not FillSummary Is Nothing Then
        Set Bars = New cGdBars
        SetBarProperties Bars, FillSummary.SymbolOrSymbolID
        
        strAccount = g.Broker.GetAccountNumber(vAccountNumberOrID)
        strKey = strAccount & vbTab & FillSummary.Symbol
        If m.BrokerPositions.Exists(strKey) Then
            Set brokerPos = m.BrokerPositions(strKey)
            If brokerPos.CarriedPosition = kNullData Then
                lCarriedPosition = brokerPos.CurrentPosition - FillSummary.NumBuysSnapshot + FillSummary.NumSellsSnapshot
                dPreviousClose = PreviousCloseForSymbol(FillSummary.SymbolOrSymbolID, dPreviousCloseTime)
                
                Set Fill = New cPtFill
                With Fill
                    .AccountID = FillSummary.AccountID
                    .AutoTradingItemID = 0&
                    .SymbolOrSymbolID = FillSummary.SymbolOrSymbolID
                    .Quantity = Abs(lCarriedPosition)
                    If (lCarriedPosition = brokerPos.CurrentPosition) And (brokerPos.AverageEntry <> 0) Then
                        bUsedAvgEntry = True
                        .Price = brokerPos.AverageEntry
                    Else
                        bUsedAvgEntry = False
                        .Price = dPreviousClose
                    End If
                    If brokerPos.ClearingDate <> 0# Then
                        .FillDate = brokerPos.ClearingDate
                    Else
                        .FillDate = ConvertToBrokerDate(dPreviousCloseTime, m.nBroker, FillSummary.Symbol, False)
                    End If
                    .SessionDate = Bars.SessionDateForTradeTime(dPreviousCloseTime)
                    .BrokerID = "CarPosFix " & .Symbol & " " & DateFormat(CurrentTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
                    .BrokerOrderID = Replace(.BrokerID, "CarPosFix", "CPF")
                    .Buy = (lCarriedPosition > 0)
                    .IsManual = True
                    .IsSnapshot = False
                    .Save
                End With
                
                If bUsedAvgEntry Then
                    DumpDebug vbTab & "Created fill for carried info: '" & mTradeTracker.FillDisplay(Fill) & "' (Used Average Entry from Broker)"
                Else
                    DumpDebug vbTab & "Created fill for carried info: '" & mTradeTracker.FillDisplay(Fill) & "' (Used Previous Close)"
                End If
            Else
                For lIndex = 1 To brokerPos.CarriedFills.Count
                    Set carFill = brokerPos.CarriedFills(lIndex)
                    
                    Set Fill = New cPtFill
                    With Fill
                        .AccountID = FillSummary.AccountID
                        .AutoTradingItemID = 0&
                        .SymbolOrSymbolID = FillSummary.SymbolOrSymbolID
                        .Quantity = carFill.Quantity
                        .Price = carFill.Price
                        .FillDate = carFill.FillDate
                        .SessionDate = Bars.SessionDateForTradeTime(ConvertToBrokerDate(.FillDate, m.nBroker, .Symbol, False))
                        If Len(carFill.BrokerID) > 0 Then
                            .BrokerID = "CarPosFix " & carFill.BrokerID
                        Else
                            .BrokerID = "CarPosFix " & .Symbol & " " & DateFormat(.FillDate, MM_DD_YYYY) & " " & Str(lIndex)
                        End If
                        If Len(carFill.BrokerOrderID) > 0 Then
                            .BrokerOrderID = carFill.BrokerOrderID
                        Else
                            .BrokerOrderID = Replace(.BrokerID, "CarPosFix", "CPF")
                        End If
                        .Buy = carFill.Buy
                        .IsManual = True
                        .IsSnapshot = False
                        .Save
                    End With
                    
                    DumpDebug vbTab & "Created fill for carried info: '" & mTradeTracker.FillDisplay(Fill) & "'"
                Next lIndex
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.CreateFillsForCarriedInfo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllowReconnect
'' Description: Allow Trade Navigator to attempt a reconnect to the broker?
'' Inputs:      Error Text
'' Returns:     True if allow reconnect, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AllowReconnect(ByVal strError As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    
    If (InStr(UCase(strError), "INVALID LOGIN") <> 0) Then
        bReturn = False
    ElseIf (InStr(UCase(strError), "THE PASSWORD IS INCORRECT") <> 0) Then
        bReturn = False
    ElseIf (InStr(UCase(strError), "USERNAME / PASSWORD NOT RECOGNIZED") <> 0) Then
        bReturn = False
    ElseIf (InStr(UCase(strError), "USERNAME DOES NOT EXIST") <> 0) Then
        bReturn = False
    ElseIf (UCase(strError) = "THE BRAND OF CLIENT IS NOT ALLOWED") Then
        bReturn = False
    ElseIf ((InStr(strError, "Perhaps someone else logged with this login.") <> 0) And (m.dLastConnection >= (CurrentTime - 1# / 1440#))) Then
        bReturn = False
    ElseIf (InStr(UCase(strError), "PERMISSION DENIED") <> 0) Then
        bReturn = False
    ElseIf (InStr(UCase(strError), "UNKNOWN APPLICATION") <> 0) Then
        bReturn = False
    ElseIf (InStr(UCase(strError), "INVALID USERID IN LOGON") <> 0) Then
        bReturn = False
    ElseIf (InStr(UCase(strError), "AUTHENTICATION FAILURE") <> 0) Then
        bReturn = False
    ElseIf (UCase(strError) = "INVALIDUSERORPASSWORD") Then
        bReturn = False
    ElseIf (UCase(strError) = "SOFTWARENOTPERMITTED") Then
        bReturn = False
    ElseIf (UCase(strError) = "EXPIRED") Then
        bReturn = False
    End If
    
    AllowReconnect = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.AllowReconnect"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAccountThreshold
'' Description: Load the account threshold
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadAccountThreshold()
On Error GoTo ErrSection:

    Dim lAccountThreshold As Long       ' Account threshold from the INI file
    
    lAccountThreshold = GetIniFileProperty("AccountThreshold", -1&, "Override", m.strIniFile)
    If (lAccountThreshold = -1&) And (Len(m.strConnectIni) > 0) Then
        lAccountThreshold = GetIniFileProperty("AccountThreshold", 10&, "Settings", m.strConnectIni)
    End If
    
    m.lAccountThreshold = lAccountThreshold

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.LoadAccountThreshold"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NumberOfAccountsReceived
'' Description: A NumberOfAccounts event has been received
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NumberOfAccountsReceived(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker message object
    Dim lNumberOfAccounts As Long       ' Number of accounts for the current login
    Dim bGetAccounts As Boolean         ' Get the accounts?
    
    bGetAccounts = True
    If (Len(strMessage) > 0) And (m.bIsBroker = False) Then
        Set brokerMsg = New cBrokerMessage
        
        brokerMsg.FromString strMessage
        lNumberOfAccounts = CLng(Val(brokerMsg("Count")))
        
        If lNumberOfAccounts > m.lAccountThreshold Then
            frmOnlineBroker.AddDialogMessage "You cannot login to more than " & Str(m.lAccountThreshold) & " accounts unless you are setup as a broker", "!", , m.strBrokerName & " Login Error"
            Disconnect False, "Too many accounts for a non-broker login"
            bGetAccounts = False
        End If
    End If
    
    If bGetAccounts Then
        GetAccounts
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.NumberOfAccountsReceived"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanSendOrder
'' Description: Determine if the given order is valid to be sent to the broker
'' Inputs:      Order, Source
'' Returns:     True if can be sent, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CanSendOrder(Order As cPtOrder, ByVal strSource As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strBrokerSymbol As String       ' Broker Symbol
    Dim strBrokerOrderType As String    ' Broker Order Type
    Dim strBrokerBase As String         ' Broker base symbol
    Dim strBrokerExchange As String     ' Broker exchange
    Dim strMessage As String            ' Message to display to user
    
    bReturn = False
    If Not m.BrokerObj Is Nothing Then
        strBrokerSymbol = m.BrokerObj.BrokerSymbol(Order.Symbol)
        strBrokerOrderType = m.BrokerObj.BrokerOrderType(Order.OrderType)
        
        If Len(strBrokerSymbol) = 0 Then
            Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
            g.Broker.ShowUnknownSymbolError Order.Symbol, m.nBroker, strSource, True, Me
        ' 12/04/2014 DAJ: With the changes in exchange fees coming from the exchange, there is more of a
        ' possibility that traders could have data turned off, but trading turned on for symbols.  Because of
        ' this, we need to let them trade symbols that they don't get data for...
        'ElseIf IsEnabledSymbol(Order.Symbol, strBrokerBase, strBrokerExchange) = False Then
        '    Order.Message = Order.Symbol & " not enabled through your " & m.strBrokerName & " account"
        '    g.Broker.ShowNotEnabledForSymbolError Order.Symbol, m.nBroker, strBrokerBase, strBrokerExchange, strSource, True, Me
        ElseIf Len(strBrokerOrderType) = 0 Then
            strMessage = OrderType(Order.OrderType) & " is not a valid order type for " & m.strBrokerName
            Order.Message = OrderType(Order.OrderType) & " cannot be converted to a " & m.strBrokerName & " order type"
        
            DumpDebug strSource & " Error shown to user: '" & Trim(Replace(strMessage, "|", " ")) & "'"
            InfBox strMessage, "!", , m.strBrokerName & " Order Error"
        ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
            Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
            g.Broker.ShowNotConnectedError Order.AccountID, m.nBroker, strSource, True, Me
        ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
            strMessage = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
            Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
            
            DumpDebug strSource & " Error shown to user: '" & Trim(Replace(strMessage, "|", " ")) & "'"
            InfBox strMessage, "!", , m.strBrokerName & " Order Error"
        ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
            strMessage = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
            Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
            
            DumpDebug strSource & " Error shown to user: '" & Trim(Replace(strMessage, "|", " ")) & "'"
            InfBox strMessage, "!", , m.strBrokerName & " Order Error"
        ElseIf g.Broker.ValidQuantityForOrder(Order) = False Then
            strMessage = Str(Order.Quantity) & " is an invalid quantity for '" & Order.Symbol & "'"
            Order.Message = strMessage
            
            DumpDebug strSource & " Error shown to user: '" & Trim(Replace(strMessage, "|", " ")) & "'"
            InfBox strMessage, "!", , m.strBrokerName & " Order Error"
        Else
            bReturn = True
        End If
    End If
    
    CanSendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.CanSendOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CopyStandalone
'' Description: Copy the standalone process
'' Inputs:      Source, Destination
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CopyStandalone(ByVal strSource As String, ByVal strDestination As String) As Boolean
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message to dump to log and display to the user

    FileCopy strSource, strDestination, True
    
ErrExit:
    CopyStandalone = True
    Exit Function
    
ErrSection:
    strMessage = "Could not copy '" & m.strProcessName & "'"
    DumpDebug strMessage
    InfBox strMessage, "!", , "File Copy Error"
    
    CopyStandalone = False
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanConvertBroker
'' Description: Can the broker be converted even though it is different?
'' Inputs:      Broker
'' Returns:     True if convert, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CanConvertBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If g.Broker.IsRithmicBroker(nBroker) And g.Broker.IsRithmicBroker(m.nBroker) Then
        bReturn = True
    ElseIf g.Broker.IsRjoBroker(nBroker) And g.Broker.IsRjoBroker(m.nBroker) Then
        bReturn = True
    ElseIf g.Broker.IsRcgPatsBroker(nBroker) And g.Broker.IsRcgPatsBroker(m.nBroker) Then
        bReturn = True
    ElseIf g.Broker.IsCqgBroker(nBroker) And g.Broker.IsCqgBroker(m.nBroker) Then
        bReturn = True
    ElseIf g.Broker.IsOecBroker(nBroker) And g.Broker.IsOecBroker(m.nBroker) Then
        bReturn = True
    End If
    
    CanConvertBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.CanConvertBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearRefreshFlags
'' Description: Clear the refresh flags
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearRefreshFlags()
On Error GoTo ErrSection:

    RefreshingAccounts = False
    RefreshingOrders = False
    RefreshingFills = False
    RefreshingPositions = False
    
    m.BrokerInfo.ClearRefreshFlags

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBroker.ClearRefreshFlags"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UseAmend
'' Description: Use Amend instead of Cancel/Replace for the given order
'' Inputs:      Order
'' Returns:     True if use Amend, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function UseAmend(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If g.Broker.IsOecBroker(Order.Broker) Then
        If IsForex(Order.Symbol) Then
            bReturn = False
        End If
    End If
    
    UseAmend = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.UseAmend"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTheWeekend
'' Description: Determine if it is between Friday night and Sunday morning
'' Inputs:      None
'' Returns:     True if the weekend, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsTheWeekend() As Boolean
On Error GoTo ErrSection:

    Dim dCurrentTime As Double          ' Current time
    Dim bReturn As Boolean              ' Return value for the function
    
    dCurrentTime = CurrentTime("NY", "", True)
    
    Select Case Weekday(dCurrentTime)
        Case vbFriday
            bReturn = (Hour(dCurrentTime) >= 17)
        
        Case vbSunday
            bReturn = (Hour(dCurrentTime) <= 12)
        
        Case vbSaturday
            bReturn = True
        
        Case Else
            bReturn = False
            
    End Select
    
    IsTheWeekend = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBroker.IsTheWeekend"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.astrAccounts = New cGdArray
    m.astrAccounts.Create eGDARRAY_Strings
    Set m.BrokerPositions = New cGdTree
    Set m.astrToFix = New cGdArray
    m.astrToFix.Create eGDARRAY_Strings
    Set m.abShowMenuItem = New cGdArray
    m.abShowMenuItem.Create eGDARRAY_TinyInts, eGDTradingMenu_NumItems, 0
    Set m.CarriedFills = New cGdTree
    Set m.RejectShown = New cGdTree
    
    Set m.WaitReplace = New cGdTree
    Set m.Originals = New cGdTree
    Set m.GidMap = New cGdTree

    m.bVerifyPositions = True
    m.bWasConnected = False
    m.bIsBroker = False
    m.lAccountThreshold = 10&
    m.bAllowBrokerOco = False
    m.dStartedProcess = 0#

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cBroker.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Termiante
'' Description: Clean up class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.astrAccounts = Nothing
    Set m.BrokerPositions = Nothing
    Set m.astrToFix = Nothing
    Set m.CarriedFills = Nothing
    Set m.RejectShown = Nothing
    Set m.WaitReplace = Nothing
    Set m.Originals = Nothing
    Set m.GidMap = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cBroker.Class_Terminate"
    
End Sub
