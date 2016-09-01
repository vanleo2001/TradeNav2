VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTransact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTransact.cls
'' Description: Wrapper class and member functions for trading through TransAct
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
'' 03/27/2009   DAJ         Implemented DropOff time for refreshes
'' 06/09/2009   DAJ         Set the BrokerInfo connection status
'' 07/30/2009   DAJ         Fix for incoming order with unknown symbol
'' 08/14/2009   DAJ         Changed simuser password from foxtrot to foxbat
'' 09/01/2009   DAJ         Use new Parked order status
'' 03/02/2010   DAJ         Use specific list of TransAct simulated accounts
'' 03/09/2010   DAJ         Put in AT Version info and new Vendor ID
'' 03/11/2010   DAJ         Use global collections
'' 03/16/2010   DAJ         Handle negative quantity orders/fills (desk entered)
'' 06/24/2010   DAJ         Changed to handle sending out the newer DLLs
'' 03/07/2011   DAJ         Broker Disconnected Alerts
'' 03/15/2011   DAJ         Fix for unsolicited amends
'' 05/27/2011   DAJ         Send Manual/Automated flag along with an order
'' 06/21/2011   DAJ         Last connection status change
'' 01/18/2012   DAJ         Enhanced logging for automated trading
'' 01/31/2012   DAJ         Handle unsolicted cancels in TradeSense order groups
'' 02/14/2012   DAJ         New status alerts for position mismatch / auto trade disabled
'' 03/07/2012   DAJ         Handle account set failure better
'' 03/21/2012   DAJ         Fix for amending a partially filled order
'' 10/30/2012   DAJ         Added tick count to each log entry
'' 11/13/2012   DAJ         Only handle fill if it came in a FilledOrder message
'' 11/14/2012   DAJ         Also handle a fill if it came in a refresh
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 12/11/2012   DAJ         Use the flatten queue for position reversals
'' 12/11/2012   DAJ         Handle broker status in automated trading item
'' 01/14/2013   DAJ         Ignore fills with zero quantity
'' 01/15/2013   DAJ         Check for Genesis symbol in Enabled symbols, not TransAct symbol
'' 02/12/2013   DAJ         Move busy box for refresh when first connected
'' 02/20/2013   DAJ         Don't try to reconnect if we get the new Error 75
'' 02/20/2013   DAJ         Log the login mode/registry values upon logon attempt
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
'' 04/17/2013   DAJ         Added direction filter to CancelWorkingOrders call
'' 04/25/2013   DAJ         Tweaked order logging, fix to alleviate duplicate order when amending
'' 07/19/2013   DAJ         New TransAct API
'' 01/16/2014   DAJ         Handle negative quantity orders ( TransAct undo's ) better
'' 04/24/2014   DAJ         Confirm flatten order when account associated to lot
'' 05/13/2014   DAJ         Fix for timeout trying to start stand alone process
'' 08/08/2014   DAJ         Reworked enabled symbols list for TransAct
'' 10/24/2014   DAJ         Fill Display
'' 10/29/2014   DAJ         Remove old synthetic order/MIT code
'' 01/21/2014   DAJ         Set g.Broker.PositionsToVerify when VerifyPositions changes not ConnectionStatus
'' 09/01/2015   DAJ         Don't count fill on order message if status is not 3, 1000, or 1001
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Development Vendor ID:
'Private Const kVendorID = "E6C6D3A3-5D5F-4b06-811B-88109AD20589"
' Production Vendor ID (Versions before 5.0.0):
Private Const kVendorID4 = "B9D60737-0BD0-4c5e-BBC7-FC8D661071BA"
' 10-Deep Vendor ID (Versions on or after 5.0.0):
Private Const kVendorID5 = "4C5099DB-9388-41db-A9E4-D9A64B95424A"
Private Const kSimUserID = "simuser" ' "simnav"
Private Const kSimUserPassword = "foxbat" '"foxtrot" '"genesis9"

Private Const kRegKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
Private Const kDefaultSymbols = "G6E,ES,NQ,YM,ZB,ZN"

Private Type mPrivate
    strIniFile As String                ' Ini file for TransAct
    nBroker As eTT_AccountType          ' Broker type for the TransAct broker
    strBrokerName As String             ' Broker name for the TransAct broker
    strProcessName As String            ' Name of the process to check
    strProcessPath As String            ' Full path of the process to start
    strControlID As String              ' Control ID
    strLogPath As String                ' Path for the log files
    strDemoAccounts As String           ' Demo accounts
    strVersion As String                ' Yes Trader version
    
    nLoginMode As eGDTransActLoginModes ' Login Mode
    strUserName As String               ' User Name to log on with
    strPassword As String               ' Password to log on with
    strIniPassword As String            ' Password from the ini file
    strIP As String                     ' IP of server to log on to
    strAccountNumber As String          ' Account number connected to
        
    strAskedFor As String               ' String of symbols attempted
    strSub As String                    ' String of symbols subscribed
        
    dHeartBeat As Double                ' Heartbeat time
    dLastStatusMsg As Double            ' Tick count of the last status message
    bReconnecting As Boolean            ' Are we currently reconnecting?
    lReconnectAttempts As Long          ' Reconnect attempts that have been made
    dLastFillTime As Double             ' Tick count of when we received the last fill
    bVerifyPositions As Boolean         ' Do we want to verify the positions with the user?
    bAllowGTC As Boolean                ' Does TransAct allow GTC orders for the set account?
    dStartedProcess As Double           ' Tick count when we started the broker stand-alone
    
    astrBaseSyms As cGdArray            ' Array of base symbol information
    astrPrices As cGdArray              ' Array of price structures
    astrAvailable As cGdArray           ' Array of subscribed symbols
    astrEnabledSymbols As cGdArray      ' Array of Genesis symbols the user is enabled for
    astrSubContract As cGdArray         ' Array of subscribed contracts
    astrLastFillTime As cGdArray        ' Last fill time for each symbol
    astrBrokerPos As cGdArray           ' Broker position information
    astrToFix As cGdArray               ' List of account/symbol pairs that need their position fixed
    
    nStatus As eGDConnectionStatus      ' Current connection status
    bWasConnected As Boolean            ' This broker was has been connected
    bClearInf As Boolean                ' Clear InfBox?
    bAskedPassword As Boolean           ' Have we asked the user for their password?
    bManualDisconnect As Boolean        ' Is this a manual disconnect?
    bAppLoaded As Boolean               ' Is the GenTransact application loaded?
    bSuccessfullyConnected As Boolean   ' Have we successfully connected once?
    bRefreshing As Boolean              ' Are we currently refreshing the account?
    bVerboseRefresh As Boolean          ' Should we tell the user what we are doing?
    bSyncInProgress As Boolean          ' Are we in the process of synchronizing positions?
    bFixCalledOnce As Boolean           ' Has the fix positions call been done at least once?
    lNumSetAccountFailures As Long      ' Number of "Set Account" failure messages
    
    dRefreshingStart As Double          ' Start of the refreshing time
    bRefreshingAccounts As Boolean      ' Are we refreshing accounts?
    bRefreshingOrders As Boolean        ' Are we refreshing orders?
    bRefreshingFills As Boolean         ' Are we refreshing fills?
    bRefreshingPositions As Boolean     ' Are we refreshing positions?
    
    SymbolBars As cGdTree
    ModifiedOrders As cGdTree           ' Collection of modified orders
    
    BrokerInfo As cBrokerInfo           ' Object of broker information
End Type
Private m As mPrivate

Public Property Get LoginMode() As eGDTransActLoginModes
    LoginMode = m.nLoginMode
End Property
Public Property Let LoginMode(ByVal nLoginMode As eGDTransActLoginModes)
    m.nLoginMode = nLoginMode
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

Public Property Get IPAddress() As String
    IPAddress = m.strIP
End Property
Public Property Let IPAddress(ByVal strIP As String)
    m.strIP = strIP
End Property

Public Property Get HeartBeat() As Double
    HeartBeat = m.dHeartBeat
End Property

Public Property Get Account() As String
    Account = m.strAccountNumber
End Property
Public Property Let Account(ByVal strAccountNumber As String)
    m.strAccountNumber = strAccountNumber
End Property

Public Property Get ConnectionStatus() As eGDConnectionStatus
    ConnectionStatus = m.nStatus
End Property
Public Property Let ConnectionStatus(ByVal nNewStatus As eGDConnectionStatus)
    If nNewStatus <> m.nStatus Then
        m.nStatus = nNewStatus
        
        If Not g.Broker Is Nothing Then
            g.Broker.LastStatusChange = Now
        End If
        'If nNewStatus = eGDConnectionStatus_Connected Then
        '    g.Broker.PositionsToVerify = True
        'End If
    End If
End Property

Public Property Get LastStatusMessage() As Double
    LastStatusMessage = m.dLastStatusMsg
End Property

Public Property Get AppLoaded() As Boolean
    AppLoaded = m.bAppLoaded
End Property

Public Property Get ModifiedOrders() As cGdTree
    Set ModifiedOrders = m.ModifiedOrders
End Property
Public Property Let ModifiedOrders(ByVal ModOrders As cGdTree)
    Set m.ModifiedOrders = ModOrders
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
    If bVerifyPositions <> m.bVerifyPositions Then
        DumpDebug vbTab & "PositionVerify changed from " & Str(m.bVerifyPositions) & " to " & Str(bVerifyPositions)
        m.bVerifyPositions = bVerifyPositions
        
        If m.bVerifyPositions = True Then
            g.Broker.PositionsToVerify = True
        End If
    End If
End Property

Public Property Get SimUserUserName() As String
    SimUserUserName = kSimUserID
End Property
Public Property Get SimUserPassword() As String
    SimUserPassword = kSimUserPassword
End Property

Public Property Get DemoAccounts() As String
    DemoAccounts = m.strDemoAccounts
End Property

Public Property Get EnabledSymbols() As cGdArray
    Set EnabledSymbols = m.astrEnabledSymbols
End Property

Public Property Get ProcessName() As String
    ProcessName = m.strProcessName
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

Public Property Get LastDateConnected() As Long
    LastDateConnected = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", m.strIniFile))))
End Property
Private Property Let LastDateConnected(ByVal lLastDateConnected As Long)
    SetIniFileProperty "Last", EncryptToHex(Str(lLastDateConnected)), "Connect", m.strIniFile
End Property

Private Property Get LKSCRC() As Long
    LKSCRC = GetRegistryValue(rkLocalMachine, kRegKey, "LKSC" & Str(m.nBroker), 0&)
End Property
Private Property Let LKSCRC(ByVal lLksCrc As Long)
    SetRegistryValue rkLocalMachine, kRegKey, "LKSC" & Str(m.nBroker), lLksCrc, True
End Property

Private Property Get LksFile() As String
    LksFile = AddSlash(App.Path) & "Provided\LKS." & Str(m.nBroker)
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
    RaiseError "cTransact.CurrentPosition.Get"
    
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
    RaiseError "cTransact.CurrentPosition.Let"
    
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
    RaiseError "cTransact.CarriedPosition.Get"
    
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
    RaiseError "cTransact.CarriedPosition.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LogonForAccountList
'' Description: Attempt to logon to the TransAct servers for an account list
'' Inputs:      User ID, Password
'' Returns:     True if request sent, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LogonForAccountList(ByVal strUserID As String, ByVal strPassword As String) As Boolean
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Time out variable
    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If (frmOnlineBroker.gdBroker.FindControlID(m.strControlID) = 0) Or (KillProcess(m.strProcessName, True) = 0) Then
        If KillProcess(m.strProcessName) > 0 Then
            Sleep 3
        End If
        
        DumpDebug "Trade Navigator version: " & FormatVersion & "." & Str(App.Revision) & " " & DateFormat(FileDate(App.Path & "\" & App.EXEName & ".EXE"), MM_DD_YYYY, HH_MM, AMPM_UPPER)
        DumpDebug "Starting Process: " & m.strProcessPath
        RunProcess m.strProcessPath, , , vbHide ' , vbMinimizedNoFocus
        StartedProcess = gdTickCount
    End If
    
    lTimeOut = 0&
    Do While (m.bAppLoaded = False) And (lTimeOut < 30&)
        Sleep 1#
        lTimeOut = lTimeOut + 1&
    Loop
    
    If m.bAppLoaded = True Then
        m.bManualDisconnect = True
        
        SendTransactMessage eGDTransActMessageType_LogonForAccountList, strUserID & vbTab & strPassword & vbTab & VendorID & vbTab & Str(m.nLoginMode)
        bReturn = True
    End If
    
    LogonForAccountList = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.LogonForAccountList"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Logon
'' Description: Attempt to logon to the TransAct servers
'' Inputs:      User ID, Password, Account
'' Returns:     True if request sent, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Logon(ByVal strUserID As String, ByVal strPassword As String, ByVal strAccountID As String) As Boolean
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Time out variable
    Dim bReturn As Boolean              ' Return value for the function

    If (g.RealTime.ConnectionStatus <> eGDConnectionStatus_Connecting) And (m.strUserName <> SimUserUserName) Then
        InfBox "Connecting to " & m.strBrokerName & " server as " & m.strUserName & "|in " & TransActLoginModeString(m.nLoginMode) & " mode||Please wait...", , , m.strBrokerName & " Logon", True
    End If
    
    bReturn = False
    If (frmOnlineBroker.gdBroker.FindControlID(m.strControlID) = 0) Or (KillProcess(m.strProcessName, True) = 0) Then
        If KillProcess(m.strProcessName) > 0 Then
            Sleep 5
        End If
        
        DumpDebug "Trade Navigator version: " & FormatVersion & "." & Str(App.Revision) & " " & DateFormat(FileDate(App.Path & "\" & App.EXEName & ".EXE"), MM_DD_YYYY, HH_MM, AMPM_UPPER)
        DumpDebug "Starting Process: " & m.strProcessPath
        RunProcess m.strProcessPath, , , vbHide
        StartedProcess = gdTickCount
    End If
    
    lTimeOut = 0&
    Do While (m.bAppLoaded = False) And (lTimeOut < 30&)
        Sleep 1#
        lTimeOut = lTimeOut + 1&
    Loop
    
    If m.bAppLoaded = True Then
        SendTransactMessage eGDTransActMessageType_Logon, strUserID & vbTab & strPassword & vbTab & VendorID & vbTab & strAccountID & vbTab & Str(m.nLoginMode)
        bReturn = True
    End If
    
    Logon = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.Logon"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Connect
'' Description: Attempt to connect to the TransAct servers
'' Inputs:      User ID, Account, Switching?
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Connect(Optional ByVal strUserID As String = "", Optional ByVal strAccountNumber As String = "", Optional ByVal bSwitching As Boolean = False) As Boolean
On Error GoTo ErrSection:
    
    Dim bConnect As Boolean             ' Should we attempt to connect?
    Dim strProcess As String            ' Name of the process to start
    Dim lCounter As Long                ' Counter variable
    Dim strPassword As String           ' Password back from InfBox
    Dim lExpireAsk As Long              ' Have we asked about the simuser expiration on an auto-connect?
    Dim lTimeOut As Long                ' Timeout value
    Dim bReturn As Boolean              ' Return value for the function
    Static bOsMessageShown As Boolean   ' Have we shown the user the Operating System Message?
    
    bReturn = False
    
    If Is9598orMe Then
        If bOsMessageShown = False Then
            bOsMessageShown = True
            InfBox m.strBrokerName & " cannot run under Windows 95, Windows 98, or Windows ME", "!", , m.strBrokerName & " Error"
            DumpDebug m.strBrokerName & " cannot run under Windows 95, Windows 98, or Windows ME"
        End If
    ElseIf (g.bUnloading = False) Then
        ' The user is not currently connected to the TransAct servers...
        If (ConnectionStatus = eGDConnectionStatus_Disconnected) Then
            ' If we have already asked for the password for this username, don't ask again...
            If (m.bAskedPassword = True) And ((strUserID = m.strUserName) Or (Len(strUserID) = 0)) Then
                If Len(strAccountNumber) > 0 Then m.strAccountNumber = strAccountNumber
                bConnect = True
                
            ' Otherwise, ask the user for the password...
            Else
                bConnect = frmTransActLogin.ShowMe(strUserID, strAccountNumber)
                m.bAskedPassword = bConnect
            End If
        
            ' If all is OK, attempt to connect to the TransAct servers...
            If (bConnect = True) And (g.bUnloading = False) Then
                If LiveTradingAllowed(m.nBroker) Then
                    ConnectionStatus = eGDConnectionStatus_Connecting
                    Logon m.strUserName, m.strPassword, m.strAccountNumber
                    bReturn = True
                End If
            End If
        
        ' The user is already connected to the TransAct servers...
        ElseIf (ConnectionStatus = eGDConnectionStatus_Connected) Then
            ' If the user ID is different, we need to disconnect and reconnect...
            If ((strUserID <> m.strUserName) And (Len(strUserID) > 0)) Then
                If InfBox("You are already connected to '" & g.Broker.AccountNameForNumber(m.strAccountNumber) & "'.|Do you want to disconnect from that account and connect to '" & g.Broker.AccountNameForNumber(strAccountNumber) & "'?|", "?", "+Yes|-No", m.strBrokerName & " Connection") = "Y" Then
                    ' Disconnect from the servers...
                    m.bManualDisconnect = True
                    Disconnect m.strAccountNumber, False, "Changing User IDs from " & m.strUserName & " to " & strUserID
                    
                    lTimeOut = 0&
                    Do While (m.bAppLoaded = True) And (lTimeOut < 30&)
                        Sleep 1#
                        lTimeOut = lTimeOut + 1&
                    Loop
                    
                    ' If successfully disconnected, attempt to connect with the new information...
                    If (ConnectionStatus = eGDConnectionStatus_Disconnected) Then
                        bReturn = Connect(strUserID, strAccountNumber, True)
                    End If
                End If
            
            ' Otherwise, if the account is different, switch accounts...
            ElseIf ((strAccountNumber <> m.strAccountNumber) And (Len(strAccountNumber) > 0)) Then
                If InfBox("You are already connected to '" & g.Broker.AccountNameForNumber(m.strAccountNumber) & "'.|Do you want to disconnect from that account and connect to '" & g.Broker.AccountNameForNumber(strAccountNumber) & "'?|", "?", "+Yes|-No", m.strBrokerName & " Connection") = "Y" Then
                    m.bManualDisconnect = True
                    DumpDebug "Changing to account " & strAccountNumber & " from " & m.strAccountNumber
                    SendTransactMessage eGDTransActMessageType_ChangeAccount, m.strAccountNumber & vbTab & strAccountNumber
                    bReturn = True
                End If
            End If
        End If
    End If
    
    Connect = bReturn
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.Connect"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disconnect
'' Description: Attempt to disconnect from the TransAct servers
'' Inputs:      Account, Force Reconnect?, Reason
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Disconnect(Optional ByVal strAccountNumber As String = "", Optional ByVal bForceReconnect As Boolean = False, Optional ByVal strReason As String = "") As Boolean
On Error GoTo ErrSection:

    If Len(strAccountNumber) = 0 Then strAccountNumber = m.strAccountNumber

    If (ConnectionStatus = eGDConnectionStatus_Connected) And (strAccountNumber = m.strAccountNumber) Then
        ConnectionStatus = eGDConnectionStatus_Disconnecting
    
        m.bManualDisconnect = Not bForceReconnect
        If m.bManualDisconnect Then
            m.dHeartBeat = 0#
            g.OrderStrategies.DeactivateExitsForBroker m.nBroker, "Disconnecting from " & m.strBrokerName
        End If
        
        If Len(strReason) = 0 Then DumpDebug "Disconnecting" Else DumpDebug "Disconnecting: " & strReason
        SendTransactMessage eGDTransActMessageType_Disconnect, strAccountNumber
        Disconnect = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.Disconnect"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SwithLoginModes
'' Description: Attempt to switch login modes
'' Inputs:      New Login Mode
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SwitchLoginModes(ByVal nNewLoginMode As eGDTransActLoginModes) As Boolean
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout value
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If nNewLoginMode <> m.nLoginMode Then
        ' Disconnect from the servers...
        m.bManualDisconnect = True
        Disconnect m.strAccountNumber, False, "Changing Login Modes from " & TransActLoginModeString(m.nLoginMode) & " to " & TransActLoginModeString(nNewLoginMode)
        
        lTimeOut = 0&
        Do While (KillProcess(m.strProcessName, True) <> 0) And (lTimeOut < 30&)
            Sleep 1#
            lTimeOut = lTimeOut + 1&
        Loop
        Sleep 1#
        
        m.nLoginMode = nNewLoginMode
        
        ' If successfully disconnected, attempt to connect with the new information...
        If (ConnectionStatus = eGDConnectionStatus_Disconnected) Then
            bReturn = Connect(m.strUserName, m.strAccountNumber, True)
        End If
    Else
        DumpDebug "User attempted to switch to existing login mode (" & TransActLoginModeString(m.nLoginMode) & ")"
    End If
    
    SwitchLoginModes = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.SwitchLoginModes"
    
End Function

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

    If NotSent(Order.Status) Then
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
    RaiseError "cTransact.SendOrder"

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

    Dim strTransactOrder As String      ' TransAct order
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    Dim strBrokerBase As String         ' Broker base symbol
    Dim strBrokerExchange As String     ' Broker exchange
    Dim strMessage As String            ' Message to display to user

    bReturn = False
    nPrevOrderType = Order.OrderType
    strTransactOrder = TransactOrder(Order)
    
    If Len(strTransactOrder) = 0 Then
        'DumpDebug "AddOrder: " & Order.Symbol & " cannot be traded on the " & m.strBrokerName & " servers"
        'InfBox Order.Symbol & " cannot be traded on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be traded on the " & m.strBrokerName & " servers"
        g.Broker.ShowUnknownSymbolError Order.Symbol, m.nBroker, "AddOrder", True
    ElseIf IsEnabledSymbol(Order.Symbol, strBrokerBase, strBrokerExchange) = False Then
        Order.Message = Order.Symbol & " not enabled through your " & m.strBrokerName & " account"
        g.Broker.ShowNotEnabledForSymbolError Order.Symbol, m.nBroker, strBrokerBase, strBrokerExchange, "AddOrder", True
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        'DumpDebug "AddOrder: Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID)
        'InfBox "Not currently connected to " & m.strBrokerName & " account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
        g.Broker.ShowNotConnectedError Order.AccountID, m.nBroker, "AddOrder", True
    ElseIf m.nLoginMode = eGDTransActLoginMode_Demo Then
        DumpDebug "AddOrder: Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in DEMO mode"
        InfBox "You cannot place an order because you are currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in DEMO mode", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in DEMO mode"
    ElseIf m.nLoginMode = eGDTransActLoginMode_SimLive Then
        DumpDebug "AddOrder: Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in SIMULATED LIVE mode"
        InfBox "You cannot place an order because you are currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in SIMULATED LIVE mode", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in SIMULATED LIVE mode"
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "AddOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "AddOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    ElseIf g.Broker.ValidQuantityForOrder(Order) = False Then
        strMessage = Str(Order.Quantity) & " is an invalid quantity for '" & Order.Symbol & "'"
        Order.Message = strMessage
        DumpDebug "AddOrder: Error shown to user: '" & Trim(Replace(strMessage, "|", " ")) & "'"
        InfBox strMessage, "!", , m.strBrokerName & " Order Error"
    Else
        Order.BrokerID = Order.GenesisOrderID
        Order.Save
        'GenesisOrderToBrokerInfo Order, eTT_OrderStatus_Sent
        
        DumpDebug "AddOrder: '" & Order.OrderText(True, True, True) & "'"
        SendTransactMessage eGDTransActMessageType_AddOrder, strTransactOrder
        bReturn = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
    AddOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.AddOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AmendOrder
'' Description: Attempt to amend the order
'' Inputs:      Order to amend
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AmendOrder(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim strTransactOrder As String      ' TransAct order
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    Dim strBrokerBase As String         ' Broker base symbol
    Dim strBrokerExchange As String     ' Broker exchange
    Dim strMessage As String            ' Message to display to user
    
    ' Modify order needs the remaining quantity, not the absolute quantity...
    Order.Quantity = Order.RemainingQuantity
    nPrevOrderType = Order.OrderType
    strTransactOrder = TransactOrder(Order)
    bReturn = False
    
    If Len(strTransactOrder) = 0 Then
        'DumpDebug "ModifyOrder: " & Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers"
        'InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be traded on the " & m.strBrokerName & " servers"
        g.Broker.ShowUnknownSymbolError Order.Symbol, m.nBroker, "AmendOrder", True
    ElseIf IsEnabledSymbol(Order.Symbol, strBrokerBase, strBrokerExchange) = False Then
        Order.Message = Order.Symbol & " not enabled through your " & m.strBrokerName & " account"
        g.Broker.ShowNotEnabledForSymbolError Order.Symbol, m.nBroker, strBrokerBase, strBrokerExchange, "AmendOrder", True
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        'DumpDebug "AmendOrder: Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID)
        'InfBox "Not currently connected to " & m.strBrokerName & " account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
        g.Broker.ShowNotConnectedError Order.AccountID, m.nBroker, "AmendOrder", True
    ElseIf m.nLoginMode = eGDTransActLoginMode_Demo Then
        DumpDebug "AmendOrder: Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in DEMO mode"
        InfBox "You cannot amend an order because you are currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in DEMO mode", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in DEMO mode"
    ElseIf m.nLoginMode = eGDTransActLoginMode_SimLive Then
        DumpDebug "AmendOrder: Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in SIMULATED LIVE mode"
        InfBox "You cannot amend an order because you are currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in SIMULATED LIVE mode", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in SIMULATED LIVE mode"
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "ModifyOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " through the " & m.strBrokerName & " servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "AmendOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    ElseIf g.Broker.ValidQuantityForOrder(Order) = False Then
        strMessage = Str(Order.Quantity) & " is an invalid quantity for '" & Order.Symbol & "'"
        Order.Message = strMessage
        DumpDebug "AmendOrder: Error shown to user: '" & Trim(Replace(strMessage, "|", " ")) & "'"
        InfBox strMessage, "!", , m.strBrokerName & " Order Error"
    Else
        DumpDebug "ModifyOrder: '" & Order.OrderText(True, True, True) & "'"
        SendTransactMessage eGDTransActMessageType_AmendOrder, strTransactOrder
        bReturn = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
    AmendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.AmendOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Attempt to cancel the order
'' Inputs:      Order to Cancel
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim strTransactOrder As String      ' TransAct order
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    Dim strBrokerBase As String         ' Broker base symbol
    Dim strBrokerExchange As String     ' Broker exchange
    
    bReturn = False
    nPrevOrderType = Order.OrderType
    strTransactOrder = TransactOrder(Order)
    
    If NotSent(Order.Status) Then
        DumpDebug "CancelOrder: '" & Order.OrderText(True, True, True) & "'"
        Order.ChangeOrderStatus eTT_OrderStatus_Cancelled, m.BrokerInfo
    ElseIf Len(strTransactOrder) = 0 Then
        'DumpDebug "CancelOrder: " & Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers"
        'InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers"
        g.Broker.ShowUnknownSymbolError Order.Symbol, m.nBroker, "CancelOrder", True
    ElseIf IsEnabledSymbol(Order.Symbol, strBrokerBase, strBrokerExchange) = False Then
        Order.Message = Order.Symbol & " not enabled through your " & m.strBrokerName & " account"
        g.Broker.ShowNotEnabledForSymbolError Order.Symbol, m.nBroker, strBrokerBase, strBrokerExchange, "CancelOrder", True
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        'DumpDebug "CancelOrder: Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID)
        'InfBox "Not currently connected to " & m.strBrokerName & " account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
        g.Broker.ShowNotConnectedError Order.AccountID, m.nBroker, "CancelOrder", True
    ElseIf m.nLoginMode = eGDTransActLoginMode_Demo Then
        DumpDebug "CancelOrder: Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in DEMO mode"
        InfBox "You cannot cancel an order because you are currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in DEMO mode", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in DEMO mode"
    ElseIf m.nLoginMode = eGDTransActLoginMode_SimLive Then
        DumpDebug "CancelOrder: Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in SIMULATED LIVE mode"
        InfBox "You cannot cancel an order because you are currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in SIMULATED LIVE mode", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in SIMULATED LIVE mode"
    Else
        If Not g.TsoGroups Is Nothing Then
            g.TsoGroups.CancelRequested Order
        End If
        
        DumpDebug "CancelOrder: '" & Order.OrderText(True, True, True) & "'"
        SendTransactMessage eGDTransActMessageType_CancelOrder, strTransactOrder
        
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
    RaiseError "cTransact.CancelOrder"
    
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

    Dim strTransactOrder As String      ' TransAct order
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    Dim strBrokerBase As String         ' Broker base symbol
    Dim strBrokerExchange As String     ' Broker exchange
    
    bReturn = False
    nPrevOrderType = Order.OrderType
    strTransactOrder = TransactOrder(Order)
    
    If Len(strTransactOrder) = 0 Then
        'DumpDebug "ParkOrder: " & Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers"
        'InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers"
        g.Broker.ShowUnknownSymbolError Order.Symbol, m.nBroker, "ParkOrder", True
    ElseIf IsEnabledSymbol(Order.Symbol, strBrokerBase, strBrokerExchange) = False Then
        Order.Message = Order.Symbol & " not enabled through your " & m.strBrokerName & " account"
        g.Broker.ShowNotEnabledForSymbolError Order.Symbol, m.nBroker, strBrokerBase, strBrokerExchange, "ParkOrder", True
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        'DumpDebug "ParkOrder: Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID)
        'InfBox "Not currently connected to " & m.strBrokerName & " account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
        g.Broker.ShowNotConnectedError Order.AccountID, m.nBroker, "ParkOrder", True
    ElseIf m.nLoginMode = eGDTransActLoginMode_Demo Then
        DumpDebug "ParkOrder: Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in DEMO mode"
        InfBox "You cannot park an order because you are currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in DEMO mode", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in DEMO mode"
    ElseIf m.nLoginMode = eGDTransActLoginMode_SimLive Then
        DumpDebug "ParkOrder: Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in SIMULATED LIVE mode"
        InfBox "You cannot park an order because you are currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in SIMULATED LIVE mode", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID) & " in SIMULATED LIVE mode"
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "ParkOrder: " & OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " through the " & m.strBrokerName & " servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed for " & Order.Symbol & " through the " & m.strBrokerName & " servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "ParkOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    Else
        DumpDebug "ParkOrder: '" & Order.OrderText(True, True, True) & "'"
        SendTransactMessage eGDTransActMessageType_CancelOrder, strTransactOrder
        
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
    RaiseError "cTransact.ParkOrder"
    
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
    RaiseError "cTransact.CancelOrderFromBrokerID"
    
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
    RaiseError "cTransact.CancelWorkingOrders"
    
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
    
    If m.nLoginMode = eGDTransActLoginMode_Demo Then
        InfBox "You cannot flatten your position because you are logged in as DEMO mode", "!", , "Flatten Error"
        DumpDebug "You cannot flatten your position because you are logged in as DEMO mode"
    ElseIf m.nLoginMode = eGDTransActLoginMode_SimLive Then
        InfBox "You cannot flatten your position because you are logged in as SIMULATED LIVE mode", "!", , "Flatten Error"
        DumpDebug "You cannot flatten your position because you are logged in as SIMULATED LIVE mode"
    Else
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
                DumpDebug "User Confirmed: " & strMessage & " in account " & strAccount
                bReturn = ExitPositionForSymbol(strAccount, strSymbol, lAtID, lPosition, strGenesisOrderID)
            End If
        End If
    End If
    
    FlattenPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.FlattenPosition"
    
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
    
    If m.nLoginMode = eGDTransActLoginMode_Demo Then
        InfBox "You cannot reverse your position because you are logged in as DEMO mode", "!", , "Reverse Error"
        DumpDebug "You cannot reverse your position because you are logged in as DEMO mode"
    ElseIf m.nLoginMode = eGDTransActLoginMode_SimLive Then
        InfBox "You cannot reverse your position because you are logged in as SIMULATED LIVE mode", "!", , "Reverse Error"
        DumpDebug "You cannot reverse your position because you are logged in as SIMULATED LIVE mode"
    Else
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
    End If
    
    ReversePosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.ReversePosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Request accounts from the GenTransact program
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAccounts()
On Error GoTo ErrSection:

    If Not Refreshing Then
        DumpDebug "Getting Accounts"
        m.bRefreshingAccounts = True
        SendTransactMessage eGDTransActMessageType_GetAccounts, ""
    Else
        DumpDebug "Account Refresh requested but already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.GetAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Ask the GenTransact program for the orders and trades that it
''              has for an account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetOrders()
On Error GoTo ErrSection:

    If m.nLoginMode = eGDTransActLoginMode_Demo Then
        InfBox "You cannot request orders because you are logged in as DEMO mode", "!", , "Request Orders Error"
        DumpDebug "You cannot request orders because you are logged in as DEMO mode"
    ElseIf m.nLoginMode = eGDTransActLoginMode_SimLive Then
        InfBox "You cannot request orders because you are logged in as SIMULATED LIVE mode", "!", , "Request Orders Error"
        DumpDebug "You cannot request orders because you are logged in as SIMULATED LIVE mode"
    Else
        If Not Refreshing Then
            DumpDebug "Getting Orders and Trades"
            m.bRefreshingOrders = True
            m.bRefreshingFills = True
            SendTransactMessage eGDTransActMessageType_GetOrders, m.strAccountNumber
            SendTransactMessage eGDTransActMessageType_GetTrades, m.strAccountNumber
        Else
            DumpDebug "Order refresh requested but already in progress"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.GetOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositions
'' Description: Ask the GenTransact program for the positions that it has for
''              an account
'' Inputs:      Symbols
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPositions(Optional ByVal strSymbols As String = "")
On Error GoTo ErrSection:

    If m.nLoginMode = eGDTransActLoginMode_Demo Then
        InfBox "You cannot request positions because you are logged in as DEMO mode", "!", , "Request Positions Error"
        DumpDebug "You cannot request positions because you are logged in as DEMO mode"
    ElseIf m.nLoginMode = eGDTransActLoginMode_SimLive Then
        InfBox "You cannot request positions because you are logged in as SIMULATED LIVE mode", "!", , "Request Positions Error"
        DumpDebug "You cannot request positions because you are logged in as SIMULATED LIVE mode"
    Else
        If Not Refreshing Then
            If Len(strSymbols) = 0 Then DumpDebug "Getting Positions" Else DumpDebug "Getting Positions for " & strSymbols
            m.bRefreshingPositions = True
            SendTransactMessage eGDTransActMessageType_GetPositions, m.strAccountNumber & vbTab & strSymbols
        Else
            DumpDebug "Positions refresh requested but already in progress"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdatePositions
'' Description: Ask the GenTransact program to query the TransAct servers for
''              the positions that it has for an account
'' Inputs:      Symbols
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdatePositions(Optional ByVal strSymbols As String = "")
On Error GoTo ErrSection:

    If m.nLoginMode = eGDTransActLoginMode_Demo Then
        InfBox "You cannot request positions because you are logged in as DEMO mode", "!", , "Positions Update Error"
        DumpDebug "You cannot request positions because you are logged in as DEMO mode"
    ElseIf m.nLoginMode = eGDTransActLoginMode_SimLive Then
        InfBox "You cannot request positions because you are logged in as SIMULATED LIVE mode", "!", , "Positions Update Error"
        DumpDebug "You cannot request positions because you are logged in as SIMULATED LIVE mode"
    Else
        If Not Refreshing Then
            If Len(strSymbols) = 0 Then DumpDebug "Updating Positions" Else DumpDebug "Updating Positions for " & strSymbols
            m.bRefreshingPositions = True
            SendTransactMessage eGDTransActMessageType_UpdatePositions, m.strAccountNumber & vbTab & strSymbols
        Else
            DumpDebug "Positions update requested but already in progress"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.UpdatePositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAccount
'' Description: Ask TransAct for the orders, trades, and positions to refresh
'' Inputs:      Update Positions?, Force Refresh?, Verbose Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshAccount(Optional ByVal bUpdatePositions As Boolean = True, Optional bForce As Boolean = False, Optional ByVal bVerbose As Boolean = False)
On Error GoTo ErrSection:

    If m.nLoginMode = eGDTransActLoginMode_Demo Then
        InfBox "You cannot refresh account because you are logged in as DEMO mode", "!", , "Refresh Account Error"
        DumpDebug "You cannot refresh account because you are logged in as DEMO mode"
    ElseIf m.nLoginMode = eGDTransActLoginMode_SimLive Then
        InfBox "You cannot refresh account because you are logged in as SIMULATED LIVE mode", "!", , "Refresh Account Error"
        DumpDebug "You cannot refresh account because you are logged in as SIMULATED LIVE mode"
    ElseIf (ConnectionStatus = eGDConnectionStatus_Connected) And ((Refreshing = False) Or (bForce = True)) Then
        DumpDebug "Refreshing Account " & m.strAccountNumber
        m.bRefreshingOrders = True
        m.bRefreshingFills = True
        m.bRefreshingPositions = True
        
        m.bVerboseRefresh = bVerbose
        
        If m.bVerboseRefresh Then
            InfBox "Synchronizing Orders, Fills, and Positions with " & m.strBrokerName & " servers.  Please wait...", , , m.strBrokerName & " Refresh", True
        End If
        
        SendTransactMessage eGDTransActMessageType_GetOrders, m.strAccountNumber
        SendTransactMessage eGDTransActMessageType_GetTrades, m.strAccountNumber
        If bUpdatePositions Then
            SendTransactMessage eGDTransActMessageType_UpdatePositions, m.strAccountNumber
        Else
            SendTransactMessage eGDTransActMessageType_GetPositions, m.strAccountNumber
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.RefreshAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResyncSubscriptionList
'' Description: Synchronize the subscription list with the GenTransact program
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
    
    DumpDebug "Resynchronizing Subscription List"
    
    If (ConnectionStatus = eGDConnectionStatus_Connected) Then
        Set astrSymbols = g.RealTime.BrokerSubscriptionList(m.nBroker)
        If Not astrSymbols Is Nothing Then
            astrSymbols.Sort
            
            ' Walk through real-time list and subscribe to any symbols that aren't already
            ' subscribed to...
            strSymbols = ""
            For lIndex = 0 To astrSymbols.Size - 1
                strSymbol = TransactSymbol(astrSymbols(lIndex))
                If m.astrSubContract.BinarySearch(strSymbol) = False Then
                    strSymbols = strSymbols & strSymbol & ","
                End If
            Next lIndex
            If Len(strSymbols) > 0 Then
                If Right(strSymbols, 1) = "," Then strSymbols = Left(strSymbols, Len(strSymbols) - 1)
                If Len(strSymbols) > 0 Then
                    DumpDebug "Subscribing to " & strSymbols
                    SendTransactMessage eGDTransActMessageType_Subscribe, m.strAccountNumber & vbTab & strSymbols
                End If
            End If
            
            ' Walk through subscribed list and unsubscribe to any symbols that are no longer
            ' in the real-time list...
            strSymbols = ""
            For lIndex = 0 To m.astrSubContract.Size - 1
                strSymbol = GenesisSymbol(m.astrSubContract(lIndex))
                If astrSymbols.BinarySearch(strSymbol) = False Then
                    strSymbols = strSymbols & m.astrSubContract(lIndex) & ","
                End If
            Next lIndex
            If Len(strSymbols) > 0 Then
                If Right(strSymbols, 1) = "," Then strSymbols = Left(strSymbols, Len(strSymbols) - 1)
                If Len(strSymbols) > 0 Then
                    DumpDebug "Unsubscribing from " & strSymbols
                    SendTransactMessage eGDTransActMessageType_Unsubscribe, m.strAccountNumber & vbTab & strSymbols
                End If
            End If
        
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.ResyncSubscriptionList"
    
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
            If strAccountNumber = m.strAccountNumber Then
                nReturn = eGDConnectionStatus_Connected
            Else
                nReturn = eGDConnectionStatus_Disconnected
            End If
            
    End Select
    
    ConnectionStatusForAccount = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransAct.ConnectionStatusForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextGenesisID
'' Description: Determine the next unique Genesis ID for the given account
'' Inputs:      Account Number
'' Returns:     Next Unique Genesis ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextGenesisID(ByVal strAccountNumber As String) As Long
On Error GoTo ErrSection:

    Dim lDate As Long                   ' Date from the ini file
    Dim lCounter As Long                ' Counter from the ini file
    Dim lCustomerID As Long             ' Customer ID from the registry
    
    lDate = GetIniFileProperty("Date", 0&, "ID", m.strIniFile)
    lCounter = GetIniFileProperty("Counter", 1&, "ID", m.strIniFile)
    lCustomerID = (RI_GetLastDataServiceID \ 1000) Mod 1000000
    
    If lDate <> Date Then
        lDate = Date
        lCounter = 1&
    Else
        lCounter = lCounter + 1
    End If

    SetIniFileProperty "Date", lDate, "ID", m.strIniFile
    SetIniFileProperty "Counter", lCounter, "ID", m.strIniFile
    
    NextGenesisID = -1 * CLng(Val(Format(lDate Mod 10000, "0000") & Format(lCounter, "00000")))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.NextGenesisID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowConfig
'' Description: Determine whether we need to show the broker config form or not
'' Inputs:      None
'' Returns:     True if we need to show config, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowConfig() As Boolean
On Error GoTo ErrSection:

    If Len(m.strUserName) = 0 Or Len(m.strIP) = 0 Then
        ShowConfig = True
    Else
        ShowConfig = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.ShowConfig"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckHeartbeat
'' Description: Check the heartbeat to make sure that we are still connected
''              to the TransAct servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckHeartbeat()
On Error GoTo ErrSection:
    
    If g.bUnloading = False Then
        If m.dHeartBeat <> 0 Then
            If (m.dHeartBeat < (gdTickCount - 120000)) Then
                If m.lReconnectAttempts = 1& Then
                    InfBox "Communication has been lost with " & m.strBrokerName & ".  Attempting to reconnect...", , "+-OK", m.strBrokerName & " Connection", True
                    DumpDebug "Communication has been lost with " & m.strBrokerName & ".  Attempting to reconnect..."
                End If
                DumpDebug "Reconnecting: Heartbeat is " & Str(gdTickCount - m.dHeartBeat) & " old"
                Reconnect
            ElseIf (ConnectionStatus = eGDConnectionStatus_Connected) Then
                If m.lReconnectAttempts > 1& Then
                    InfBox "TradeNavigator has reconnected to " & m.strBrokerName & ".", , "+-OK", m.strBrokerName & " Connection", True
                    DumpDebug "TradeNavigator has reconnected to " & m.strBrokerName & "."
                End If
                m.lReconnectAttempts = 0&
            End If
            
        ElseIf (ConnectionStatus = eGDConnectionStatus_Connecting) And (StartedProcess > 0#) Then
            If StartedProcess < gdTickCount - 15000 Then
                DumpDebug "Timed out waiting for application to load"
                ConnectionStatus = eGDConnectionStatus_Disconnected
            
                If KillProcess(m.strProcessName) > 0 Then
                    DumpDebug "Process '" & m.strProcessName & "' killed"
                End If
            
                StartedProcess = 0#
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.CheckHeartbeat"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reconnect
'' Description: Attempt to reconnect to the TransAct servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Reconnect()
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Time out variable

    If m.bReconnecting = False Then
        DumpDebug "Reconnecting"
        m.bReconnecting = True
        
        If ConnectionStatus = eGDConnectionStatus_Connected Then
            Disconnect "", True, "Reconnecting"
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
    RaiseError "cTransact.Reconnect"
    
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
    RaiseError "cTransact.DumpDebug"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming App Mail message from the TransAct program
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDTransActMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strBaseSymbol As String         ' Translated base symbol
    Dim strSymbol As String             ' Translated symbol
    Dim astrSymbols As New cGdArray     ' List of symbols subscribed to
    Dim strMessage2 As String           ' Message to send
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position in an array
    Dim strKey As String                ' Key into the registry
    Dim strAccount As String            ' Account
    Dim lError As Long                  ' Error code
    Dim frm As frmAlertPopup            ' Alert popup form
    Dim bReloadAccount As Boolean       ' Do we want to reload the account?
    Static bOrdersDone As Boolean       ' Have we received all of the orders?
    Static bTradesDone As Boolean       ' Have we received all of the trades?
    
    If g.bUnloading Then Exit Sub
    
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"

    If (nType <> eGDTransActMessageType_PriceUpdate) And (nType <> eGDTransActMessageType_ConnectionInfo) And (nType <> eGDTransActMessageType_Heartbeat) Then
        DumpDebug vbTab & m.strBrokerName & " Message Received (" & MessageTypeToString(nType) & "): " & strMessage
    End If

    Select Case nType
        Case TransactMessageType(eGDTransActMessageType_Subscribed)
            m.dHeartBeat = gdTickCount
            DumpDebug "Subscribed to: " & Parse(strMessage, vbTab, 2)
            astrSymbols.SplitFields Parse(strMessage, vbTab, 2), ";"
            For lIndex = 0 To astrSymbols.Size - 1
                If Parse(strMessage, vbTab, 3) = "0" Then
                    If m.astrSubContract.BinarySearch(astrSymbols(lIndex), lPos) = False Then
                        m.astrSubContract.Add astrSymbols(lIndex), lPos
                    End If
                End If
                
                strBaseSymbol = Parse(GenesisSymbol(astrSymbols(lIndex)), "-", 1)
                If Len(strBaseSymbol) > 0 Then
                    g.RealTime.AddBrokerRtSymbol strBaseSymbol & "-", "Transact/D"
                    
                    If m.astrEnabledSymbols.BinarySearch(strBaseSymbol, lPos) = False Then
                        m.astrEnabledSymbols.Add strBaseSymbol, lPos
                    End If
                End If
            Next lIndex
            
            SaveEnabledSymbols
            
            If Parse(strMessage, vbTab, 3) = "1" Then
                g.RealTime.UseBrokerFeed = m.nBroker
                ResyncSubscriptionList
                m.strAskedFor = Parse(strMessage, vbTab, 2) & ";"
            ElseIf Len(m.strAskedFor) > 0 Then
                If InStr(m.strAskedFor, Parse(strMessage, vbTab, 2) & ";") <> 0 Then
                    m.strAskedFor = Replace(m.strAskedFor, Parse(strMessage, vbTab, 2) & ";", "")
                End If
                If Len(m.strAskedFor) = 0 Then ResyncSubscriptionList
            End If
            
        Case TransactMessageType(eGDTransActMessageType_Unsubscribed)
            m.dHeartBeat = gdTickCount
            DumpDebug "Unsubscribed from: " & Parse(strMessage, vbTab, 2)
            astrSymbols.SplitFields Parse(strMessage, vbTab, 2), ";"
            For lIndex = 0 To astrSymbols.Size - 1
                If m.astrSubContract.BinarySearch(astrSymbols(lIndex), lPos) = True Then
                    m.astrSubContract.Remove lPos
                End If
            Next lIndex
            
        Case TransactMessageType(eGDTransActMessageType_CancelledOrder)
            m.dHeartBeat = gdTickCount
            'UpdateBrokerInfoOrder strMessage, False
            HandleOrderCallback nType, strMessage, False
            
        Case TransactMessageType(eGDTransActMessageType_ChangedOrder)
            m.dHeartBeat = gdTickCount
            'UpdateBrokerInfoOrder strMessage, False
            HandleOrderCallback nType, strMessage, False
            
        Case TransactMessageType(eGDTransActMessageType_FilledOrder)
            m.dHeartBeat = gdTickCount
            'UpdateBrokerInfoOrder strMessage, False
            HandleOrderCallback nType, strMessage, False
            
        Case TransactMessageType(eGDTransActMessageType_ExpiredOrder)
            m.dHeartBeat = gdTickCount
            'UpdateBrokerInfoOrder strMessage, False
            HandleOrderCallback nType, strMessage, False
            
        Case TransactMessageType(eGDTransActMessageType_SentOrder)
            m.dHeartBeat = gdTickCount
            'UpdateBrokerInfoOrder strMessage, False
            HandleOrderCallback nType, strMessage, False
            
        Case TransactMessageType(eGDTransActMessageType_GetOrder)
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingOrders = True
                m.BrokerInfo.RefreshingFills = True
                m.BrokerInfo.ClearOrders
            ElseIf UCase(strMessage) = "END" Then
                bOrdersDone = True
                m.bRefreshingOrders = False
                
                If bOrdersDone And bTradesDone Then
                    m.BrokerInfo.RefreshingOrders = False
                    m.BrokerInfo.RefreshingFills = False
                    
                    bOrdersDone = False
                    bTradesDone = False
                    m.bRefreshing = False
                    InfBox ""
                End If
            Else
                'UpdateBrokerInfoOrder strMessage, True
                HandleOrderCallback nType, strMessage, True
            End If
            DoEvents
            
        Case TransactMessageType(eGDTransActMessageType_GetTrade)
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.ClearFills
                m.BrokerInfo.RefreshingOrders = True
                m.BrokerInfo.RefreshingFills = True
            ElseIf UCase(strMessage) = "END" Then
                bTradesDone = True
                m.bRefreshingFills = False
                
                If bOrdersDone And bTradesDone Then
                    m.BrokerInfo.RefreshingOrders = False
                    m.BrokerInfo.RefreshingFills = False
                    
                    bOrdersDone = False
                    bTradesDone = False
                    m.bRefreshing = False
                    InfBox ""
                End If
            Else
                'UpdateBrokerInfoOrder strMessage, True
                HandleOrderCallback nType, strMessage, True
            End If
            DoEvents
            
        Case TransactMessageType(eGDTransActMessageType_AppLoaded)
            m.bAppLoaded = True
            StartedProcess = 0
            'If LiveTradingAllowed(m.nBroker) Then
            '    DumpDebug "Application Loaded -- Connecting"
            '    strMessage2 = m.strUserName & vbTab & m.strPassword & vbTab & m.strIP & vbTab & m.strAccountNumber
            '    SendTransactMessage eGDTransActMessageType_Connect, strMessage2
            'Else
            '    DumpDebug "Unloading the application since live trading is not allowed"
            '    SendTransactMessage eGDTransactMessageType_UnloadApp, ""
            'End If
        
        Case TransactMessageType(eGDTransActMessageType_AppUnloaded)
            m.bAppLoaded = False
            DumpDebug "Application Unloaded"
            If m.bManualDisconnect = True Then
                m.dHeartBeat = 0#
                m.bVerifyPositions = True
            ElseIf m.lNumSetAccountFailures > 0 Then
                m.dHeartBeat = gdTickCount - 60000
            Else
                m.dHeartBeat = gdTickCount - 110000
            End If
            m.bManualDisconnect = False
            If FormIsLoaded("frmTransActLogin") Then
                frmTransActLogin.AppUnloaded
            End If
        
        Case TransactMessageType(eGDTransActMessageType_Heartbeat)
            m.dHeartBeat = gdTickCount
            If m.dLastFillTime > 0 Then
                If gdTickCount > m.dLastFillTime + 5000 Then
                    GetPositions
                    m.dLastFillTime = 0#
                End If
            End If
            
        Case TransactMessageType(eGDTransActMessageType_PriceUpdate)
            m.dHeartBeat = gdTickCount
            UpdatePrice strMessage
            
        Case TransactMessageType(eGDTransActMessageType_Error)
            strAccount = Parse(strMessage, vbTab, 1)
            lError = CLng(Val(Parse(strMessage, vbTab, 2)))
            
            ''If (strAccount <> kTransActSimUserAccount) And (strAccount <> kTransActOldSimUserAccount) Then
            If (Not TransActSimulatedAccount(strAccount)) And (m.nLoginMode = eGDTransActLoginMode_Live) Then
                Select Case lError
                    Case 1
                        If (ConnectionStatus = eGDConnectionStatus_Disconnected) And (m.bSuccessfullyConnected = False) Then
                            If Len(Parse(strMessage, vbTab, 1)) > 0 Then
                                Set frm = New frmAlertPopup
                                frm.ShowMessageBox "Error: " & Parse(strMessage, vbTab, 2) & "|Account: " & Parse(strMessage, vbTab, 1) & ":|" & Parse(strMessage, vbTab, 3) & "|", m.strBrokerName & " Error", vbLeftJustify
                            Else
                                Set frm = New frmAlertPopup
                                frm.ShowMessageBox "Error: " & Parse(strMessage, vbTab, 2) & ":|" & Parse(strMessage, vbTab, 3) & "|", m.strBrokerName & " Error", vbLeftJustify
                            End If
                            
                            InfBox ""

                            m.dHeartBeat = 0#
                            m.bVerifyPositions = True
                        End If
                    
                    Case 13
                    
                    Case 1004
                        InfBox ""
                        Set frm = New frmAlertPopup
                        frm.ShowMessageBox "There are no accounts for user '" & m.strUserName & "' on the " & m.strBrokerName & " server '" & m.strIP & "'.  Please call your broker and verify your user name and account information.", m.strBrokerName & " Error", vbCenter
                        Disconnect "", False, "No Accounts"
                        
                    Case 1008
                        InfBox ""
                        Set frm = New frmAlertPopup
                        frm.ShowMessageBox "There are no symbols set up for account '" & m.strAccountNumber & "' on the " & m.strBrokerName & " server '" & m.strIP & "'.  Please call your broker and verify your user name and account information.", m.strBrokerName & " Error", vbCenter
                        Disconnect "", False, "No Symbols for Account"
                    
                    Case 1009           ' Subscribe failed
                        If InStr(UCase(Parse(strMessage, vbTab, 5)), "NEVER BE SEEN") <> 0 Then
                            Disconnect "", True, "Never be seen error"
                        Else
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox "Error: " & Parse(strMessage, vbTab, 4) & "|Account: " & Parse(strMessage, vbTab, 1) & ":|" & Parse(strMessage, vbTab, 3) & "|" & Parse(strMessage, vbTab, 5) & "|", m.strBrokerName & " Error", vbLeftJustify
                        End If
                    
                    Case 1005, 1006, 1007
                        HandleOrderError strMessage
                        
                    Case Else
                        If Len(Parse(strMessage, vbTab, 1)) > 0 Then
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox "Error: " & Parse(strMessage, vbTab, 2) & "|Account: " & Parse(strMessage, vbTab, 1) & ":|" & Parse(strMessage, vbTab, 3) & "|", m.strBrokerName & " Error", vbLeftJustify
                        Else
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox "Error: " & Parse(strMessage, vbTab, 2) & ":|" & Parse(strMessage, vbTab, 3) & "|", m.strBrokerName & " Error", vbLeftJustify
                        End If
                End Select
                
                Select Case lError
                    Case 13
                        m.bAskedPassword = False
                        m.strPassword = ""
                        Connect "", True
                    
                    Case 56, 57, 60, 75
                        m.bAskedPassword = False
                        m.strPassword = ""
                        
                        m.bManualDisconnect = True
                        m.dHeartBeat = 0#
                        m.dLastStatusMsg = 0#
                        
                    Case 1002
                        ' This number starts at zero, so when it gets to 4, it has failed 5 times...
                        If m.lNumSetAccountFailures < 4 Then
                            m.dHeartBeat = gdTickCount - 60000
                            m.lNumSetAccountFailures = m.lNumSetAccountFailures + 1
                        Else
                            m.bManualDisconnect = True
                            m.dLastStatusMsg = 0#
                            m.dHeartBeat = 0#
                            m.bVerifyPositions = True
                            m.bSuccessfullyConnected = False
                            m.lReconnectAttempts = -1&
                            m.lNumSetAccountFailures = 0
                            
                            DumpDebug "Not reconnecting anymore -- 5 set account failures"
                        End If
                        
                End Select
            Else
                Select Case lError
                    Case 13, 56, 57, 60, 75
                        m.bManualDisconnect = True
                        m.dHeartBeat = 0#
                        m.dLastStatusMsg = 0#
                        Set frm = New frmAlertPopup
                        frm.ShowMessageBox Parse(strMessage, vbTab, 3), m.strBrokerName & " Error " & Str(lError), vbCenter
                        
                End Select
            End If
            
        Case TransactMessageType(eGDTransActMessageType_Account)
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingAccounts = True
                m.BrokerInfo.ClearAccounts
            ElseIf UCase(strMessage) = "END" Then
                m.BrokerInfo.RefreshingAccounts = False
                If m.nLoginMode = eGDTransActLoginMode_Live Then
                    RefreshAccount False, True
                End If
                m.bRefreshingAccounts = False
            Else
                bReloadAccount = CreateAccount(Parse(strMessage, vbTab, 1), Parse(strMessage, vbTab, 2))
                If Parse(strMessage, vbTab, 1) = m.strAccountNumber Then
                    m.BrokerInfo.AddAccount Parse(strMessage, vbTab, 1) & vbTab, True, bReloadAccount
                    m.BrokerInfo.SetConnectionStatusForAccount eGDConnectionStatus_Connected, Parse(strMessage, vbTab, 1)
                Else
                    m.BrokerInfo.AddAccount Parse(strMessage, vbTab, 1) & vbTab, False, bReloadAccount
                    m.BrokerInfo.SetConnectionStatusForAccount eGDConnectionStatus_Disconnected, Parse(strMessage, vbTab, 1)
                End If
            End If
            
        Case TransactMessageType(eGDTransActMessageType_Position)
            m.dHeartBeat = gdTickCount
            AddPositionToBrokerInfo strMessage, False
            strSymbol = GenesisSymbol(Parse(strMessage, vbTab, 3))
            If Len(strSymbol) > 0 Then
                SynchronizePositions strSymbol
            End If
            
        Case TransactMessageType(eGDTransActMessageType_ConnectionInfo)
            UpdateConnectionInfo strMessage
            For lIndex = 0 To m.astrLastFillTime.Size - 1
                If Val(Parse(m.astrLastFillTime(lIndex), vbTab, 2)) > 0 Then
                    If gdTickCount > Val(Parse(m.astrLastFillTime(lIndex), vbTab, 2)) + 5000 Then
                        UpdatePositions Parse(m.astrLastFillTime(lIndex), vbTab, 1)
                        m.astrLastFillTime(lIndex) = Parse(m.astrLastFillTime(lIndex), vbTab, 1) & vbTab & "0"
                    End If
                End If
            Next lIndex
            
        Case TransactMessageType(eGDTransActMessageType_GetPosition)
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.ClearPositions
                m.BrokerInfo.RefreshingPositions = True
                m.astrBrokerPos.Clear
            ElseIf UCase(strMessage) = "END" Then
                m.BrokerInfo.RefreshingPositions = False
                SynchronizePositions
                m.bRefreshingPositions = False
                
                InfBox ""
                m.bVerboseRefresh = False
            Else
                AddPositionToBrokerInfo strMessage, True
            End If
            
        Case TransactMessageType(eGDTransActMessageType_AccountList)
            If FormIsLoaded("frmTransActLogin") Then
                frmTransActLogin.SetAccounts strMessage
            ElseIf Len(m.strUserName) > 0 Then
                SetIniFileProperty "AccountList", Replace(strMessage, vbTab, ","), m.strUserName, m.strIniFile
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
    RaiseError "cTransact.HandleMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendTransactMessage
'' Description: Send an App Mail message to the TransAct program
'' Inputs:      Type of Message, Message, Send Now?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendTransactMessage(ByVal nType As eGDTransActMessageTypes, ByVal strMessage As String, Optional ByVal bSendNow As Boolean = True)
On Error GoTo ErrSection:

    Dim astrMessage As New cGdArray     ' Message to dump to the debug

    If nType = eGDTransActMessageType_Connect Then
        astrMessage.SplitFields strMessage, vbTab
        astrMessage(1) = SetStrToChar(astrMessage(1), "*")
        DumpDebug vbTab & "Sending " & m.strBrokerName & " Message (" & MessageTypeToString(nType) & "): " & astrMessage.JoinFields(vbTab)
    ElseIf (nType = eGDTransActMessageType_LogonForAccountList) Or (nType = eGDTransActMessageType_Logon) Then
        astrMessage.SplitFields strMessage, vbTab
        astrMessage(1) = SetStrToChar(astrMessage(1), "*")
        astrMessage(2) = SetStrToChar(astrMessage(2), "*")
        DumpDebug vbTab & "Sending " & m.strBrokerName & " Message (" & MessageTypeToString(nType) & "): " & astrMessage.JoinFields(vbTab)
    Else
        DumpDebug vbTab & "Sending " & m.strBrokerName & " Message (" & MessageTypeToString(nType) & "): " & strMessage
    End If
    frmOnlineBroker.gdBroker.CreateMessage m.strControlID, nType, strMessage, , bSendNow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.SendTransactMessage"
    
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
    
    If UCase(m.strUserName) = UCase(SimUserUserName) Then
        m.bVerifyPositions = False
    ElseIf ((m.bVerifyPositions = True) Or (bForce = True)) And (m.bFixCalledOnce = True) Then
        ' Only do this if there is not a modal dialog up...
        If frmMain.Enabled Then
            m.bVerifyPositions = False
            
            If m.nLoginMode = eGDTransActLoginMode_Live Then
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
    RaiseError "cTransact.VerifyPositions"
    
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

    IsTradeableSymbol = (Len(TransactSymbol(strGenesisSymbol)) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.IsTradeableSymbol"
    
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
    Dim strTransActSymbol As String     ' TransAct symbol
    
    strTransActSymbol = TransactSymbol(strGenesisSymbol)
    
    bReturn = (Len(strTransActSymbol) > 0)
    If bReturn Then
        strBrokerBase = Parse(strTransActSymbol, "|", 1)
        strBrokerExchange = ""
        
        bReturn = m.astrEnabledSymbols.BinarySearch(Parse(strGenesisSymbol, "-", 1))
    End If
    
    IsEnabledSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.IsEnabledSymbol"
    
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
    Dim strTransActSymbol As String     ' TransAct symbol for the Genesis symbol passed in
    Dim strBaseSymbol As String         ' TransAct base symbol
    Dim lPos As Long                    ' Position in the base symbol array
    Dim strMask As String               ' Order type mask from the file
    
    bReturn = False
    
    If nOrderType <> eTT_OrderType_Adjustment Then
        strTransActSymbol = TransactSymbol(strGenesisSymbol)
        If Len(strTransActSymbol) > 0 Then
            strBaseSymbol = Parse(strTransActSymbol, "|", 1)
            If m.astrBaseSyms.BinarySearch(strBaseSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
                strMask = Parse(m.astrBaseSyms(lPos), vbTab, 5)
                If Len(strMask) = 0 Then strMask = "11110000"
                If Len(strMask) = 8 Then strMask = strMask & "0"
                If (nOrderType >= 0) And (nOrderType < Len(strMask)) Then
                    bReturn = (Mid(strMask, nOrderType + 1, 1) = "1")
                End If
            End If
        End If
    End If
    
    OrderTypeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.OrderTypeAllowed"
    
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
    Dim strTransActSymbol As String     ' TransAct symbol for the Genesis symbol passed in
    Dim strBaseSymbol As String         ' TransAct base symbol
    Dim lPos As Long                    ' Position in the base symbol array
    Dim strMask As String               ' Order type mask from the file
    
    bReturn = False
    
    strTransActSymbol = TransactSymbol(strGenesisSymbol)
    If Len(strTransActSymbol) > 0 Then
        If nTimeInForce = eTT_TimeInForce_GTC Then
            bReturn = m.bAllowGTC
        Else
            strBaseSymbol = Parse(strTransActSymbol, "|", 1)
            If m.astrBaseSyms.BinarySearch(strBaseSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
                strMask = Parse(m.astrBaseSyms(lPos), vbTab, 6)
                If Len(strMask) = 0 Then strMask = "100"
                If (nTimeInForce >= 0) And (nTimeInForce < Len(strMask)) Then
                    bReturn = (Mid(strMask, nTimeInForce + 1, 1) = "1")
                End If
            End If
        End If
    End If
    
    TimeInForceAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.TimeInForceAllowed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DropOffTime
'' Description: Determine what time a symbol is likely to fall out of a refresh
''              because TransAct has done their daily maintenance on it
'' Inputs:      Symbol
'' Returns:     Drop Off Time
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DropOffTime(ByVal strGenesisSymbol As String) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim strTransActSymbol As String     ' TransAct symbol
    Dim strBaseSymbol As String         ' TransAct base symbol
    Dim lPos As Long                    ' Position in the base symbol array
    Dim dCurrentTime As Double          ' Current time
    
    strTransActSymbol = TransactSymbol(strGenesisSymbol)
    If Len(strTransActSymbol) > 0 Then
        strBaseSymbol = Parse(strTransActSymbol, "|", 1)
        dCurrentTime = CurrentTime("CHI")
        
        If m.astrBaseSyms.BinarySearch(strBaseSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
            dReturn = Int(dCurrentTime) + (Val(Parse(m.astrBaseSyms(lPos), vbTab, 8)) / 1440#)
        Else
            dReturn = Int(dCurrentTime) + (((15 * 60) + 20) / 1440#)
        End If
    End If
    
    DropOffTime = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.DropOffTime"
    
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
    RaiseError "cTransact.FixPositions"
    
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
                            frm.ShowMessageBox m.strBrokerName & " is reporting that you carried a " & UCase(g.Broker.TextPosition(lBrokerCarried)) & " position for " & strSymbol & " in account " & strAccount & ", but Trade Navigator reports a carried position of " & UCase(FillSummary.CurrentPositionString) & ".||Because this inconsistency could cause incorrect orders to be placed, auto exits and automated trading strategies are being disabled for this symbol.|", "Inconsistent Carried Position", vbCenter
                            
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
    RaiseError "cTransact.FixPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ATVersion
'' Description: Attempt to determine the version of TransAct AT installed
'' Inputs:      None
'' Returns:     TransAct AT Version (-999999 if not installed)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ATVersion() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    
    dReturn = kNullData
    If Len(GetRegistryValue(rkLocalMachine, kTransActRegistryKey, "AppName", "")) > 0 Then
        m.strVersion = GetRegistryValue(rkLocalMachine, kTransActRegistryKey, "Version", "")
        If Len(m.strVersion) = 0 Then
            dReturn = 0#
        ElseIf Left(m.strVersion, 1) = "1" Then
            dReturn = 1#
        ElseIf Left(m.strVersion, 1) = "5" Then
            dReturn = 5#
        End If
    End If
    
    ATVersion = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransAct.ATVersion"

End Function

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

    Dim lIndex As Long                  ' Index into a for loop
    Dim lLastDate As Long               ' Last successful connection to broker
    
    If g.Broker.IsBrokerUser(m.nBroker) Then
        lLastDate = LastDateConnected
        
        If g.Broker.IsBrokerSimUser(m.nBroker) Or (lLastDate >= Date - 30) Then
            DumpDebug "Symbols enabled for real-time: '" & m.astrEnabledSymbols.JoinFields(";") & "'"
            For lIndex = 0 To m.astrEnabledSymbols.Size - 1
                g.RealTime.AddBrokerRtSymbol m.astrEnabledSymbols(lIndex) & "-", m.strBrokerName & "/D"
            Next lIndex
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.AddBrokerRtSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateAccount
'' Description: If the given account does not exist, create it
'' Inputs:      Account Number, Account Name
'' Returns:     Was Account Created or Changed?
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateAccount(ByVal strAccountNumber As String, ByVal strAccountName As String) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lAccountID As Long              ' Account ID of the account just created
    Dim bReturn As Boolean              ' Return value for the function
    Dim bOverwriteName As Boolean       ' Do we want to overwrite the account name?
    
    bReturn = False
    
    ' If the account number is the simulated account or a demo account, change the name
    ' to something a little bit more descriptive to distinguish it from live accounts...
    If CLng(Val(strAccountNumber)) = kTransActSimUserAccount Then
        strAccountName = "TransAct Simulator"
        bOverwriteName = True
    ElseIf TransActSimulatedAccount(strAccountNumber) Then
        strAccountName = "TransAct Demo " & strAccountNumber
        bOverwriteName = True
    Else
        bOverwriteName = False
    End If
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountNumber]='" & strAccountNumber & "';", dbOpenDynaset)
    If (rs.BOF And rs.EOF) Then
        DumpDebug "Creating account " & strAccountNumber
        rs.AddNew
        rs!AccountNumber = strAccountNumber
        rs!Name = strAccountName
        rs!StartingBalance = 0
        rs!CurrentBalance = 0
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
        bReturn = True
    Else
        lAccountID = rs!AccountID
        If rs!UserName <> m.strUserName Then
            rs.Edit
            rs!UserName = m.strUserName
            EncryptField rs!Password, m.strPassword
            rs.Update
            bReturn = True
        ElseIf (rs!Name <> strAccountName) And (bOverwriteName = True) Then
            rs.Edit
            rs!Name = strAccountName
            rs.Update
            bReturn = True
        End If
    End If
    
    mTradeTracker.RefreshAccount lAccountID
    
    CreateAccount = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.CreateAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Convert a TransAct symbol to a Genesis symbol
'' Inputs:      TransAct Symbol
'' Returns:     Genesis Symbol (Blank if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisSymbol(ByVal strTransActSymbol As String) As String
On Error GoTo ErrSection:

    Dim strBase As String               ' Base symbol
    Dim lPos As Long                    ' Position of base symbol in the array
    Dim strReturn As String             ' Return from the function
    
    strBase = Parse(strTransActSymbol, "|", 1)
    strReturn = ""
    
    If m.astrBaseSyms.BinarySearch(strBase & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        strReturn = Parse(m.astrBaseSyms(lPos), vbTab, 3) & "-" & Parse(strTransActSymbol, "|", 2) & Format(Parse(strTransActSymbol, "|", 3), "00")
    End If

    GenesisSymbol = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.GenesisSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactSymbol
'' Description: Convert a Genesis symbol to a TransAct symbol
'' Inputs:      Genesis Symbol
'' Returns:     TransAct Symbol (Blank if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TransactSymbol(ByVal strGenesisSymbol As String) As String
On Error GoTo ErrSection:

    Dim strBase As String               ' Base symbol
    Dim strContract As String           ' Contract
    Dim strReturn As String             ' Return from the function
    Dim lIndex As Long                  ' Index into a for loop
    
    strBase = Parse(strGenesisSymbol, "-", 1)
    strContract = Parse(strGenesisSymbol, "-", 2)
    strReturn = ""
    
    For lIndex = 0 To m.astrBaseSyms.Size - 1
        If Parse(m.astrBaseSyms(lIndex), vbTab, 3) = strBase Then
            strReturn = Parse(m.astrBaseSyms(lIndex), vbTab, 1) & "|" & Left(strContract, 4) & "|" & Right(strContract, 2)
            Exit For
        End If
    Next lIndex

    TransactSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.TransactSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleOrderCallback
'' Description: Handle an order callback from the TransAct app
'' Inputs:      Message Type, Message
'' Returns:     None
''
'' Fields:      Error, Account, Symbol, Expiration, Fill Price, Fill Qty,
''              Fill Time, B/S, O/C, Last Updated, Order ID, Order Number,
''              Order Type, Original Order ID, Original Qty, Limit Price,
''              Quantity, Remaining Qty, Restrictions, Order Status, Stop Price,
''              Time Sent, Fill ID, Yes ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleOrderCallback(ByVal nType As eGDTransActMessageTypes, ByVal strMessage As String, Optional ByVal bRefresh As Boolean = False)
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Order split out in fields
    Dim Order As cPtOrder               ' Order object
    Dim lFillID As Long                 ' ID of the fill
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As New cPtFill             ' Temporary fill object
    Dim bReceivedFill As Boolean        ' Did we receive a fill with the order?
    Dim bFillExisted As Boolean         ' Did the fill already exist?
    Dim bFillChanged As Boolean         ' Did the fill change?
    Dim nPrevStatus As eTT_OrderStatus  ' Previous order status
    Dim bModified As Boolean            ' Is this a modified order?
    Dim lOldOrderID As Long             ' Old order id if it was modified
    Dim dFillTime As Double             ' Old Fill Time
    Dim dFillPrice As Double            ' Old Fill Price
    Dim lFillQty As Long                ' Old Fill Quantity
    Dim lPos As Long                    ' Position in an array
    Dim frm As frmAlertPopup            ' Alert popup form
    Dim Bars As New cGdBars             ' Bars object
    Dim nOrderType As eTT_OrderType     ' Order Type
    Dim PrevOrder As cPtOrder           ' Previous order
    
    astrOrder.SplitFields strMessage, vbTab

    ''If (astrOrder(1) <> kTransActSimUserAccount) And (astrOrder(1) <> kTransActOldSimUserAccount) Then
    If (Not TransActSimulatedAccount(astrOrder(1))) And (m.nLoginMode = eGDTransActLoginMode_Live) Then
        If Val(astrOrder(22)) < 0 Then
            Set Order = LoadOrder(astrOrder(1), astrOrder(22), astrOrder(10), astrOrder(13))
        Else
            Set Order = LoadOrder(astrOrder(1), "", astrOrder(10), astrOrder(13))
        End If

        If Order Is Nothing Then
            Set Order = GenesisOrder(astrOrder)
            If Order Is Nothing Then
                DumpDebug "Order cannot be created because of an unknown symbol"
            Else
                DumpDebug "Creating Order: '" & Order.OrderText(True, True, True) & "'"
            End If
        ElseIf (Order.Symbol <> GenesisSymbol(astrOrder(2))) Then
            Set Order = New cPtOrder
            Set Order = GenesisOrder(astrOrder)
            DumpDebug "Creating Order (different symbol): '" & Order.OrderText(True, True, True) & "'"
            
            Set frm = New frmAlertPopup
            frm.ShowMessageBox "An order has come from the " & m.strBrokerName & " servers that had the same identifier as an order in Trade Navigator with a different symbol.  Trade Navigator has created a new order based on the order that came from the " & m.strBrokerName & " servers.", m.strBrokerName & " Order Error", vbCenter
        Else
            DumpDebug "Found Order: '" & Order.OrderText(True, True, True) & "'"
        End If
                
        If Not Order Is Nothing Then
            nPrevStatus = Order.Status
            
            If astrOrder(0) = "0" Then
                With Order
                    bModified = False
                    SetBarProperties Bars, .Symbol
                    lOldOrderID = 0&
                    
                    If m.ModifiedOrders.Exists(astrOrder(13)) Then
                        bModified = True
                        lOldOrderID = .OrderID
                        
                        ' DAJ 04/25/2013: Only do this if the broker ID of the found order matches
                        ' the broker id of the order in the ModifiedOrders collection.  The thinking
                        ' here is that if the order id's don't match, we have already created the
                        ' new order...
                        If .BrokerID = m.ModifiedOrders(astrOrder(13)).BrokerID Then
                            .OrderID = 0&
                            .Fills.Clear
                            .History.Clear
                        End If
                        
                        .PreviousBrokerID = m.ModifiedOrders(astrOrder(13)).BrokerID
                        
                        DumpDebug "Order Modified: '" & m.ModifiedOrders(astrOrder(13)).OrderText(True, True, True) & "'"
                        m.ModifiedOrders(astrOrder(13)).GenesisOrderID = ""
                        m.ModifiedOrders(astrOrder(13)).Status = eTT_OrderStatus_Amended
                        m.ModifiedOrders(astrOrder(13)).StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, m.ModifiedOrders(astrOrder(13)).Symbol)
                        m.ModifiedOrders(astrOrder(13)).Save
                        
                        .PreviousOrder = m.ModifiedOrders(astrOrder(13))
                        
                        OrderCallback m.ModifiedOrders(astrOrder(13))
                        m.BrokerInfo.AddOrder m.ModifiedOrders(astrOrder(13)), bRefresh
                        g.ActivityLogs.AddOrderToActivityLog m.nBroker, m.ModifiedOrders(astrOrder(13))
                        
                        m.ModifiedOrders.Remove astrOrder(13)
                    ElseIf (nType = eGDTransActMessageType_ChangedOrder) And (.BrokerID = astrOrder(13)) Then
                        Set PrevOrder = Order.MakeCopy
                        bModified = True
                        lOldOrderID = .OrderID
                        
                        .OrderID = 0&
                        .Fills.Clear
                        .History.Clear
                        .PreviousBrokerID = PrevOrder.BrokerID
                        
                        DumpDebug "Order Modified: '" & PrevOrder.OrderText(True, True, True) & "'"
                        PrevOrder.GenesisOrderID = ""
                        PrevOrder.Status = eTT_OrderStatus_Amended
                        PrevOrder.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, PrevOrder.Symbol)
                        PrevOrder.Save
                        
                        .PreviousOrder = PrevOrder
                        
                        OrderCallback PrevOrder
                        m.BrokerInfo.AddOrder PrevOrder, bRefresh
                        g.ActivityLogs.AddOrderToActivityLog m.nBroker, PrevOrder
                    End If
                    
                    If astrOrder(10) <> "0" Then .BrokerID = astrOrder(10)
                    .Buy = (astrOrder(7) <> "0")
                    .ExchangeID = astrOrder(11)
                    nOrderType = GenesisOrderType(astrOrder(12))
                    .OrderType = GenesisOrderType(astrOrder(12))
                    .LimitPrice = GenesisPrice(astrOrder(2), astrOrder(15))
                    .StopPrice = GenesisPrice(astrOrder(2), astrOrder(20))
                    .Quantity = CLng(Val(astrOrder(16)))
                                                
                    ' 03/16/2010 DAJ: If we get a fill with a negative quantity (which is a special
                    ' case that the TransAct desk will do in rare cases), we need to change the quantity
                    ' to a positive quantity and reverse the Buy/Sell code...
                    ' 01/16/2014 DAJ: When we get one of these negative quantity situations, the
                    ' time stamp coming back to us is in Chicago time instead of GMT like all the
                    ' other orders are.  Change it here to make it consistent...
                    If .Quantity < 0 Then
                        .Quantity = .Quantity * -1&
                        .Buy = Not .Buy
                        .StatusDate = ConvertTimeZone(.StatusDate, "CHI", "GMT")
                    End If
                    
                    ' Restrictions: 1=None, 2=All or None, 3=Immediate, 5=GTC
                    If astrOrder(18) = "5" Then
                        .Expiration = 0&
                    ElseIf CLng(Val(astrOrder(3))) = 0& Then
                        .Expiration = -1&
                    Else
                        .Expiration = CLng(Val(astrOrder(3))) * -1&
                    End If
                    
                    .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.OrderDate, m.nBroker, .Symbol, False))
                    .IsSnapshot = True
                                        
                    ' If we don't have an Order ID yet (this is a new order), we need to save
                    ' the order now so that we can get an Order ID for the fill...
                    If .OrderID = 0& Then
                        .Save
                    End If
                    
                    ' DAJ 04/25/2013: Moved this out of the '.OrderID = 0&' block above because
                    ' there will be some times that we won't be blanking out the OrderID in the
                    ' modified blocks above...
                    If bModified Then
                        OrderIDChanged lOldOrderID, .OrderID
                    End If
                    
                    ' 11/12/2012 DAJ: TransAct made a change back at the beginning of the year that can cause
                    ' fill information to come through on a CancelledOrder message, but that fill information
                    ' has a different ID than the FilledOrder message causing me to create duplicate fills (and
                    ' therefore to create an "over filled" order.  The fix as per Jim at TransAct is to only
                    ' pay attention to a fill if it comes in a FilledOrder message...
                    ' 11/14/2012 DAJ: Realized that the fix made me skip the fill information all together in
                    ' a refresh situation ( eGDTransActMessageType_GetTrade ), so we will also handle the fill
                    ' information in a refersh now...
                    ' 12/20/2012 DAJ: Also going to ignore "fills" that come in with a zero quantity.  Looks like
                    ' if you try to cancel an order that has filled ( i.e. two ships passing in the night ),
                    ' a subsequent refresh can come through with an order that has a fill price, but a zero
                    ' quantity.  That causes us to create a zero quantity fill which creates an error in
                    ' cAccountPosition.CreateTradeLines...
                    ' 01/16/2014 DAJ: Need to allow a fill with a negative quantity through here, too.  That it
                    ' is a rare case where TransAct is correcting or "undo-ing" a fill...
                    ' 09/01/2015 DAJ: Also, only count is as a fill if the status is 3 ( filled ), 1000 ( balance
                    ' cancelled ), or 1001 ( partial filled ) because a customer is getting records back on a
                    ' refresh that look like fills, but are not...
                    lFillID = Val(astrOrder(22))
                    lFillQty = CLng(Val(astrOrder(5)))
                    If ((nType = eGDTransActMessageType_FilledOrder) Or (bRefresh = True)) And (lFillID > 0) And (Val(astrOrder(4)) <> 0) And (lFillQty <> 0) And (StatusIsFill(astrOrder(19)) = True) Then
                        bReceivedFill = True
                        bFillExisted = False
                        bFillChanged = True
                        
                        ' 01/16/2014 DAJ: Normal fills come in as GMT, but the negative quantity fills seem to
                        ' be coming through as central time.  Convert the central time fills to GMT so that
                        ' they get put in the right spot...
                        dFillTime = Val(astrOrder(6))
                        If lFillQty < 0 Then
                            dFillTime = ConvertTimeZone(dFillTime, "CHI", "GMT")
                        End If
                        
                        dFillPrice = GenesisPrice(astrOrder(2), astrOrder(4))
                        
                        For lIndex = 1 To .Fills.Count
                            If .Fills(lIndex).BrokerID = Str(lFillID) Then
                                Set Fill = .Fills(lIndex)
                                bFillExisted = True
                            
                                bFillChanged = (Fill.FillDate <> dFillTime) Or (Fill.Price <> dFillPrice) Or (Fill.Quantity <> lFillQty)
                            
                                Exit For
                            End If
                        Next lIndex
                        
                        With Fill
                            .AccountID = Order.AccountID
                            .AutoTradingItemID = Order.AutoTradeItemID
                            .BrokerID = Str(lFillID)
                            .BrokerOrderID = Order.BrokerID
                            .Buy = Order.Buy
                            .FillDate = dFillTime
                            .IsManual = False
                            .IsSnapshot = True
                            .OrderID = Order.OrderID
                            .Price = dFillPrice
                            .Quantity = Abs(lFillQty)
                            .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.FillDate, m.nBroker, .Symbol, False))
                            .Symbol = Order.Symbol
                            .SymbolID = Order.SymbolID
                            
                            ' 03/16/2010 DAJ: If we get a fill with a negative quantity (which is a special
                            ' case that the TransAct desk will do in rare cases), we need to change the quantity
                            ' to a positive quantity and reverse the Buy/Sell code...
                            If .Quantity < 0 Then
                                .Quantity = .Quantity * -1&
                                .Buy = Not .Buy
                            End If
                            
                            ' If we don't have a Fill ID yet (this is a new fill), we need to save
                            ' the fill to get a Fill ID for the key into the Fills collection of
                            ' the order...
                            If .FillID = 0& Then .Save
                        End With
                        
                        Set .Fills(Str(Fill.FillID)) = Fill
                        DumpDebug mTradeTracker.FillDisplay(Fill)
                    Else
                        bReceivedFill = False
                    End If
        
                    .Status = GenesisOrderStatus(astrOrder(19))
                    If astrOrder(6) = "0" Then
                        .StatusDate = Val(astrOrder(21))
                    Else
                        .StatusDate = Val(astrOrder(6))
                    End If
                    
                    If (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_ParkPending) Then
                        .Status = eTT_OrderStatus_Parked
                    ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_Parked) Then
                        .Status = eTT_OrderStatus_Parked
                    ElseIf (Order.FillQuantity > Order.Quantity) Then
                        .Status = eTT_OrderStatus_OverFilled
                    ElseIf (.Status = eTT_OrderStatus_Filled) And (Order.FillQuantity < Order.Quantity) Then
                        .Status = eTT_OrderStatus_Partial
                    ElseIf (.Status = eTT_OrderStatus_Cancelled) And (Order.FillQuantity = Order.Quantity) Then
                        .Status = eTT_OrderStatus_Filled
                    ElseIf (.Status = eTT_OrderStatus_Cancelled) And (Order.FillQuantity > 0&) Then
                        .Status = eTT_OrderStatus_BalCancelled
                    ElseIf (.Status = eTT_OrderStatus_BalCancelled) And (Order.FillQuantity = Order.Quantity) Then
                        .Status = eTT_OrderStatus_Filled
                    End If
                    
                    .Save
                    DumpDebug Order.OrderText(True, True, True) & ": " & OrderStatus(Order.Status)
                    
                    If Order.Status = eTT_OrderStatus_OverFilled Then
                        Set frm = New frmAlertPopup
                        frm.ShowMessageBox "Trade Navigator has received duplicate fill|information for order " & Order.BrokerID & ".||This will cause this order to appear over-filled and will likely result in a position mismatch.||PLEASE CALL YOUR BROKER AND VERIFY YOUR POSITIONS.", "Over Filled Order on " & m.strBrokerName, vbCenter
                    End If
        
                    ' If this is a successful order status change of an order that is in the
                    ' modified array, update the order in the modified array...
                    If m.ModifiedOrders.Exists(astrOrder(10)) Then
                        Set m.ModifiedOrders(astrOrder(10)) = Order
                    End If
                    
                    If bReceivedFill = True Then
                        m.BrokerInfo.AddOrder Order, bRefresh
                        m.BrokerInfo.AddFill Fill, bRefresh
                    
                        FillCallback Order, Fill, bFillExisted, True, bFillChanged
                        If (nType = eGDTransActMessageType_FilledOrder) And (bFillExisted = False) Then
                            g.OrderStrategies.FillCallback Fill, Order
                        End If
                        If astrOrder(25) = "0" Then
                            g.ActivityLogs.AddFillToActivityLog m.nBroker, Fill, False
                            
                            lPos = -1&
                            For lIndex = 0 To m.astrLastFillTime.Size - 1
                                If Parse(m.astrLastFillTime(lIndex), vbTab, 1) = astrOrder(2) Then
                                    lPos = lIndex
                                    Exit For
                                End If
                            Next lIndex
                            
                            If lPos = -1& Then
                                m.astrLastFillTime.Add astrOrder(2) & vbTab & Str(gdTickCount)
                            Else
                                m.astrLastFillTime(lPos) = astrOrder(2) & vbTab & Str(gdTickCount)
                            End If
                            
                            m.dLastFillTime = gdTickCount
                        End If
                    Else
                        m.BrokerInfo.AddOrder Order, bRefresh
                        
                        OrderCallback Order
                        If astrOrder(25) = "0" Then
                            g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                        End If
                    End If
                End With
            Else
                ' If the "old" order ID exists in the modified orders collection, we need to
                ' restore the original order to "Working" since the modification likely failed...
                If m.ModifiedOrders.Exists(astrOrder(13)) Then
                    Set Order = m.ModifiedOrders(astrOrder(13))
                    m.ModifiedOrders.Remove astrOrder(13)
                    Order.Save
                    
                    m.BrokerInfo.AddOrder Order, bRefresh
                    OrderCallback Order
                    If astrOrder(25) = "0" Then
                        g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                    End If
                    
                    DumpDebug "Restoring Order: '" & Order.OrderText(True, True, True) & "'"
                
                ' If the "new" order ID exists in the modified orders collection, we need to
                ' restore the original order to "Working" since the modification likely failed...
                ElseIf m.ModifiedOrders.Exists(astrOrder(10)) Then
                    Set Order = m.ModifiedOrders(astrOrder(10))
                    m.ModifiedOrders.Remove astrOrder(10)
                    Order.Save
                    
                    m.BrokerInfo.AddOrder Order, bRefresh
                    OrderCallback Order
                    If astrOrder(25) = "0" Then
                        g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                    End If
                    
                    DumpDebug "Restoring Order: '" & Order.OrderText(True, True, True) & "'"
                    
                ' Otherwise mark the order as Error (or Balance Cancelled if partially filled) and
                ' update it everywhere...
                Else
                    With Order
                        If astrOrder(10) <> "0" Then .BrokerID = astrOrder(10)
                        .ExchangeID = astrOrder(11)
                        If Order.FillQuantity = 0 Then
                            .Status = eTT_OrderStatus_Error
                        Else
                            .Status = eTT_OrderStatus_BalCancelled
                        End If
                        If astrOrder(6) = "0" Then
                            .StatusDate = Val(astrOrder(21))
                        Else
                            .StatusDate = Val(astrOrder(6))
                        End If
                        .Message = astrOrder(24)
            
                        .Save
                        
                        m.BrokerInfo.AddOrder Order, bRefresh
                        OrderCallback Order
                        If astrOrder(25) = "0" Then
                            g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                        End If
                        
                        DumpDebug "Order Error: '" & Order.OrderText(True, True, True) & "': " & Order.Message
                    End With
                End If
                    
                If (m.strUserName <> kSimUserID) And ((nPrevStatus <> eTT_OrderStatus_Error) And (nPrevStatus <> eTT_OrderStatus_BalCancelled)) Then
                    Set frm = New frmAlertPopup
                    frm.ShowMessageBox m.strBrokerName & " order error " & astrOrder(0) & "||" & astrOrder(24) & "|", m.strBrokerName & " Order Error", vbCenter
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.HandleOrderCallback"
    
End Sub

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
    Dim Order As cPtOrder               ' Order to return
    Dim bFound As Boolean               ' Have we found the order?

    bFound = False
    lAccountID = g.Broker.AccountIDForNumber(strAccountNumber)
    
    Set Order = m.BrokerInfo.Order(strBrokerID, strGenesisOrderID)
    If Not Order Is Nothing Then
        bFound = True
    Else
        Set Order = m.BrokerInfo.Order(strOriginalID, strGenesisOrderID)
        If Not Order Is Nothing Then
            bFound = True
        End If
    End If
    
    If bFound = False Then
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
        If (bFound = False) And (Len(strBrokerID) > 0) And (strBrokerID <> "0") Then
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
        If (bFound = False) And (Len(strOriginalID) > 0) And (strOriginalID <> "0") Then
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
    End If
    
    If bFound Then
        Set LoadOrder = Order
    Else
        Set LoadOrder = Nothing
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.LoadOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrder
'' Description: Convert an array of TransAct order info to a Genesis order
'' Inputs:      Order Array
'' Returns:     Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrder(ByVal astrOrder As cGdArray) As cPtOrder
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order to return
    
    With Order
        .SymbolOrSymbolID = GenesisSymbol(astrOrder(2))
        .AccountID = g.Broker.AccountIDForNumber(astrOrder(1))
        If astrOrder(10) <> "0" Then .BrokerID = astrOrder(10)
        .Buy = (astrOrder(7) <> "0")
        .DateSent = Val(astrOrder(21))
        .ExchangeID = astrOrder(11)
        If CLng(Val(astrOrder(3))) = 0& Then
            .Expiration = -1&
        Else
            .Expiration = CLng(Val(astrOrder(3))) * -1&
        End If
        .LimitPrice = GenesisPrice(astrOrder(2), astrOrder(15))
        .OrderDate = Val(astrOrder(21))
        .OrderType = GenesisOrderType(astrOrder(12))
        .Quantity = Val(astrOrder(16))
        .Status = GenesisOrderStatus(astrOrder(19))
        If astrOrder(6) = "0" Then
            .StatusDate = Val(astrOrder(21))
        Else
            .StatusDate = Val(astrOrder(6))
        End If
        .StopPrice = GenesisPrice(astrOrder(2), astrOrder(20))
    End With
    
    If Len(Order.Symbol) > 0 Then
        Set GenesisOrder = Order
    Else
        Set GenesisOrder = Nothing
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.GenesisOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactOrder
'' Description: Convert a Genesis Order to a delimited string to send to TransAct
'' Inputs:      Genesis Order
'' Returns:     TransAct Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TransactOrder(ByVal Order As cPtOrder) As String
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information
    Dim strTransActSymbol As String     ' TransAct symbol for the Genesis symbol
    
    strTransActSymbol = TransactSymbol(Order.Symbol)
    If Len(strTransActSymbol) > 0 Then
        astrOrder(0) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(1) = strTransActSymbol
        If Order.Buy Then astrOrder(2) = "1" Else astrOrder(2) = "0"
        astrOrder(3) = TransactOrderType(Order.OrderType)
        astrOrder(4) = TransactPrice(strTransActSymbol, Order.StopPrice)
        astrOrder(5) = TransactPrice(strTransActSymbol, Order.LimitPrice)
        astrOrder(6) = Str(Order.Quantity)
        astrOrder(7) = m.strUserName
        astrOrder(8) = Str(CLng(Val(Order.GenesisOrderID)))
        astrOrder(9) = Order.BrokerID
        If Order.Expiration = 0& Then
            astrOrder(10) = "1"
        Else
            astrOrder(10) = "0"
        End If
        If Order.IsAutomated Then
            astrOrder(11) = "1"
        Else
            astrOrder(11) = "0"
        End If
        
        TransactOrder = astrOrder.JoinFields(vbTab)
    Else
        TransactOrder = ""
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.TransactOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisPrice
'' Description: Convert a TransAct price to a Genesis Price
'' Inputs:      TransAct Symbol, TransAct Price
'' Returns:     Genesis Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisPrice(ByVal strTransActSymbol As String, ByVal strTransactPrice As String) As Double
On Error GoTo ErrSection:

    Dim strBase As String               ' TransAct base symbol
    Dim dMult As Double                 ' Multiplier
    Dim lPos As Long                    ' Position in array
    Dim dReturn As Double               ' Return value from the function
    
    strBase = Parse(strTransActSymbol, "|", 1)
    
    dReturn = Val(strTransactPrice)
    If m.astrBaseSyms.BinarySearch(strBase & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        dMult = Val(Parse(m.astrBaseSyms(lPos), vbTab, 4))
        If dMult <> 0 Then
            Select Case dMult
                Case 8
                    dReturn = (Val(Left(strTransactPrice, Len(strTransactPrice) - 1)) * 100) + (Val(Right(strTransactPrice, 1)) / 8)
            
                Case 32
                    dReturn = Val(Left(strTransactPrice, Len(strTransactPrice) - 2)) + (Val(Right(strTransactPrice, 2)) / 32)
                
                Case 32.25
                    If Right(strTransactPrice, 1) = "2" Or Right(strTransactPrice, 1) = "7" Then
                        strTransactPrice = strTransactPrice & "5"
                    Else
                        strTransactPrice = strTransactPrice & "0"
                    End If
                    
                    dReturn = Val(Left(strTransactPrice, Len(strTransactPrice) - 4)) + ((Val(Right(strTransactPrice, 4)) / 100) / 32)
                
                Case 32.5
                    dReturn = Val(Left(strTransactPrice, Len(strTransactPrice) - 3)) + ((Val(Right(strTransactPrice, 3)) / 10) / 32)
                
                Case Else
                    dReturn = Val(strTransactPrice) * dMult
                    
            End Select
        End If
    End If
    
    GenesisPrice = RoundToSigDigits(dReturn)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.GenesisPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactPrice
'' Description: Convert a Genesis price to a TransAct Price
'' Inputs:      TransAct Symbol, Genesis Price
'' Returns:     TransAct Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TransactPrice(ByVal strTransActSymbol As String, ByVal dGenesisPrice As Double) As String
On Error GoTo ErrSection:

    Dim strBase As String               ' TransAct base symbol
    Dim dMult As Double                 ' Multiplier
    Dim lPos As Long                    ' Position in array
    Dim strReturn As String             ' Return value from the function
    
    strBase = Parse(strTransActSymbol, "|", 1)
    
    strReturn = Str(dGenesisPrice)
    If m.astrBaseSyms.BinarySearch(strBase & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        dMult = Val(Parse(m.astrBaseSyms(lPos), vbTab, 4))
        If dMult <> 0 Then
            Select Case dMult
                Case 8
                    strReturn = Replace(Format(Int(dGenesisPrice) / 100#, "0.00"), ",", ".") & Format((dGenesisPrice - Int(dGenesisPrice)) * 8, "0")
            
                Case 32
                    strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 32, "00")
                
                Case 32.25
                    strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 3200, "0000")
                    strReturn = Left(strReturn, Len(strReturn) - 1)
                    
                Case 32.5
                    strReturn = Str(Int(dGenesisPrice)) & "." & Format((dGenesisPrice - Int(dGenesisPrice)) * 320, "000")
                    
                Case Else
                    strReturn = Str(dGenesisPrice / dMult)

            End Select
        End If
    End If
    
    TransactPrice = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.GenesisPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderType
'' Description: Convert a TransAct order type to a Genesis order type
'' Inputs:      TransAct order type
'' Returns:     Genesis order type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderType(ByVal strTransactOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Select Case strTransactOrderType
        Case "1"
            GenesisOrderType = eTT_OrderType_Limit
        Case "2"
            GenesisOrderType = eTT_OrderType_Market
        Case "3"
            GenesisOrderType = eTT_OrderType_StopWithLimit
        Case "4"
            GenesisOrderType = eTT_OrderType_Stop
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.GenesisOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactOrderType
'' Description: Convert a Genesis order type to a TransAct order type
'' Inputs:      TransAct order type
'' Returns:     Genesis order type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TransactOrderType(ByVal nGenesisOrderType As eTT_OrderType) As String
On Error GoTo ErrSection:

    Select Case nGenesisOrderType
        Case eTT_OrderType_Limit, eTT_OrderType_LimitCloseOnly
            TransactOrderType = "1"
        
        Case eTT_OrderType_Market, eTT_OrderType_MarketOnClose
            TransactOrderType = "2"
        
        Case eTT_OrderType_StopWithLimit, eTT_OrderType_StopWithLimitCloseOnly
            TransactOrderType = "3"
            
        Case eTT_OrderType_Stop, eTT_OrderType_StopCloseOnly
            TransactOrderType = "4"
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.TransactOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderStatus
'' Description: Convert a TransAct order status to a Genesis order status
'' Inputs:      TransAct order status
'' Returns:     Genesis order status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderStatus(ByVal strTransactOrderStatus As String) As eTT_OrderStatus
On Error GoTo ErrSection:

    Select Case strTransactOrderStatus
        Case "1"    'Unsent
            GenesisOrderStatus = eTT_OrderStatus_Open
        Case "2"    'Open
            GenesisOrderStatus = eTT_OrderStatus_Working
        Case "3"    'Filled
            GenesisOrderStatus = eTT_OrderStatus_Filled
        Case "4"    'Modified
            GenesisOrderStatus = eTT_OrderStatus_Amended
        Case "5"    'Cancelled
            GenesisOrderStatus = eTT_OrderStatus_Cancelled
        Case "6"    'Unfilled
            GenesisOrderStatus = eTT_OrderStatus_Working
        Case "7"    'OrderInProgress
            GenesisOrderStatus = eTT_OrderStatus_Working
        Case "8"    'CancelInProgress
            GenesisOrderStatus = eTT_OrderStatus_CancelPending
        Case "9"    'ManualCancel
            GenesisOrderStatus = eTT_OrderStatus_Cancelled
        Case "10"   'ExchangeCancel
            GenesisOrderStatus = eTT_OrderStatus_Cancelled
        
        Case "1000"
            GenesisOrderStatus = eTT_OrderStatus_BalCancelled
        Case "1001"
            GenesisOrderStatus = eTT_OrderStatus_Partial
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.GenesisOrderStatus"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdatePrice
'' Description: Update a price from a price callback
'' Inputs:      Price
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdatePrice(ByVal strPrice As String)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' TransAct symbol
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim bJustAdded As Boolean           ' Did we add this price for the first time?
    Dim lPos As Long                    ' Position of the symbol in the array
    Dim astrPrice As New cGdArray       ' Current price
    Dim astrPrevPrice As New cGdArray   ' Previous price
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrDepth As New cGdArray       ' Current depth
    Dim astrPrevDepth As New cGdArray   ' Previous depth
    Dim astrMessage As New cGdArray     ' Message to send to real-time
    Dim bBidChanged As Boolean          ' Did the bid change?
    Dim bAskChanged As Boolean          ' Did the ask change?
    Dim dExchTime As Double             ' Time of the tick in exchange time
    Dim strExchTimeZone As String       ' Exchange time zone information for the symbol
    Dim Bars As cGdBars                 ' Bars object for getting exchange time zone information
    Dim dTickTime As Double             ' Tick time from the TransAct feed
    
    If Not g.RealTime.Active Then
        m.astrPrices.Clear
        Exit Sub
    End If
    
    strSymbol = Parse(strPrice, vbTab, 1)
    strGenesisSymbol = GenesisSymbol(strSymbol)
    
    If Len(strGenesisSymbol) > 0 Then
        bJustAdded = False
        If m.astrPrices.BinarySearch(strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = False Then
            bJustAdded = True
            m.astrPrices(lPos) = strPrice
        End If
        
        astrPrice.SplitFields strPrice, vbTab
        astrPrevPrice.SplitFields m.astrPrices(lPos), vbTab
        
        ' convert current time to exchange time (NY for all US exchanges)
        dExchTime = g.RealTime.FeedTime
        Set Bars = m.SymbolBars(strGenesisSymbol)
        If Bars Is Nothing Then
            Set Bars = New cGdBars
            SetBarProperties Bars, strGenesisSymbol
            m.SymbolBars.Add Bars, strGenesisSymbol
        End If
        strExchTimeZone = Bars.Prop(eBARS_ExchangeTimeZoneInf)
        If Len(strExchTimeZone) > 0 And UCase(strExchTimeZone) <> "NY" And dExchTime > 0 Then
            dExchTime = ConvertTimeZone(dExchTime, "NY", strExchTimeZone)
        End If
        Set Bars = Nothing
        
        'If (astrPrice(5) <> astrPrevPrice(5)) Then
        If Val(astrPrice(16)) <> 0 And dExchTime > 0 Then
            dTickTime = Val(astrPrice(17))
            If (dTickTime <> 0) And (Len(strExchTimeZone) > 0) And (UCase(strExchTimeZone) <> "GMT") Then
                dTickTime = ConvertTimeZone(dTickTime, "GMT", strExchTimeZone)
            End If
            
            astrMessage.Clear
            astrMessage(0) = strGenesisSymbol
            astrMessage(1) = Str(dExchTime)
            astrMessage(2) = Str(GenesisPrice(strSymbol, astrPrice(3)))
            astrMessage(3) = astrPrice(4)
            g.RealTime.StreamedMessage 20, astrMessage.JoinFields(vbTab), m.nBroker
        End If
        
        For lIndex = 6 To 15
            astrMessage.Clear
            astrMessage(0) = strGenesisSymbol
            astrMessage(1) = Str(lIndex - 6)
            astrMessage(2) = Str(dExchTime)       ' Date/time
            astrMessage(7) = ""         ' Bid tick
            astrMessage(8) = "Y"        ' Quote condition
            
            astrDepth.SplitFields astrPrice(lIndex), "|"
            astrPrevDepth.SplitFields astrPrevPrice(lIndex), "|"
            
            bBidChanged = False
            bAskChanged = False
            
            If (astrDepth(3) <> astrPrevDepth(3)) Or (astrDepth(4) <> astrPrevDepth(4)) Or bJustAdded Then
                bBidChanged = True
                astrMessage(3) = Str(GenesisPrice(strSymbol, astrDepth(3)))
                astrMessage(4) = astrDepth(4)
                If lIndex = 6 Then
                    g.RealTime.StreamedMessage 25, astrMessage(0) & vbTab & astrMessage(2) & vbTab & astrMessage(3) & vbTab & astrMessage(4), m.nBroker
                End If
            End If
        
            If (astrDepth(1) <> astrPrevDepth(1)) Or (astrDepth(2) <> astrPrevDepth(2)) Or bJustAdded Then
                bAskChanged = True
                astrMessage(5) = Str(GenesisPrice(strSymbol, astrDepth(1)))
                astrMessage(6) = astrDepth(2)
                If lIndex = 6 Then
                    g.RealTime.StreamedMessage 26, astrMessage(0) & vbTab & astrMessage(2) & vbTab & astrMessage(5) & vbTab & astrMessage(6), m.nBroker
                End If
            End If
            
            If bBidChanged Or bAskChanged Then
                g.RealTime.StreamedMessage 52, astrMessage.JoinFields(vbTab), m.nBroker
            End If
        Next lIndex
        
        m.astrPrices(lPos) = strPrice
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.UpdatePrice"
    
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

#If 0 Then

    Dim fh As Integer                   ' File handle to open file with

    fh = FreeFile
    Open AddSlash(m.strLogPath) & "TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append Shared As #fh
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
'' Function:    ClearPrevPrices
'' Description: Clear the prices array
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearPrevPrices()
On Error GoTo ErrSection:

    m.astrPrices.Clear
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.ClearPrevPrices"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SynchronizePositions
'' Description: Synchronize the positions between Genesis and TransAct
'' Inputs:      Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SynchronizePositions(Optional ByVal strSymbol As String = "")
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
    
    frmPositionConfirm.ShowMe m.nBroker, "", strSymbol, False, False
    
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
                    If lPosition <> lBrokerPos Then
                        If InStr(strFirstTime, "," & strAccount & ";" & .Symbol & ",") <> 0 Then
                            DumpDebug "Position Mismatch (" & strAccount & ", " & .Symbol & "): Second Time"
                            strFirstTime = Replace(strFirstTime, "," & strAccount & ";" & .Symbol & ",", "")
                            
                            m.BrokerInfo.ConsistentBroker(strAccount, .Symbol) = False
                        
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
                        End If
                    End If
                End If
            End With
        Next lIndex
    End If
    
    If Len(strFirstTime) > 0 Then
        ' Set the last fill time so that 5 seconds from now we will ask for positions again...
        m.dLastFillTime = gdTickCount
    Else
        m.dLastFillTime = 0#
    End If

    ' Set the variable to notify that we are no longer synchronizing positions...
    m.bSyncInProgress = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.SynchronizePositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateConnectionInfo
'' Description: Update the connection information
'' Inputs:      Connection Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateConnectionInfo(ByVal strConnectionInfo As String)
On Error GoTo ErrSection:

    Dim dTickCount As Double            ' Current tick count from message
    Dim nStatus As eGDConnectionStatus  ' Connection status
    Dim strAccount As String            ' Account number
    Static strLastUser As String        ' Last user name
    
    dTickCount = Val(Parse(strConnectionInfo, vbTab, 1))
    nStatus = CLng(Val(Parse(strConnectionInfo, vbTab, 2)))
    strAccount = Parse(strConnectionInfo, vbTab, 3)
    If Parse(strConnectionInfo, vbTab, 4) = "1" Then
        m.bAllowGTC = True
    Else
        m.bAllowGTC = False
    End If
    
    Select Case nStatus
        Case eGDConnectionStatus_Disconnected
            m.BrokerInfo.SetConnectionStatus nStatus
            
            If ConnectionStatus <> eGDConnectionStatus_Disconnected Then
                If m.bManualDisconnect Then
                    m.dLastStatusMsg = 0#
                    m.dHeartBeat = 0#
                    m.bVerifyPositions = True
                    m.bSuccessfullyConnected = False
                    m.lReconnectAttempts = -1&
                Else
                    m.dLastStatusMsg = dTickCount
                    m.dHeartBeat = dTickCount - 110000
                End If
                
                If m.bWasConnected = True Then
                    m.bWasConnected = False
                    g.Alerts.CheckBrokerStatusAlert m.nBroker
                End If
                
                m.astrSubContract.Size = 0
                g.RealTime.UseBrokerFeed = 0
                
                DumpDebug "Disconnected from " & m.strBrokerName & " -- Unloading the application"
                SendTransactMessage eGDTransactMessageType_UnloadApp, "", True
            End If
        
        Case eGDConnectionStatus_Disconnecting
            m.BrokerInfo.SetConnectionStatusForUser nStatus, strLastUser
            m.dLastStatusMsg = dTickCount
        
        Case eGDConnectionStatus_Connecting
            m.BrokerInfo.SetConnectionStatusForUser nStatus, m.strUserName
            m.dLastStatusMsg = dTickCount
        
        Case eGDConnectionStatus_Connected
            m.dLastStatusMsg = dTickCount
            m.bSuccessfullyConnected = True
            m.bWasConnected = True
            If ConnectionStatus <> eGDConnectionStatus_Connected Then
                DumpDebug "Connected to " & m.strBrokerName
                
                LastDateConnected = Date
                GetAccounts
                
                ' 02/12/2013 DAJ: See comment below
'                If m.nLoginMode = eGDTransActLoginMode_Live Then
'                    If ((m.lReconnectAttempts > 1&) Or (m.lReconnectAttempts = -1&)) And (m.strUserName <> SimUserUserName) Then
'                        InfBox "Synchronizing Orders, Fills, and Positions with " & m.strBrokerName & " servers.  Please wait...", , , m.strBrokerName & " Refresh", True
'                    End If
'                End If
                
                m.strAccountNumber = strAccount
                SetIniFileProperty "LastAccount", m.strAccountNumber, "User", m.strIniFile
            End If
            m.lNumSetAccountFailures = 0
                
    End Select
    
    If m.strUserName <> strLastUser Then
        strLastUser = m.strUserName
    End If
    
    If ConnectionStatus <> nStatus Then
        DumpDebug vbTab & m.strBrokerName & " Message Received (" & MessageTypeToString(eGDTransActMessageType_ConnectionInfo) & "): " & strConnectionInfo
        ConnectionStatus = nStatus
        
        g.Broker.ConnectionStatusChanged m.nBroker, nStatus
        
        If ConnectionStatus = eGDConnectionStatus_Connected Then ResyncSubscriptionList
        If g.bUnloading = False Then
            RefreshAccountCombos
        End If
        ClearPrevPrices
       
        ' 02/12/2013 DAJ: Due to the fact that a customer was getting events in a little bit of a wierd
        ' order, I have decided to move this block down here from up above.  I think that the DoEvents in
        ' the ShowMe of the InfBox is causing the issue...
        If ConnectionStatus = eGDConnectionStatus_Connected Then
            If m.nLoginMode = eGDTransActLoginMode_Live Then
                If ((m.lReconnectAttempts > 1&) Or (m.lReconnectAttempts = -1&)) And (m.strUserName <> SimUserUserName) Then
                    InfBox "Synchronizing Orders, Fills, and Positions with " & m.strBrokerName & " servers.  Please wait...", , , m.strBrokerName & " Refresh", True
                End If
            Else
                InfBox ""
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.UpdateConnectionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MessageTypeToString
'' Description: Convert a message type enumeration to a string
'' Inputs:      Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MessageTypeToString(ByVal nType As eGDTransActMessageTypes) As String
On Error GoTo ErrSection:

    Select Case nType
        Case eGDTransActMessageType_Connect
            MessageTypeToString = "eGDTransActMessageType_Connect"
        Case eGDTransActMessageType_Disconnect
            MessageTypeToString = "eGDTransActMessageType_Disconnect"
        Case eGDTransActMessageType_AddOrder
            MessageTypeToString = "eGDTransActMessageType_AddOrder"
        Case eGDTransActMessageType_AmendOrder
            MessageTypeToString = "eGDTransActMessageType_AmendOrder"
        Case eGDTransActMessageType_CancelOrder
            MessageTypeToString = "eGDTransActMessageType_CancelOrder"
        Case eGDTransactMessageType_UnloadApp
            MessageTypeToString = "eGDTransActMessageType_UnloadApp"
        Case eGDTransActMessageType_GetOrders
            MessageTypeToString = "eGDTransActMessageType_GetOrders"
        Case eGDTransActMessageType_GetTrades
            MessageTypeToString = "eGDTransActMessageType_GetTrades"
        Case eGDTransActMessageType_Subscribe
            MessageTypeToString = "eGDTransActMessageType_Subscribe"
        Case eGDTransActMessageType_Unsubscribe
            MessageTypeToString = "eGDTransActMessageType_Unsubscribe"
        Case eGDTransActMessageType_ChangeAccount
            MessageTypeToString = "eGDTransActMessageType_ChangeAccount"
        Case eGDTransActMessageType_GetPositions
            MessageTypeToString = "eGDTransActMessageType_GetPositions"
        Case eGDTransActMessageType_UpdatePositions
            MessageTypeToString = "eGDTransActMessageType_UpdatePositions"
        Case eGDTransActMessageType_GetAccounts
            MessageTypeToString = "eGDTransActMessageType_GetAccounts"
        Case eGDTransActMessageType_LogonForAccountList
            MessageTypeToString = "eGDTransActMessageType_LogonForAccountList"
        Case eGDTransActMessageType_Logon
            MessageTypeToString = "eGDTransActMessageType_Logon"
        
        Case eGDTransActMessageType_Connected
            MessageTypeToString = "eGDTransActMessageType_Connected"
        Case eGDTransActMessageType_Disconnected
            MessageTypeToString = "eGDTransActMessageType_Disconnected"
        Case eGDTransActMessageType_Subscribed
            MessageTypeToString = "eGDTransActMessageType_Subscribed"
        Case eGDTransActMessageType_Unsubscribed
            MessageTypeToString = "eGDTransActMessageType_Unsubscribed"
        Case eGDTransActMessageType_CancelledOrder
            MessageTypeToString = "eGDTransActMessageType_CancelledOrder"
        Case eGDTransActMessageType_ChangedOrder
            MessageTypeToString = "eGDTransActMessageType_ChangedOrder"
        Case eGDTransActMessageType_FilledOrder
            MessageTypeToString = "eGDTransActMessageType_FilledOrder"
        Case eGDTransActMessageType_ExpiredOrder
            MessageTypeToString = "eGDTransActMessageType_ExpiredOrder"
        Case eGDTransActMessageType_SentOrder
            MessageTypeToString = "eGDTransActMessageType_SentOrder"
        Case eGDTransActMessageType_GetOrder
            MessageTypeToString = "eGDTransActMessageType_GetOrder"
        Case eGDTransActMessageType_GetTrade
            MessageTypeToString = "eGDTransActMessageType_GetTrade"
        Case eGDTransActMessageType_AppLoaded
            MessageTypeToString = "eGDTransActMessageType_AppLoaded"
        Case eGDTransActMessageType_AppUnloaded
            MessageTypeToString = "eGDTransActMessageType_AppUnloaded"
        Case eGDTransActMessageType_Heartbeat
            MessageTypeToString = "eGDTransActMessageType_Heartbeat"
        Case eGDTransActMessageType_PriceUpdate
            MessageTypeToString = "eGDTransActMessageType_PriceUpdate"
        Case eGDTransActMessageType_Error
            MessageTypeToString = "eGDTransActMessageType_Error"
        Case eGDTransActMessageType_Account
            MessageTypeToString = "eGDTransActMessageType_Account"
        Case eGDTransActMessageType_Position
            MessageTypeToString = "eGDTransActMessageType_Position"
        Case eGDTransActMessageType_ConnectionInfo
            MessageTypeToString = "eGDTransActMessageType_ConnectionInfo"
        Case eGDTransActMessageType_GetPosition
            MessageTypeToString = "eGDTransActMessageType_GetPosition"
        Case eGDTransActMessageType_AccountList
            MessageTypeToString = "eGDTransActMessageType_AccountList"
            
        Case Else
            MessageTypeToString = Str(nType)
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.MessageTypeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StringToMessageType
'' Description: Convert a string to a message type enumeration
'' Inputs:      String
'' Returns:     Message Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StringToMessageType(ByVal strMessageType As String) As eGDTransActMessageTypes
On Error GoTo ErrSection:

    Dim nReturn As eGDTransActMessageTypes  ' Return value for the function

    Select Case strMessageType
        Case "eGDTransActMessageType_Connect"
            nReturn = eGDTransActMessageType_Connect
        Case "eGDTransActMessageType_Disconnect"
            nReturn = eGDTransActMessageType_Disconnect
        Case "eGDTransActMessageType_AddOrder"
            nReturn = eGDTransActMessageType_AddOrder
        Case "eGDTransActMessageType_AmendOrder"
            nReturn = eGDTransActMessageType_AmendOrder
        Case "eGDTransActMessageType_CancelOrder"
            nReturn = eGDTransActMessageType_CancelOrder
        Case "eGDTransactMessageType_UnloadApp"
            nReturn = eGDTransactMessageType_UnloadApp
        Case "eGDTransActMessageType_GetOrders"
            nReturn = eGDTransActMessageType_GetOrders
        Case "eGDTransActMessageType_GetTrades"
            nReturn = eGDTransActMessageType_GetTrades
        Case "eGDTransActMessageType_Subscribe"
            nReturn = eGDTransActMessageType_Subscribe
        Case "eGDTransActMessageType_Unsubscribe"
            nReturn = eGDTransActMessageType_Unsubscribe
        Case "eGDTransActMessageType_ChangeAccount"
            nReturn = eGDTransActMessageType_ChangeAccount
        Case "eGDTransActMessageType_GetPositions"
            nReturn = eGDTransActMessageType_GetPositions
        Case "eGDTransActMessageType_UpdatePositions"
            nReturn = eGDTransActMessageType_UpdatePositions
        Case "eGDTransActMessageType_GetAccounts"
            nReturn = eGDTransActMessageType_GetAccounts
        Case "eGDTransActMessageType_LogonForAccountList"
            nReturn = eGDTransActMessageType_LogonForAccountList
        Case "eGDTransActMessageType_Logon"
            nReturn = eGDTransActMessageType_Logon
        
        Case "eGDTransActMessageType_Connected"
            nReturn = eGDTransActMessageType_Connected
        Case "eGDTransActMessageType_Disconnected"
            nReturn = eGDTransActMessageType_Disconnected
        Case "eGDTransActMessageType_Subscribed"
            nReturn = eGDTransActMessageType_Subscribed
        Case "eGDTransActMessageType_Unsubscribed"
            nReturn = eGDTransActMessageType_Unsubscribed
        Case "eGDTransActMessageType_CancelledOrder"
            nReturn = eGDTransActMessageType_CancelledOrder
        Case "eGDTransActMessageType_ChangedOrder"
            nReturn = eGDTransActMessageType_ChangedOrder
        Case "eGDTransActMessageType_FilledOrder"
            nReturn = eGDTransActMessageType_FilledOrder
        Case "eGDTransActMessageType_ExpiredOrder"
            nReturn = eGDTransActMessageType_ExpiredOrder
        Case "eGDTransActMessageType_SentOrder"
            nReturn = eGDTransActMessageType_SentOrder
        Case "eGDTransActMessageType_GetOrder"
            nReturn = eGDTransActMessageType_GetOrder
        Case "eGDTransActMessageType_GetTrade"
            nReturn = eGDTransActMessageType_GetTrade
        Case "eGDTransActMessageType_AppLoaded"
            nReturn = eGDTransActMessageType_AppLoaded
        Case "eGDTransActMessageType_AppUnloaded"
            nReturn = eGDTransActMessageType_AppUnloaded
        Case "eGDTransActMessageType_Heartbeat"
            nReturn = eGDTransActMessageType_Heartbeat
        Case "eGDTransActMessageType_PriceUpdate"
            nReturn = eGDTransActMessageType_PriceUpdate
        Case "eGDTransActMessageType_Error"
            nReturn = eGDTransActMessageType_Error
        Case "eGDTransActMessageType_Account"
            nReturn = eGDTransActMessageType_Account
        Case "eGDTransActMessageType_Position"
            nReturn = eGDTransActMessageType_Position
        Case "eGDTransActMessageType_ConnectionInfo"
            nReturn = eGDTransActMessageType_ConnectionInfo
        Case "eGDTransActMessageType_GetPosition"
            nReturn = eGDTransActMessageType_GetPosition
        Case "eGDTransActMessageType_AccountList"
            nReturn = eGDTransActMessageType_AccountList
        
        Case Else
            nReturn = 0&
    End Select
    
    StringToMessageType = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.StringToMessageType"
    

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleOrderError
'' Description: Handle an order error callback from the TransAct app
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleOrderError(ByVal strOrderInfo As String)
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Order split out in fields
    Dim Order As cPtOrder               ' Order object
    Dim nPrevStatus As eTT_OrderStatus  ' Previous order status
    Dim frm As frmAlertPopup            ' Alert popup form
    
    astrOrder.SplitFields strOrderInfo, vbTab
    
    Set Order = LoadOrder(astrOrder(0), astrOrder(2), astrOrder(3), "")
    If Not Order Is Nothing Then
        nPrevStatus = Order.Status
        If Order.FillQuantity = 0 Then
            Order.Status = eTT_OrderStatus_Error
        Else
            Order.Status = eTT_OrderStatus_BalCancelled
        End If
        Order.StatusDate = ConvertTimeZone(CurrentTime, "", "GMT")
        Order.Message = astrOrder(4)
        Order.Save
        
        m.BrokerInfo.AddOrder Order, False
        OrderCallback Order
        g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
    
        If (m.strUserName <> kSimUserID) And ((nPrevStatus <> eTT_OrderStatus_Error) And (nPrevStatus <> eTT_OrderStatus_BalCancelled)) Then
            If astrOrder(1) = "1007" Then
                Set frm = New frmAlertPopup
                frm.ShowMessageBox "You cannot trade " & Order.Symbol & " with " & m.strBrokerName & " account " & astrOrder(0) & ".||Please contact your broker if you wish to trade this symbol.|", m.strBrokerName & " Order Error", vbLeftJustify
            Else
                Set frm = New frmAlertPopup
                frm.ShowMessageBox m.strBrokerName & " order error " & astrOrder(1) & "||" & astrOrder(4) & "|", m.strBrokerName & " Order Error", vbLeftJustify
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.HandleOrderError"
    
End Sub

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

    Dim astrPosition As cGdArray        ' Array of position information split out from string
    Dim astrToAdd As cGdArray           ' Array of information to pass to broker info class
    Dim lPos As Long                    ' Position in an array
    
    If Len(strPositionString) > 0 Then
        Set astrPosition = New cGdArray
        astrPosition.Create eGDARRAY_Strings
        Set astrToAdd = New cGdArray
        astrToAdd.Create eGDARRAY_Strings
        
        astrPosition.SplitFields strPositionString, vbTab
        
        ''If (astrPosition(1) <> kTransActSimUserAccount) And (astrPosition(1) <> kTransActOldSimUserAccount) Then
        If (Not TransActSimulatedAccount(astrPosition(1))) And (m.nLoginMode = eGDTransActLoginMode_Live) Then
            astrToAdd(0) = astrPosition(1)
            astrToAdd(1) = GenesisSymbol(astrPosition(2))
            astrToAdd(2) = astrPosition(7)
            astrToAdd(3) = ""
            astrToAdd(4) = astrPosition(4)
            
            m.BrokerInfo.AddPosition astrToAdd.JoinFields(vbTab), bRefresh
            
            ' Add the broker information to the local broker position array...
            If m.astrBrokerPos.BinarySearch(astrToAdd(0) & vbTab & astrToAdd(1) & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
                m.astrBrokerPos(lPos) = astrToAdd(0) & vbTab & astrToAdd(1) & vbTab & astrToAdd(2) & vbTab & astrToAdd(4)
            Else
                m.astrBrokerPos.Add astrToAdd(0) & vbTab & astrToAdd(1) & vbTab & astrToAdd(2) & vbTab & astrToAdd(4), lPos
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.AddPositionToBrokerInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadEnabledSymbols
'' Description: Load the enabled symbols array
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadEnabledSymbols()
On Error GoTo ErrSection:

    Dim strFileName As String           ' Filename for the LKSyms file
    Dim lIndex As Long                  ' Index into a for loop

    strFileName = LksFile
    
    If FileExist(AddSlash(App.Path) & "Provided\LKTran.SYM") Then
        m.astrEnabledSymbols.FromFile (AddSlash(App.Path) & "Provided\LKTran.SYM")
        
        For lIndex = 0 To m.astrEnabledSymbols.Size - 1
            m.astrEnabledSymbols(lIndex) = Replace(m.astrEnabledSymbols(lIndex), "-", "")
        Next lIndex
        
        SaveEnabledSymbols
        
        KillFile AddSlash(App.Path) & "Provided\LKTran.SYM"
    ElseIf FileExist(AddSlash(App.Path) & "Provided\LKSyms.TRN") Then
        m.astrEnabledSymbols.Serialize AddSlash(App.Path) & "Provided\LKSyms.TRN", False
        
        For lIndex = 0 To m.astrEnabledSymbols.Size - 1
            m.astrEnabledSymbols(lIndex) = Replace(m.astrEnabledSymbols(lIndex), "-", "")
        Next lIndex
    
        SaveEnabledSymbols
        
        KillFile AddSlash(App.Path) & "Provided\LKSyms.TRN"
    End If
    
    If CalcFileCrc(strFileName) <> LKSCRC Then
        KillFile strFileName
    End If
    
    If m.astrEnabledSymbols.Serialize(strFileName, False) = False Then
        ' DAJ 04/22/2014: Because of the new CME rules, TransAct is going to be required to
        ' get rid of the "simuser" demo account ( they will need to go to individual simulated
        ' accounts per user ).  Because of this, we are going to stop connecting to "simuser",
        ' but utilize the BRKRDEMO enablement to still give the user real-time data on the same
        ' handful of symbols for the first month ( or until they login with a real login )...
        'If (g.Broker.IsBrokerSimUser(eTT_AccountType_TransAct) = True) And (g.Transact.UserName = g.Transact.SimUserUserName) Then
        If g.Broker.IsBrokerSimUser(eTT_AccountType_TransAct) = True Then
            'strSymbols = "G6A-,G6B-,G6C-,G6E-,G6J-,G6S-,E7-,EMD-,ER2-,ES-,GE-,NQ-,QG-,QM-,XK-,XY-,YM-,ZB-,ZC-,ZF-,ZG-,ZI-,ZL-,ZN-,ZS-,ZT-,ZU-,ZW-"
            'strSymbols = "G6E-,ER2-,ES-,NQ-,YM-,ZB-,ZN-"
            m.astrEnabledSymbols.SplitFields kDefaultSymbols, ","
            
            SaveEnabledSymbols
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.LoadEnabledSymbols"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveEnabledSymbols
'' Description: Save the enabled symbols array
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveEnabledSymbols()
On Error GoTo ErrSection:

    Dim strFileName As String           ' Filename for the LKSyms file

    strFileName = LksFile
    
    m.astrEnabledSymbols.Serialize strFileName, True
    LKSCRC = CalcFileCrc(strFileName)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.SaveEnabledSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VendorID
'' Description: Vendor ID to pass to TransAct login
'' Inputs:      None
'' Returns:     Vendor ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VendorID() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    strReturn = kVendorID4
    If Len(m.strVersion) > 0 Then
        If CLng(Val(Parse(m.strVersion, ".", 1))) >= 5& Then
            strReturn = kVendorID5
        End If
    End If
    
    VendorID = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransact.VendorID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StatusIsFill
'' Description: Determine if the status is valid for a fill
'' Inputs:      Status
'' Returns:     True if for fill, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StatusIsFill(ByVal strStatus As String) As Boolean
On Error GoTo ErrSection:

    StatusIsFill = ((strStatus = "3") Or (strStatus = "1000") Or (strStatus = "1001"))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTransAct.StatusIsFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TaSetup
'' Description: Setup the TransAct stuff
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TaSetup()
On Error GoTo ErrSection:

    Dim dAtVersion As Double            ' Version of the AT DLLs
    Dim lReturn As Long                 ' Return code from the registration
    
    If ATVersion = kNullData Then
        ' Attempt to register the necessary files...
        If Is9598orMe = False Then
            ' Make sure to ignore the error code for the msxml4.dll file because it sometimes
            ' returns an error code even though it appears to register correctly...
            ChangePath WinSysPath
            lReturn = gdRegisterFiles(WinSysPath & "msxml4.dll", ByVal 0)
            
            ChangePath AddSlash(CommonFilesPath) & "TransActAPI"
            lReturn = gdRegisterFiles(AddSlash(CommonFilesPath) & "TransActAPI\spr32x30.ocx", ByVal 0)
            lReturn = gdRegisterFiles(AddSlash(CommonFilesPath) & "TransActAPI\YesPrice.DLL", ByVal 0)
            lReturn = gdRegisterFiles(AddSlash(CommonFilesPath) & "TransActAPI\YesTypes.DLL", ByVal 0)
            lReturn = gdRegisterFiles(AddSlash(CommonFilesPath) & "TransActAPI\YesTradeEngineServer.DLL", ByVal 0)
            lReturn = gdRegisterFiles(AddSlash(CommonFilesPath) & "TransActAPI\YesControls.DLL", ByVal 0)
            
            ChangePath App.Path
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransAct.TaSetup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the object is created
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Dim strIpsIni As String             ' Ini file for TransAct IP's

    m.bAppLoaded = False
    
    m.strIniFile = AddSlash(App.Path) & "TransAct.INI"
    m.nBroker = eTT_AccountType_TransAct
    m.strBrokerName = g.Broker.BrokerName(m.nBroker)
    m.strProcessName = "GenTransAct"
    m.strProcessPath = AddSlash(App.Path) & "..\Brokers\TransAct\GenTransAct.EXE"
    m.strControlID = "Transact"
    m.strLogPath = AddSlash(App.Path) & "TransAct"
    m.lNumSetAccountFailures = 0
    
    ' Load up the list of demo accounts from the ini file...
    strIpsIni = AddSlash(App.Path) & "Provided\TaIps.INI"
    m.strDemoAccounts = DecryptFromHex(GetIniFileProperty("Accounts", "", "Demo", strIpsIni))
    If Len(m.strDemoAccounts) = 0 Then
        m.strDemoAccounts = GetIniFileProperty("Accts", "", "Demo", strIpsIni)
        SetIniFileProperty "Accounts", EncryptToHex(m.strDemoAccounts), "Demo", strIpsIni
        SetIniFileProperty "Accts", "", "Demo", strIpsIni
    End If
    
    If Not DirExist(m.strLogPath) Then MakeDir m.strLogPath
    KillFile AddSlash(m.strLogPath) & "*.LOG /o=-30"
    
    Set m.BrokerInfo = New cBrokerInfo
    m.BrokerInfo.Init m.nBroker, m.strLogPath
    
    Set m.SymbolBars = New cGdTree
    
    TaSetup
    
    If UCase(m.strUserName) <> UCase(SimUserUserName) Then
        m.strAccountNumber = GetIniFileProperty("LastAccount", "", "User", m.strIniFile)
    End If
    m.bAskedPassword = False
    
    m.lReconnectAttempts = -1

    Set m.astrBaseSyms = New cGdArray
    m.astrBaseSyms.Create eGDARRAY_Strings
    m.astrBaseSyms.FromFile AddSlash(App.Path) & "Provided\TrnToGen.TXT"
    m.astrBaseSyms.Sort
    
    Set m.astrPrices = New cGdArray
    m.astrPrices.Create eGDARRAY_Strings
    
    Set m.astrEnabledSymbols = New cGdArray
    m.astrEnabledSymbols.Create eGDARRAY_Strings
    LoadEnabledSymbols
    
    Set m.astrSubContract = New cGdArray
    m.astrSubContract.Create eGDARRAY_Strings
    
    Set m.astrLastFillTime = New cGdArray
    m.astrLastFillTime.Create eGDARRAY_Strings
    
    Set m.astrBrokerPos = New cGdArray
    m.astrBrokerPos.Create eGDARRAY_Strings
    
    Set m.astrToFix = New cGdArray
    m.astrToFix.Create eGDARRAY_Strings
    
    Set m.ModifiedOrders = New cGdTree
    
    ' Set up the YesTrader registry information if not set already...
    If GetRegistryValue(rkLocalMachine, kTransActRegistryKey, "LogFileLevel", -99999) = -99999 Then
        SetRegistryValue rkLocalMachine, kTransActRegistryKey, "LogFileLevel", 3, True
    End If
    If Len(GetRegistryValue(rkLocalMachine, kTransActRegistryKey, "LogFilePath", "")) = 0 Then
        SetRegistryValue rkLocalMachine, kTransActRegistryKey, "LogFilePath", m.strLogPath, False
    End If
    If GetRegistryValue(rkLocalMachine, kTransActRegistryKey, "OrderRetry", -99999) = -99999 Then
        SetRegistryValue rkLocalMachine, kTransActRegistryKey, "OrderRetry", 2, True
    End If
    If GetRegistryValue(rkLocalMachine, kTransActRegistryKey, "OrderWaitMIL", -99999) = -99999 Then
        SetRegistryValue rkLocalMachine, kTransActRegistryKey, "OrderWaitMIL", 300, True
    End If
    If GetRegistryValue(rkLocalMachine, kTransActRegistryKey, "PriceWaitMIL", -99999) = -99999 Then
        SetRegistryValue rkLocalMachine, kTransActRegistryKey, "PriceWaitMIL", 300, True
    End If
    If GetRegistryValue(rkLocalMachine, kTransActRegistryKey, "PriceRetry", -99999) = -99999 Then
        SetRegistryValue rkLocalMachine, kTransActRegistryKey, "PriceRetry", 2, True
    End If
    
    m.bWasConnected = False
    m.bVerifyPositions = True
    m.bSyncInProgress = False
    m.bFixCalledOnce = False
    m.dStartedProcess = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.Class_Initialize"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the object is destroyed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:
    
    Set m.BrokerInfo = Nothing
    Set m.SymbolBars = Nothing
    Set m.astrBaseSyms = Nothing
    Set m.astrPrices = Nothing
    Set m.astrEnabledSymbols = Nothing
    Set m.astrSubContract = Nothing
    Set m.astrLastFillTime = Nothing
    Set m.astrBrokerPos = Nothing
    Set m.astrToFix = Nothing
    Set m.ModifiedOrders = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTransact.Class_Terminate"
    
End Sub
