VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cRan.cls
'' Description: Wrapper class for trading through Rolfe & Nolan servers
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
'' 03/07/2011   DAJ         Broker Disconnected Alerts
'' 06/21/2011   DAJ         Last connection status change
'' 01/18/2012   DAJ         Enhanced logging for automated trading
'' 01/31/2012   DAJ         Handle unsolicted cancels in TradeSense order groups
'' 03/21/2012   DAJ         Fix for amending a partially filled order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum eGDRanOrderMode
    eGDRanOrderMode_Park = 0
    eGDRanOrderMode_Submit
    eGDRanOrderMode_Validate
End Enum

Private Enum eGDTableCols
    eGDTblCol_RanSymbol = 0
    eGDTblCol_Currency
    eGDTblCol_Description
    eGDTblCol_Exchange
    eGDTblCol_GenesisSymbol
    eGDTblCol_Session
    eGDTblCol_Multiplier
    eGDTblCol_SymbolType
    eGDTblCol_OrderTypeMask
    eGDTblCol_TimeInForceMask
    eGDTblCol_AllowRT
End Enum

Private Type mPrivate
    strIniFile As String                ' Ini file for Rolfe & Nolan broker
    nBroker As eTT_AccountType          ' Broker type for the Rolfe & Nolan broker
    strBrokerName As String             ' Broker name for the TT broker
    strProcessName As String            ' Name of the process to check
    strProcessPath As String            ' Full path of the process to start
    strControlID As String              ' Control ID
    strLogPath As String                ' Path for the log files
    
    strUserName As String               ' User Name to log on with
    strPassword As String               ' Password to log on with
    strIP As String                     ' IP of server to log on to
    strPort As String                   ' Port for the server connection
    
    nStatus As eGDConnectionStatus      ' Connection status
    bWasConnected As Boolean            ' This broker was has been connected
    dHeartBeat As Double                ' Heartbeat
    dLastStatusMsg As Double            ' Tick count of the last status message
    dConnectTime As Double              ' Tick count of the time we got connected
    bReconnecting As Boolean            ' Are we currently reconnecting?
    lReconnectAttempts As Long          ' Reconnect attempts that have been made
    bAppLoaded As Boolean               ' Is the Rolfe & Nolan broker application loaded?
    bManualDisconnect As Boolean        ' Are we doing a manual disconnect?
    bAskedPassword As Boolean           ' Have we asked for the password yet?
    dLastFillTime As Double             ' Tick count of when we received the last fill
    bGetAccountInfo As Boolean          ' Do we want to get account info?
    bVerifyPositions As Boolean         ' Do we want to verify the positions with the user?
    bVerboseRefresh As Boolean          ' Do we want to tell the user about the refresh?
    bSyncInProgress As Boolean          ' Are we in the process of synchronizing positions?
    bFixCalledOnce As Boolean           ' Has the fix positions call been done at least once?
    
    tblSymbols As cGdTable              ' Table of symbol translations
    aRanIdx As cGdArray                 ' Index of symbol table sorted by Rolfe & Nolan symbol
    aGenesisIdx As cGdArray             ' Index of symbol table sorted by Genesis symbol
    astrAccounts As cGdArray            ' List of accounts currently connected to
    astrBrokerPos As cGdArray           ' List of current positions
    astrOrderMap As cGdArray            ' Order number/type map
    astrToFix As cGdArray               ' List of account/symbol pairs that need their position fixed

    dRefreshingStart As Double          ' Start of the refreshing time
    bRefreshingAccounts As Boolean      ' Are we refreshing accounts?
    bRefreshingOrders As Boolean        ' Are we refreshing orders?
    bRefreshingFills As Boolean         ' Are we refreshing fills?
    bRefreshingPositions As Boolean     ' Are we refreshing positions?
    bRefreshingStatement As Boolean     ' Are we refreshing the statement?
    bRefreshingAcctStatus As Boolean    ' Are we refreshing the account status?

    astrAmendPending As cGdArray        ' Pending order amend information array
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

Public Property Get IPAddress() As String
    IPAddress = m.strIP
End Property
Public Property Let IPAddress(ByVal strIPAddress As String)
    m.strIP = strIPAddress
End Property

Public Property Get Port() As String
    Port = m.strPort
End Property
Public Property Let Port(ByVal strPort As String)
    m.strPort = strPort
End Property

Public Property Get HeartBeat()
    HeartBeat = m.dHeartBeat
End Property
Public Property Get LastStatusMessage() As Double
    LastStatusMessage = m.dLastStatusMsg
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
        If nNewStatus = eGDConnectionStatus_Connected Then
            g.Broker.PositionsToVerify = True
        End If
    End If
End Property

Private Function TblCol(ByVal nCol As eGDTableCols) As Long
    TblCol = nCol
End Function

Private Function DebugMode() As Boolean
    DebugMode = FileExist(AddSlash(App.Path) & "RanDebug.FLG")
End Function

Public Function IsBroker() As Boolean
    IsBroker = (InStr(UCase(m.strUserName), "-ALR") <> 0)
End Function

Public Property Get BrokerInfo() As cBrokerInfo
    Set BrokerInfo = m.BrokerInfo
End Property

Public Property Get Refreshing() As Boolean
    Refreshing = m.bRefreshingAccounts Or m.bRefreshingPositions Or m.bRefreshingOrders Or m.bRefreshingFills Or m.bRefreshingAcctStatus Or m.bRefreshingStatement
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
    RaiseError "cRan.CurrentPosition.Get"
    
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
    RaiseError "cRan.CurrentPosition.Let"
    
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
    RaiseError "cRan.CarriedPosition.Get"
    
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
    RaiseError "cRan.CarriedPosition.Let"
    
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
    Dim astrSymbols As New cGdArray     ' Array of symbols from a file
    
    Select Case nBroker
        Case eTT_AccountType_Alaron
            m.strIniFile = AddSlash(App.Path) & "Alrn.INI"
            m.nBroker = eTT_AccountType_Alaron
            m.strBrokerName = g.Broker.BrokerName(nBroker)
            m.strProcessName = "GenAlaron"
            m.strProcessPath = AddSlash(App.Path) & "..\Brokers\Alaron\GenAlaron.EXE"
            m.strControlID = "Alaron"
            m.strLogPath = AddSlash(App.Path) & "Alaron"
            strSymbolFile = AddSlash(App.Path) & "Provided\AlrToGen.TXT"
            
        Case eTT_AccountType_Cadent
            m.strIniFile = AddSlash(App.Path) & "Cadent.INI"
            m.nBroker = eTT_AccountType_Cadent
            m.strBrokerName = g.Broker.BrokerName(nBroker)
            m.strProcessName = "GenCadent"
            m.strProcessPath = AddSlash(App.Path) & "..\Brokers\Cadent\GenCadent.EXE"
            m.strControlID = "Cadent"
            m.strLogPath = AddSlash(App.Path) & "Cadent"
            strSymbolFile = AddSlash(App.Path) & "Provided\CadToGen.TXT"
    
        Case eTT_AccountType_Lotus
            m.strIniFile = AddSlash(App.Path) & "Lotus.INI"
            m.nBroker = eTT_AccountType_Lotus
            m.strBrokerName = g.Broker.BrokerName(nBroker)
            m.strProcessName = "GenLotus"
            m.strProcessPath = AddSlash(App.Path) & "..\Brokers\Lotus\GenLotus.EXE"
            m.strControlID = "Lotus"
            m.strLogPath = AddSlash(App.Path) & "Lotus"
            strSymbolFile = AddSlash(App.Path) & "Provided\LbsToGen.TXT"
    
    End Select
        
    If Not DirExist(m.strLogPath) Then MkDir m.strLogPath
    KillFile AddSlash(m.strLogPath) & "*.LOG /o=-30"
    
    Set m.BrokerInfo = New cBrokerInfo
    m.BrokerInfo.Init m.nBroker, m.strLogPath
    
    Set m.tblSymbols = New cGdTable
    Set m.aRanIdx = New cGdArray
    Set m.aGenesisIdx = New cGdArray
    
    Set m.astrAccounts = New cGdArray
    m.astrAccounts.Create eGDARRAY_Strings
    Set m.astrBrokerPos = New cGdArray
    m.astrBrokerPos.Create eGDARRAY_Strings
    
    Set m.astrOrderMap = New cGdArray
    m.astrOrderMap.Create eGDARRAY_Strings
    m.astrOrderMap.Serialize AddSlash(m.strLogPath) & "Order.Map", False
    
    With m.tblSymbols
        .CreateField eGDARRAY_Strings, TblCol(eGDTblCol_RanSymbol), "RanSymbol"
        .CreateField eGDARRAY_Strings, TblCol(eGDTblCol_Currency), "Currency"
        .CreateField eGDARRAY_Strings, TblCol(eGDTblCol_Description), "Description"
        .CreateField eGDARRAY_Strings, TblCol(eGDTblCol_Exchange), "Exchange"
        .CreateField eGDARRAY_Strings, TblCol(eGDTblCol_GenesisSymbol), "GenesisSymbol"
        .CreateField eGDARRAY_Strings, TblCol(eGDTblCol_Session), "Session"
        .CreateField eGDARRAY_Doubles, TblCol(eGDTblCol_Multiplier), "Multiplier"
        .CreateField eGDARRAY_Strings, TblCol(eGDTblCol_SymbolType), "SymbolType"
        .CreateField eGDARRAY_Strings, TblCol(eGDTblCol_OrderTypeMask), "OrderTypeMask"
        .CreateField eGDARRAY_Strings, TblCol(eGDTblCol_TimeInForceMask), "TimeInForceMask"
        .CreateField eGDARRAY_Strings, TblCol(eGDTblCol_AllowRT), "AllowRT"
    End With
    astrSymbols.FromFile strSymbolFile, , , ";"
    m.tblSymbols.FromString astrSymbols.JoinFields(vbLf), vbLf, vbTab
    
    Set m.aRanIdx = m.tblSymbols.CreateSortedIndex(TblCol(eGDTblCol_RanSymbol))
    Set m.aGenesisIdx = m.tblSymbols.CreateSortedIndex(TblCol(eGDTblCol_GenesisSymbol))
    
    m.bGetAccountInfo = True
    m.bVerifyPositions = True
    
    Set m.astrAmendPending = New cGdArray
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.Init"
    
End Sub

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

    If (Len(m.strUserName) = 0) Or (Len(m.strIP) = 0) Or (Len(m.strPort) = 0) Then
        ShowConfig = True
    Else
        ShowConfig = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.ShowConfig"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Connect
'' Description: Attempt a connection to the Rolfe & Nolan servers if not
''              already connected
'' Inputs:      New Account, Are we switching?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Connect(Optional ByVal strNewAccount As String = "", Optional ByVal bSwitching As Boolean = False)
On Error GoTo ErrSection:

    Dim bConnect As Boolean             ' Should we still try to connect?
    Dim lTimeOut As Long                ' Timeout variable
    
    If g.bUnloading = False Then
        'If Not IsIDE Then
        '    InfBox "You cannot connect to the " & m.strBrokerName & " servers with this version of Trade Navigator.", "i", , m.strBrokerName &  " Connect Message"
        'Else
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
                        Disconnect False, "Changing accounts to " & m.strUserName
                        
                        lTimeOut = 0&
                        Do While (ConnectionStatus <> eGDConnectionStatus_Disconnected) And (lTimeOut < 30&)
                            Sleep 1
                            lTimeOut = lTimeOut + 1&
                        Loop
                        
                        bConnect = (ConnectionStatus = eGDConnectionStatus_Disconnected)
                        If bConnect And FormIsLoaded("frmBrokerScreen") Then
                            frmBrokerScreen.ClearForm
                        End If
                    End If
                End If
            End If
                
            If (bConnect = True) And (g.bUnloading = False) Then
                If LiveTradingAllowed(m.nBroker) Then
                    ConnectionStatus = eGDConnectionStatus_Connecting
                    
                    If ((frmOnlineBroker.gdBroker.FindControlID(m.strControlID) = 0) Or (KillProcess(m.strProcessName, True) = 0)) And (DebugMode = False) Then
                        If KillProcess(m.strProcessName) > 0 Then
                            Sleep 3
                        End If
                            
                        FileCopy AddSlash(App.Path) & "..\Brokers\GenRan.EXE", m.strProcessPath, True
                    
                        DumpDebug "Trade Navigator version: " & FormatVersion & "." & Str(App.Revision) & " " & DateFormat(FileDate(App.Path & "\" & App.EXEName & ".EXE"), MM_DD_YYYY, HH_MM, AMPM_UPPER)
                        DumpDebug "Loading application " & m.strProcessPath
                        RunProcess m.strProcessPath, , , vbHide ' , vbMinimizedNoFocus
                    Else
                        DumpDebug "Connecting to " & m.strBrokerName & " as " & m.strUserName & "(" & m.strIP & ":" & m.strPort & ")"
                        If IsBroker Then
                            SendRanMessage eGDRanMessageType_Connect, m.strUserName & vbTab & m.strPassword & vbTab & "B" & vbTab & m.strIP & vbTab & m.strPort
                        Else
                            SendRanMessage eGDRanMessageType_Connect, m.strUserName & vbTab & m.strPassword & vbTab & "U" & vbTab & m.strIP & vbTab & m.strPort
                        End If
                    End If
                End If
            End If
        'End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.Connect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disconnect
'' Description: Attempt to disconnect from the Rolfe & Nolan servers
'' Inputs:      Force Reconnect?, Reason
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Disconnect(Optional ByVal bForceReconnect As Boolean = False, Optional ByVal strReason As String = "") As Boolean
On Error GoTo ErrSection:
    
    If ConnectionStatus = eGDConnectionStatus_Connected Then
        ConnectionStatus = eGDConnectionStatus_Disconnecting
        
        m.bManualDisconnect = Not bForceReconnect
        If m.bManualDisconnect Then
            m.dHeartBeat = 0#
            g.OrderStrategies.DeactivateExitsForBroker m.nBroker, "Disconnecting from " & m.strBrokerName
        End If
        
        If Len(strReason) = 0 Then DumpDebug "Disconnecting from " & m.strBrokerName Else DumpDebug "Disconnecting from " & m.strBrokerName & ": " & strReason
        SendRanMessage eGDRanMessageType_Disconnect, ""
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.Disconnect"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reconnect
'' Description: Attempt to reconnect to the Rolfe & Nolan servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Reconnect()
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Time out variable

    If m.bReconnecting = False Then
        m.bReconnecting = True
        
        If (ConnectionStatus = eGDConnectionStatus_Connected) Then
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
    RaiseError "cRan.Reconnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Ask the server for the available accounts for this user
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAccounts(Optional ByVal bGetAccountInfo As Boolean = True)
On Error GoTo ErrSection:

    If Not Refreshing Then
        DumpDebug "Getting Accounts"
        m.bRefreshingAccounts = True
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for account information...", , , m.strBrokerName & " Refresh", True
        End If
        SendRanMessage eGDRanMessageType_GetAccounts, ""
        m.bGetAccountInfo = bGetAccountInfo
    Else
        DumpDebug "Refresh requested but already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.GetAccounts"
    
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
    SendRanMessage eGDRanMessageType_GetContracts, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.GetContracts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Ask the server for the available orders for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetOrders(ByVal strAccount As String)
On Error GoTo ErrSection:

    If Not Refreshing Then
        If Len(strAccount) = 0 Then DumpDebug "Getting Orders" Else DumpDebug "Getting orders for " & strAccount
        m.bRefreshingOrders = True
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for order information...", , , m.strBrokerName & " Refresh", True
        End If
        SendRanMessage eGDRanMessageType_GetOrders, strAccount
    Else
        DumpDebug "Refresh requested but already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.GetOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFills
'' Description: Ask the server for the available fills for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetFills(ByVal strAccount As String)
On Error GoTo ErrSection:

    If Not Refreshing Then
        If Len(strAccount) = 0 Then DumpDebug "Getting Fills" Else DumpDebug "Getting fills for " & strAccount
        m.bRefreshingFills = True
        If m.bVerboseRefresh Then
            InfBox "Asking " & m.strBrokerName & " servers for fill information...", , , m.strBrokerName & " Refresh", True
        End If
        SendRanMessage eGDRanMessageType_GetFills, strAccount
    Else
        DumpDebug "Refresh requested but already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.GetFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositions
'' Description: Ask the server for the available positions for the given account
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
                SendRanMessage eGDRanMessageType_GetPositions, m.astrAccounts(lIndex)
            Next lIndex
        Else
            DumpDebug "Getting Positions for " & strAccount
            SendRanMessage eGDRanMessageType_GetPositions, strAccount
        End If
    Else
        DumpDebug "Refresh requested but already in progress"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSecurityList
'' Description: Ask the server for the available security list for this user
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetSecurityList(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    If Len(strAccount) > 0 Then DumpDebug "Getting Security List for " & strAccount Else DumpDebug "Getting Security List"
    SendRanMessage eGDRanMessageType_GetSecurityList, strAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.GetSecurityList"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetStatement
'' Description: Ask the server for the statement for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetStatement(ByVal strAccount As String)
On Error GoTo ErrSection:

    DumpDebug "Getting Statement"
    SendRanMessage eGDRanMessageType_GetStatement, strAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.GetStatement"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSystemMessages
'' Description: Ask the server for the system messages
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetSystemMessages()
On Error GoTo ErrSection:

    DumpDebug "Getting System Messages"
    SendRanMessage eGDRanMessageType_GetSystemMessages, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.GetSystemMessages"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetUserMessages
'' Description: Ask the server for the user messages for this user
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetUserMessages()
On Error GoTo ErrSection:

    DumpDebug "Getting User Messages"
    SendRanMessage eGDRanMessageType_GetUserMessages, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.GetUserMessages"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccountStatus
'' Description: Ask the server for the status for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAccountStatus(ByVal strAccount As String)
On Error GoTo ErrSection:

    DumpDebug "Getting Account Status"
    SendRanMessage eGDRanMessageType_GetAccountStatus, strAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.GetAccountStatus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Refresh
'' Description: Allow the user to refresh their account
'' Inputs:      Verbose Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Refresh(Optional ByVal bVerbose As Boolean = False)
On Error GoTo ErrSection:

    If (ConnectionStatus = eGDConnectionStatus_Connected) Then
        DumpDebug "Refreshing"
        m.bVerboseRefresh = bVerbose
        GetAccounts
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.Refresh"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAccountInfo
'' Description: Allow the user to refresh their account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshAccountInfo(ByVal strAccount As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrAccounts As cGdArray        ' Array of accounts to refresh

    Set astrAccounts = New cGdArray
    astrAccounts.Create eGDARRAY_Strings

    If (ConnectionStatus = eGDConnectionStatus_Connected) And (Not Refreshing) Then
        m.bRefreshingOrders = True
        m.bRefreshingFills = True
        m.bRefreshingPositions = True
        m.bRefreshingAcctStatus = True
        m.bRefreshingStatement = True
        
        If Len(strAccount) > 0 Then
            astrAccounts.Add strAccount
        Else
            For lIndex = 0 To m.astrAccounts.Size - 1
                astrAccounts.Add m.astrAccounts(lIndex)
            Next lIndex
        End If
        
        For lIndex = 0 To astrAccounts.Size - 1
            DumpDebug "Refreshing account " & astrAccounts(lIndex)
            SendRanMessage eGDRanMessageType_GetOrders, astrAccounts(lIndex)
            ' DAJ 08/02/2007: Fill information comes with the orders, so we don't need
            ' to ask for fills specifically...
            ''SendRanMessage eGDRanMessageType_GetFills, astrAccounts(lIndex)
            SendRanMessage eGDRanMessageType_GetPositions, astrAccounts(lIndex)
            SendRanMessage eGDRanMessageType_GetStatement, astrAccounts(lIndex)
            SendRanMessage eGDRanMessageType_GetAccountStatus, astrAccounts(lIndex)
        Next lIndex
    Else
        DumpDebug "Account Info Refresh called for but refresh already in progress"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.RefreshAccountInfo"
    
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
    RaiseError "cRan.SendOrder"

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

    Dim strRanOrder As String           ' Rolfe & Nolan order
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    nPrevOrderType = Order.OrderType
    strRanOrder = RanOrder(Order)
    
    If Len(strRanOrder) = 0 Then
        DumpDebug "AddOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        DumpDebug "AddOrder: Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to " & m.strBrokerName & " account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "AddOrder: " & OrderType(Order.OrderType) & " orders not allowed on the " & m.strBrokerName & " servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed on the " & m.strBrokerName & " servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "AddOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    Else
        'GenesisOrderToBrokerInfo Order, eTT_OrderStatus_Sent
        DumpDebug "AddOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
        SendRanMessage eGDRanMessageType_AddOrder, "Submit" & vbTab & strRanOrder
        bReturn = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
    
    AddOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.AddOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddBrokerOrder
'' Description: Attempt to add the order
'' Inputs:      Order to add
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddBrokerOrder(ByVal strBrokerOrder As String, ByVal nOrderMode As eGDRanOrderMode)
On Error GoTo ErrSection:

    Select Case nOrderMode
        Case eGDRanOrderMode_Park
            SendRanMessage eGDRanMessageType_AddOrder, "Park" & vbTab & strBrokerOrder
        Case eGDRanOrderMode_Submit
            SendRanMessage eGDRanMessageType_AddOrder, "Submit" & vbTab & strBrokerOrder
        Case eGDRanOrderMode_Validate
            SendRanMessage eGDRanMessageType_AddOrder, "Validate" & vbTab & strBrokerOrder
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.AddBrokerOrder"
    
End Sub

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
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.AddAmendPendingInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AmendOrder
'' Description: Attempt to amend an order
'' Inputs:      Order to amend
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AmendOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim strRanOrder As String           ' Rolfe & Nolan order
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    nPrevOrderType = Order.OrderType
    strRanOrder = RanOrder(Order)
    Order.BrokerID = ""
    
    If Len(strRanOrder) = 0 Then
        DumpDebug "AmendOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to an " & m.strBrokerName & " symbol"
    ElseIf ConnectionStatusForAccount(g.Broker.AccountNumberForID(Order.AccountID)) <> eGDConnectionStatus_Connected Then
        DumpDebug "AmendOrder: Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNumberForID(Order.AccountID)
        InfBox "Not currently connected to " & m.strBrokerName & " account '" & g.Broker.AccountNameForID(Order.AccountID) & "'", "!", , m.strBrokerName & " Order Error"
        Order.Message = "Not currently connected to " & m.strBrokerName & " account " & g.Broker.AccountNameForID(Order.AccountID)
    ElseIf OrderTypeAllowed(Order.OrderType, Order.Symbol) = False Then
        DumpDebug "AmendOrder: " & OrderType(Order.OrderType) & " orders not allowed on the " & m.strBrokerName & " servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed on the " & m.strBrokerName & " servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "AmendOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    Else
        DumpDebug "AmendOrder: " & Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & ")"
        SendRanMessage eGDRanMessageType_AmendOrder, "Submit" & vbTab & strRanOrder
        bReturn = True
    End If
    
    If nPrevOrderType <> Order.OrderType Then
        Order.OrderType = nPrevOrderType
    End If
        
    AmendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.AmendOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AmendBrokerOrder
'' Description: Attempt to amend an order for the broker screen
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AmendBrokerOrder(ByVal strBrokerOrder As String, ByVal nOrderMode As eGDRanOrderMode)
On Error GoTo ErrSection:

    Select Case nOrderMode
        Case eGDRanOrderMode_Park
            SendRanMessage eGDRanMessageType_AmendOrder, "Park" & vbTab & strBrokerOrder
        Case eGDRanOrderMode_Submit
            SendRanMessage eGDRanMessageType_AmendOrder, "Submit" & vbTab & strBrokerOrder
        Case eGDRanOrderMode_Validate
            SendRanMessage eGDRanMessageType_AmendOrder, "Validate" & vbTab & strBrokerOrder
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.AmendBrokerOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Attempt to cancel an order
'' Inputs:      Order to cancel
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelOrder(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim strRanOrder As String           ' Rolfe & Nolan order
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    nPrevOrderType = Order.OrderType
    strRanOrder = RanOrder(Order)
    
    If Len(strRanOrder) = 0 Then
        DumpDebug "CancelOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
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
        DumpDebug "CancelOrder: " & OrderType(Order.OrderType) & " orders not allowed on the " & m.strBrokerName & " servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed on the " & m.strBrokerName & " servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "CancelOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    Else
        If Not g.TsoGroups Is Nothing Then
            g.TsoGroups.CancelRequested Order
        End If
        
        DumpDebug "CancelOrder: " & Order.OrderText & "(" & Order.GenesisOrderID & ", " & Order.BrokerID & ")"
        SendRanMessage eGDRanMessageType_CancelOrder, g.Broker.AccountNumberForID(Order.AccountID) & vbTab & Order.BrokerID
        
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
    RaiseError "cRan.CancelOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Attempt to Park an order
'' Inputs:      Order to Park
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ParkOrder(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim strRanOrder As String           ' Rolfe & Nolan order
    Dim bReturn As Boolean              ' Return value for the function
    Dim nPrevOrderType As eTT_OrderType ' Previous order type
    
    bReturn = False
    nPrevOrderType = Order.OrderType
    strRanOrder = RanOrder(Order)
    
    If Len(strRanOrder) = 0 Then
        DumpDebug "ParkOrder: " & Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
        InfBox Order.Symbol & " cannot be traded through the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = Order.Symbol & " cannot be converted to a " & m.strBrokerName & " symbol"
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
        DumpDebug "ParkOrder: " & OrderType(Order.OrderType) & " orders not allowed on the " & m.strBrokerName & " servers"
        InfBox OrderType(Order.OrderType) & " orders not allowed on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = OrderType(Order.OrderType) & " orders not allowed on the " & m.strBrokerName & " servers"
    ElseIf TimeInForceAllowed(Order.TimeInForce, Order.Symbol) = False Then
        DumpDebug "ParkOrder: " & TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
        InfBox TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers", "!", , m.strBrokerName & " Order Error"
        Order.Message = TimeInForce(Order.TimeInForce) & " orders not allowed for " & Order.Symbol & " on the " & m.strBrokerName & " servers"
    Else
        DumpDebug "ParkOrder: " & Order.OrderText & "(" & Order.GenesisOrderID & ", " & Order.BrokerID & ")"
        SendRanMessage eGDRanMessageType_CancelOrder, g.Broker.AccountNumberForID(Order.AccountID) & vbTab & Order.BrokerID
        
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
    RaiseError "cRan.ParkOrder"
    
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

    Dim Order As cPtOrder               ' Order to cancel
    
    Set Order = m.BrokerInfo.Order(strBrokerID, strGenesisID)
    If Not Order Is Nothing Then
        CancelOrder Order
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.CancelOrderFromBrokerID"
    
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
    
    Dim Orders As cPtOrders             ' Collection of orders to cancel
    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Number of orders cancelled
    
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
    RaiseError "cRan.CancelWorkingOrders"
    
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
    RaiseError "cRan.FlattenPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelBrokerOrder
'' Description: Attempt to cancel an order from the broker screen
'' Inputs:      Account Number, Broker Order Number
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelBrokerOrder(ByVal strAccountNumber As String, ByVal strBrokerOrderNumber As String)
On Error GoTo ErrSection:

    SendRanMessage eGDRanMessageType_CancelOrder, strAccountNumber & vbTab & strBrokerOrderNumber

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.CancelBrokerOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckHeartbeat
'' Description: Check the heartbeat to make sure that we are still connected
''              to the Rolfe & Nolan servers
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
    RaiseError "cRan.CheckHeartbeat"
    
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
            If m.astrAccounts.BinarySearch(strAccountNumber) Or ((strAccountNumber = "Alaron Broker Management") And (InStr(UCase(m.strUserName), "-ALR") <> 0)) Then
                nReturn = eGDConnectionStatus_Connected
            Else
                nReturn = eGDConnectionStatus_Disconnected
            End If
            
    End Select
    
    ConnectionStatusForAccount = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.ConnectionStatusForAccount"
    
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
    RaiseError "cRan.NextGenesisID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle the incoming message from the Rolfe & Nolan stand-alone app
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDRanMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    If (nType <> eGDRanMessageType_ConnectionInfo) And (nType <> eGDRanMessageType_Heartbeat) Then
        DumpDebug vbTab & m.strBrokerName & " Message Received (" & MessageTypeToString(nType) & "): " & strMessage
    End If

    Select Case nType
        Case eGDRanMessageType_ConnectionInfo
            HandleConnectionInfo strMessage
            If m.dLastFillTime > 0 Then
                If gdTickCount > m.dLastFillTime + 5000 Then
                    m.dLastFillTime = 0#
                    SynchronizePositions
                End If
            End If
            
        Case eGDRanMessageType_AppLoaded
            If DebugMode = False Then
                m.bAppLoaded = True
                If LiveTradingAllowed(m.nBroker) Then
                    DumpDebug "Connecting to " & m.strBrokerName & " as " & m.strUserName & " (" & m.strIP & ":" & m.strPort & ")"
                    If IsBroker Then
                        SendRanMessage eGDRanMessageType_Connect, m.strUserName & vbTab & m.strPassword & vbTab & "B" & vbTab & m.strIP & vbTab & m.strPort
                    Else
                        SendRanMessage eGDRanMessageType_Connect, m.strUserName & vbTab & m.strPassword & vbTab & "U" & vbTab & m.strIP & vbTab & m.strPort
                    End If
                Else
                    DumpDebug "Unloading the application since live trading is not allowed"
                    SendRanMessage eGDRanMessageType_UnloadApp, ""
                End If
            End If
        
        Case eGDRanMessageType_AppUnloaded
            m.bAppLoaded = False
            If m.bManualDisconnect = True Then
                m.dHeartBeat = 0#
                m.bVerifyPositions = True
            Else
                m.dHeartBeat = gdTickCount - 45000
            End If
            m.bManualDisconnect = False
        
        Case eGDRanMessageType_Account
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
                If m.bGetAccountInfo = True Then
                    RefreshAccountInfo ""
                Else
                    m.bGetAccountInfo = True
                End If
            Else
                HandleAccountInfo strMessage
                m.BrokerInfo.AddAccount Parse(strMessage, vbTab, 1) & vbTab
                m.BrokerInfo.SetConnectionStatusForAccount eGDConnectionStatus_Connected, Parse(strMessage, vbTab, 1)
            End If
            If FormIsLoaded("frmBrokerScreen") Then
                frmBrokerScreen.AccountCallback strMessage
            End If
            
        Case eGDRanMessageType_Contract
            m.dHeartBeat = gdTickCount
            
        Case eGDRanMessageType_Fill
            m.dHeartBeat = gdTickCount
            ' DAJ 08/02/2007: We are going to be pulling the fill information from the order
            ' callback, so we may not need to handle it here...
            'If UCase(strMessage) = "BEGIN" Then
            '    m.BrokerInfo.RefreshingFills = True
            '    If m.bVerboseRefresh Then
            '        InfBox "Refreshing " & m.strBrokerName & " fill informationcopy ta...", , , m.strBrokerName & " Refresh", True
            '    End If
            'ElseIf UCase(strMessage) = "END" Then
            '    m.BrokerInfo.RefreshingFills = False
            '    m.bRefreshingFills = False
            'Else
            '    'AddFillToBrokerInfo strMessage, m.bRefreshingFills
            '    HandleFillCallback strMessage, m.bRefreshingFills
            'End If
            
        Case eGDRanMessageType_Order
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingOrders = True
                m.BrokerInfo.RefreshingFills = True
                m.BrokerInfo.ClearOrders
                m.BrokerInfo.ClearFills
                m.BrokerInfo.ClearFillSummary
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " order and fill information...", , , m.strBrokerName & " Refresh", True
                End If
            ElseIf UCase(strMessage) = "END" Then
                m.BrokerInfo.RefreshingOrders = False
                m.BrokerInfo.RefreshingFills = False
                CleanOutOrderMap
                m.bRefreshingOrders = False
                m.bRefreshingFills = False
            Else
                strMessage = FixOrderType(strMessage)
                'UpdateBrokerInfoOrder strMessage, m.bRefreshingOrders
                HandleOrderCallback strMessage, m.bRefreshingOrders
            End If
            If FormIsLoaded("frmBrokerScreen") Then
                frmBrokerScreen.OrderCallback strMessage
            End If
                    
        Case eGDRanMessageType_Position
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
                m.BrokerInfo.RefreshingPositions = True
                m.BrokerInfo.ClearPositions
                m.astrBrokerPos.Clear
                If m.bVerboseRefresh Then
                    InfBox "Refreshing " & m.strBrokerName & " position information...", , , m.strBrokerName & " Refresh", True
                End If
            ElseIf UCase(strMessage) = "END" Then
                AddPositionsToBrokerInfo
                m.BrokerInfo.RefreshingPositions = False
                m.bRefreshingPositions = False
            Else
                ''AddPositionToBrokerInfo strMessage, m.bRefreshingPositions
                HandlePositionCallback strMessage
            End If
            If FormIsLoaded("frmBrokerScreen") Then
                frmBrokerScreen.PositionCallback strMessage
            End If
            
        Case eGDRanMessageType_SecurityList
            m.dHeartBeat = gdTickCount
            HandleSecurityInfo strMessage
            
        Case eGDRanMessageType_Statement
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
            ElseIf UCase(strMessage) = "END" Then
                HandleStatement ""
                m.bRefreshingStatement = False
            Else
                HandleStatement strMessage
            End If
            
        Case eGDRanMessageType_SystemMessage
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
            ElseIf UCase(strMessage) = "END" Then
                HandleSystemMessage ""
            Else
                HandleSystemMessage strMessage
            End If
            
        Case eGDRanMessageType_UserMessage
            m.dHeartBeat = gdTickCount
            If UCase(strMessage) = "BEGIN" Then
            ElseIf UCase(strMessage) = "END" Then
                HandleUserMessage ""
            Else
                HandleUserMessage strMessage
            End If
            
        Case eGDRanMessageType_AccountStatus
            m.dHeartBeat = gdTickCount
            HandleAccountStatusCallback strMessage
            m.bRefreshingAcctStatus = False
            
        Case eGDRanMessageType_Heartbeat
            m.dHeartBeat = gdTickCount
            
    End Select
    
    If Refreshing And (m.dRefreshingStart = 0#) Then
        m.dRefreshingStart = gdTickCount
    ElseIf Refreshing And (gdTickCount > m.dRefreshingStart + 30000#) Then
        m.bRefreshingAccounts = False
        m.bRefreshingOrders = False
        m.bRefreshingFills = False
        m.bRefreshingPositions = False
        m.bRefreshingAcctStatus = False
        m.bRefreshingStatement = False
        If m.bVerboseRefresh Then
            InfBox ""
        End If
    ElseIf (Not Refreshing) And (m.dRefreshingStart <> 0#) Then
        m.dRefreshingStart = 0#
        If m.bVerboseRefresh Then
            InfBox "Synchronizing " & m.strBrokerName & " position information...", , , m.strBrokerName & " Refresh", True
        End If
        SynchronizePositions
        If m.bVerboseRefresh Then
            InfBox ""
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.HandleMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendRanMessage
'' Description: Send an App Mail message to the Rolfe & Nolan program
'' Inputs:      Type of Message, Message, Send Now?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendRanMessage(ByVal nType As eGDRanMessageTypes, ByVal strMessage As String, Optional ByVal bSendNow As Boolean = True)
On Error GoTo ErrSection:

    Dim astrMessage As New cGdArray     ' Message split out into an array
    
    If nType = eGDRanMessageType_Connect Then
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
    RaiseError "cRan.SendRanMessage"
    
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

    If Not IsBroker Then
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
    RaiseError "cRan.VerifyPositions"
    
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

    Dim strBaseSymbol As String         ' Genesis base symbol
    
    strBaseSymbol = Parse(strGenesisSymbol, "-", 1)
    IsTradeableSymbol = m.tblSymbols.SearchAsIndex(m.aGenesisIdx, TblCol(eGDTblCol_GenesisSymbol), strBaseSymbol)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.IsTradeableSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsPitSymbol
'' Description: Is the given symbol in the conversion table as a pit symbol?
'' Inputs:      Symbol
'' Returns:     True if pit, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsPitSymbol(ByVal strGenesisSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim strBaseSymbol As String         ' Genesis base symbol
    Dim lPos As Long                    ' Position of symbol in the table
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    strBaseSymbol = Parse(strGenesisSymbol, "-", 1)
    If m.tblSymbols.SearchAsIndex(m.aGenesisIdx, TblCol(eGDTblCol_GenesisSymbol), strBaseSymbol, lPos) = True Then
        If m.tblSymbols.Item(TblCol(eGDTblCol_SymbolType), m.aGenesisIdx(lPos)) = "P" Then
            bReturn = True
        End If
    End If
    
    IsPitSymbol = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.IsPitSymbol"
    
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
        
        For lIndex = 0 To m.tblSymbols.NumRecords - 1
            If m.tblSymbols(TblCol(eGDTblCol_GenesisSymbol), lIndex) = strBaseSymbol Then
                strMask = m.tblSymbols(TblCol(eGDTblCol_OrderTypeMask), lIndex)
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
    RaiseError "cRan.OrderTypeAllowed"
    
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
    
    For lIndex = 0 To m.tblSymbols.NumRecords - 1
        If m.tblSymbols(TblCol(eGDTblCol_GenesisSymbol), lIndex) = strBaseSymbol Then
            strMask = m.tblSymbols(TblCol(eGDTblCol_TimeInForceMask), lIndex)
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
    RaiseError "cRan.TimeInForceAllowed"
    
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
    RaiseError "cRan.FixPositions"
    
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
            DumpDebug "Carried Position Mismatch: " & m.strBrokerName & " = " & Str(lBrokerCarried) & " ; Trade Navigator = " & Str(FillSummary.CurrentPosition)
            
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
    RaiseError "cRan.FixPosition"
    
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
    RaiseError "cRan.DumpDebug", eGDRaiseError_Raise
    
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
    Open AddSlash(m.strLogPath) & "TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append Shared As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " - " & strMessage
        Close #fh
    End If

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
    Dim strError As String              ' Error message back
    Static dLastAsk As Double           ' When did we last ask the user to reconnect?
    Static bInProgress As Boolean       ' Are we currently asking the user?
    Static strLastUser As String        ' Last user name
    
    dTickCount = Val(Parse(strConnectionInfo, vbTab, 1))
    nStatus = CLng(Val(Parse(strConnectionInfo, vbTab, 2)))
    strError = Parse(strConnectionInfo, vbTab, 3)
    
    Select Case nStatus
        Case eGDConnectionStatus_Disconnected
            m.BrokerInfo.SetConnectionStatus nStatus
            
            If ConnectionStatus <> eGDConnectionStatus_Disconnected Then
                m.dConnectTime = 0#
                m.astrAccounts.Clear
                
                If (Len(strError) > 0) Or m.bManualDisconnect Then
                    m.dLastStatusMsg = 0#
                    m.dHeartBeat = 0#
                    m.bVerifyPositions = True
                Else
                    m.dLastStatusMsg = dTickCount
                End If
                
                If m.bWasConnected = True Then
                    m.bWasConnected = False
                    g.Alerts.CheckBrokerStatusAlert m.nBroker
                End If
                
                If Len(strError) > 0 Then
                    InfBox strError, "!", "+-OK", m.strBrokerName & " Connection Error", True
                End If
                
                If InStr(UCase(strError), "ACCOUNT ACCESS DENIED") <> 0 Then
                    m.bAskedPassword = False
                    m.strPassword = ""
                End If
                
                If DebugMode = False Then
                    DumpDebug "Unloading application since we have been disconnected"
                    SendRanMessage eGDRanMessageType_UnloadApp, ""
                End If
            End If
        
        Case eGDConnectionStatus_Disconnecting
            m.BrokerInfo.SetConnectionStatusForUser nStatus, strLastUser
            m.dLastStatusMsg = dTickCount
            
        Case eGDConnectionStatus_Connecting
            m.BrokerInfo.SetConnectionStatusForUser nStatus, m.strUserName
            m.dLastStatusMsg = dTickCount
            
        Case eGDConnectionStatus_Connected
            m.dLastStatusMsg = dTickCount
            m.bWasConnected = True
            If (ConnectionStatus <> eGDConnectionStatus_Connected) Then
                m.dConnectTime = Now
                SetIniFileProperty "Last", EncryptToHex(Str(Date)), "Connect", m.strIniFile
                
                If (IsBroker = False) Then
                    GetAccounts
                End If
                
                GetUserMessages
                GetSystemMessages
            End If
            
    End Select
    
    If m.strUserName <> strLastUser Then
        strLastUser = m.strUserName
    End If
    
    If ConnectionStatus <> nStatus Then
        ConnectionStatus = nStatus
        
        g.ActivityLogs.AddStatusToActivityLog m.nBroker, nStatus
        
        DumpDebug vbTab & m.strBrokerName & " Message Received (" & MessageTypeToString(eGDRanMessageType_ConnectionInfo) & "): " & strConnectionInfo
    End If
    
    ' Since Rolfe & Nolan times out the order connection after 8 hours, we need to reconnect the user
    ' after 7 hours to make sure that we have continuous service...
    If (m.dConnectTime <> 0#) And (bInProgress = False) Then
        If Now > (m.dConnectTime + (7# / 24#)) Then
            If (dLastAsk = 0#) Or (Now > dLastAsk + (10# / 1440#)) Then
                bInProgress = True
                If InfBox("Trade Navigator needs to reconnect to the " & m.strBrokerName & " servers because the connection to the order server is about to time out.||NOT RECONNECTING NOW MAY RESULT IN TRADE NAVIGATOR NOT GETTING ORDER AND FILL MESSAGES FROM " & UCase(m.strBrokerName) & ".|", "!", "+Reconnect|-Cancel", m.strBrokerName & " Connection", , 20) = "R" Then
                    dLastAsk = 0#
                    DumpDebug "Reconnecting because of the 7 hour timeout (Started: " & DateFormat(m.dConnectTime, MM_DD_YY, HH_MM_SS, AMPM_UPPER) & ")"
                    Reconnect
                Else
                    dLastAsk = Now
                    DumpDebug "User asked about 7 hour timeout, but cancelled the dialog (Started: " & DateFormat(m.dConnectTime, MM_DD_YY, HH_MM_SS, AMPM_UPPER) & ")"
                End If
                bInProgress = False
            End If
        End If
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.HandleConnectionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleAccountInfo
'' Description: Handle account information from a callback
'' Inputs:      Account Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleAccountInfo(ByVal strAccountInfo As String)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strAccountNumber As String      ' Account Number from the callback
    Dim lAccountID As Long              ' Account ID
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the security found?

    If IsBroker = False Then
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
            
            RefreshAccount lAccountID
            
            If m.astrAccounts.BinarySearch(strAccountNumber) = False Then
                m.astrAccounts.Add strAccountNumber
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.HandleAccountInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleSecurityInfo
'' Description: Handle security information from a callback
'' Inputs:      Security Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleSecurityInfo(ByVal strSecurityInfo As String)
On Error GoTo ErrSection:

    If FormIsLoaded("frmBrokerScreen") Then
        frmBrokerScreen.SecurityCallback strSecurityInfo
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.HandleSecurityList"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleUserMessage
'' Description: Handle message from a callback
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleUserMessage(ByVal strUserMessage As String)
On Error GoTo ErrSection:

    Dim frm As frmMessage               ' Form to display the message in
    Static astrMessage As New cGdArray  ' Message to display to the user
    
    If Len(strUserMessage) > 0 Then
        astrMessage.Add strUserMessage
    Else
        Set frm = New frmMessage
        frm.ShowMe m.strBrokerName & " User Message", astrMessage.JoinFields(vbCrLf)
        astrMessage.Clear
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.HandleUserMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleSystemMessage
'' Description: Handle system message from a callback
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleSystemMessage(ByVal strSystemMessage As String)
On Error GoTo ErrSection:

    Dim frm As frmMessage               ' Form to display the message in
    Static astrMessage As New cGdArray  ' Message to display to the user
    
    If Len(strSystemMessage) > 0 Then
        astrMessage.Add strSystemMessage
    Else
        Set frm = New frmMessage
        frm.ShowMe m.strBrokerName & " System Message", astrMessage.JoinFields(vbCrLf)
        astrMessage.Clear
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.HandleSystemMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleStatement
'' Description: Handle statment from a callback
'' Inputs:      Statement
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleStatement(ByVal strStatement As String)
On Error GoTo ErrSection:

    Dim frm As frmMessage               ' Form to display the message in
    Static astrMessage As New cGdArray  ' Message to display to the user
        
    If Len(strStatement) > 0 Then
        astrMessage.Add strStatement
    Else
        If FormIsLoaded("frmBrokerScreen") Then
            frmBrokerScreen.StatementCallback astrMessage
        Else
            Set frm = New frmMessage
            frm.ShowMe m.strBrokerName & " Account Statment", astrMessage.JoinFields(vbCrLf)
        End If
        
        astrMessage.Clear
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.HandleStatement"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleOrderCallback
'' Description: Handle an order from a callback
'' Inputs:      Order String
'' Returns:     None
''
'' RecordFrom:  Fill Index, Order ID, Exchange Order ID, User Name, Account,
''              Exchange Name, Base Symbol, Contact, B/S, Price, Price2,
''              Quantity, Linked, Qty Filled, Status, Date Exch Recd, Time Exch Recd,
''              XRef, Expiration, Original Order ID, Credit, Currency, Desc,
''              Exchange#, Order Type1, Order Type2, Salesman, Session, P/C,
''              Parked, Last Price, Broker Symbol, P/L, Date Filled, Time Filled,
''              Fill Price, Fill Flag, Order Type String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleOrderCallback(ByVal strOrder As String, Optional ByVal bRefresh As Boolean = False)
On Error GoTo ErrSection:
    
    Dim Order As cPtOrder               ' Order object
    Dim astrOrder As New cGdArray       ' Array of order information from callback
    Dim strRanSymbol As String          ' Rolfe & Nolan symbol
    Dim strGenesisSymbol As String      ' Genesis version of Rolfe & Nolan symbol
    Dim dMult As Double                 ' Price multiplier
    Dim nPrevStatus As eTT_OrderStatus  ' Previous order status for the order
    Dim nOrderType As eTT_OrderType     ' Order type
    Dim frm As frmAlertPopup            ' Alert popup form
    Dim Bars As New cGdBars             ' Bars object
    Dim lPos As Long                    ' Position in an array
    Dim Fill As cPtFill                 ' Fill object
    Dim strFillID As String             ' Fill ID (Order ID - Fill Seq)
    Dim bFillReceived As Boolean        ' Did we receive a fill?
    Dim bFillExisted As Boolean         ' Did the fill already exist?
    Dim bFillChanged As Boolean         ' Did the fill change?
    Dim lIndex As Long                  ' Index into a for loop
    Dim dFillTime As Double             ' Old Fill Time
    Dim dFillPrice As Double            ' Old Fill Price
    Dim lFillQty As Long                ' Old Fill Quantity
    Dim PrevOrder As cPtOrder           ' Previous order
    
    If IsBroker = False Then
        astrOrder.SplitFields strOrder, vbTab
        
        strRanSymbol = astrOrder(6) & "-" & astrOrder(7) & "-" & astrOrder(28)
        If astrOrder(17) = "NONE" Then astrOrder(17) = ""
        If Len(strRanSymbol) = 2 Then
            If (GenesisOrderStatus(astrOrder(14)) = eTT_OrderStatus_AmendPending) Then
                If m.astrAmendPending.BinarySearch(astrOrder(19) & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
                    Set Order = LoadOrder(astrOrder(4), Parse(m.astrAmendPending(lPos), vbTab, 3), astrOrder(19), "")
                    If Not Order Is Nothing Then
                        Order.Status = eTT_OrderStatus_Amended
                        Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
                        Order.Save
                        
                        m.BrokerInfo.AddOrder Order, bRefresh
                        
                        OrderCallback Order
                        If bRefresh = False Then
                            g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                        End If
                    
                        DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                    End If
                
                    Set Order = LoadOrder(astrOrder(4), Parse(m.astrAmendPending(lPos), vbTab, 5), astrOrder(1), "")
                    If Not Order Is Nothing Then
                        Order.BrokerID = astrOrder(1)
                        Order.Status = eTT_OrderStatus_Sent
                        Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
                        Order.Save
                        
                        m.BrokerInfo.AddOrder Order, bRefresh
                        
                        OrderCallback Order
                        If bRefresh = False Then
                            g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                        End If
                    
                        DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                    End If
                    
                    OrderIDChanged CLng(Val(Parse(m.astrAmendPending(lPos), vbTab, 2))), CLng(Val(Parse(m.astrAmendPending(lPos), vbTab, 4)))
                    m.astrAmendPending.Remove lPos
                End If
            Else
                Set Order = LoadOrder(astrOrder(4), astrOrder(17), astrOrder(1), astrOrder(19))
                If Not Order Is Nothing Then
                    nPrevStatus = Order.Status
                    Order.Status = GenesisOrderStatus(astrOrder(14))
                    If (Order.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_ParkPending) Then
                        Order.Status = eTT_OrderStatus_Parked
                    ElseIf (Order.Status = eTT_OrderStatus_Queued) And (nPrevStatus = eTT_OrderStatus_Working) Then
                        Order.Status = eTT_OrderStatus_Working
                    End If
                    Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, strGenesisSymbol)
                    Order.Save
                    
                    m.BrokerInfo.AddOrder Order, bRefresh
                    
                    OrderCallback Order
                    g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                End If
            End If
        ElseIf GenesisSymbol(strRanSymbol, strGenesisSymbol, dMult) Then
            SetBarProperties Bars, strGenesisSymbol
            If GenesisOrderStatus(astrOrder(14)) = eTT_OrderStatus_Rejected Then
                Set Order = LoadOrder(astrOrder(4), astrOrder(17), "", "")
            Else
                Set Order = LoadOrder(astrOrder(4), astrOrder(17), astrOrder(1), astrOrder(19))
            End If
            If Order Is Nothing Then Set Order = New cPtOrder
            
            If Len(astrOrder(0)) > 0 Then
                strFillID = astrOrder(1) & "-" & astrOrder(0)
            Else
                strFillID = ""
            End If
            
            With Order
                nPrevStatus = .Status
                
                .AccountID = g.Broker.AccountIDForNumber(astrOrder(4))
                If GenesisOrderStatus(astrOrder(14)) <> eTT_OrderStatus_Rejected Then
                    .BrokerID = astrOrder(1)
                End If
                .Buy = (UCase(astrOrder(8)) = "B")
                .DateExchRecd = DateFromString(astrOrder(15), astrOrder(16))
                .ExchangeID = astrOrder(2)
                If Len(Trim(astrOrder(18))) = 0 Then
                    .Expiration = -1&
                ElseIf UCase(Trim(astrOrder(18))) = "GTC" Then
                    .Expiration = 0&
                Else
                    .Expiration = CLng(Val(astrOrder(18)))
                End If
                .GenesisOrderID = astrOrder(17)
                .Message = astrOrder(39)
                If .OrderDate = 0 Then .OrderDate = ConvertToBrokerDate(CurrentTime, m.nBroker, strGenesisSymbol)
                If UCase(astrOrder(27)) = "E" Then .Session = "E" Else .Session = "P"
                
                nOrderType = GenesisOrderType(astrOrder(37))
                If Not ((nOrderType = eTT_OrderType_Market) And (.OrderType = eTT_OrderType_MIT)) Then
                    If nOrderType <> -1& Then .OrderType = nOrderType
                    Select Case .OrderType
                        Case eTT_OrderType_Market
                            .LimitPrice = 0#
                            .StopPrice = 0#
                        Case eTT_OrderType_Stop
                            .LimitPrice = 0#
                            .StopPrice = GenesisPrice(Val(astrOrder(9)), dMult)
                        Case eTT_OrderType_Limit
                            .LimitPrice = GenesisPrice(Val(astrOrder(9)), dMult)
                            .StopPrice = 0#
                        Case eTT_OrderType_StopWithLimit
                            .LimitPrice = GenesisPrice(Val(astrOrder(10)), dMult)
                            .StopPrice = GenesisPrice(Val(astrOrder(9)), dMult)
                    End Select
                End If
                .Quantity = CLng(Val(astrOrder(11)))
                
                If Len(strFillID) > 0 Then
                    bFillReceived = True
                    bFillExisted = False
                    bFillChanged = True
                    
                    dFillTime = CalendarDate(DateFromString(astrOrder(33), astrOrder(34)), strGenesisSymbol, Bars)
                    dFillPrice = GenesisPrice(Val(astrOrder(35)), dMult)
                    lFillQty = CLng(Val(astrOrder(13)))
                    
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
                        .FillDate = dFillTime
                        .IsManual = False
                        .IsSnapshot = True
                        .OrderID = Order.OrderID
                        .Price = dFillPrice
                        .Quantity = lFillQty
                        .SymbolOrSymbolID = strGenesisSymbol
                        .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.FillDate, m.nBroker, .Symbol, False))
                        
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
                
                .Status = GenesisOrderStatus(astrOrder(14))
                If (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_ParkPending) Then
                    .Status = eTT_OrderStatus_Parked
                ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_Parked) Then
                    .Status = eTT_OrderStatus_Parked
                ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_AmendPending) Then
                    .Status = eTT_OrderStatus_Amended
                ElseIf (.Status = eTT_OrderStatus_Cancelled) And (nPrevStatus = eTT_OrderStatus_Amended) Then
                    .Status = eTT_OrderStatus_Amended
                End If
                .StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, strGenesisSymbol)
                .SymbolOrSymbolID = strGenesisSymbol
                
                .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(Order.OrderDate, m.nBroker, Order.Symbol, False))
                .IsSnapshot = True
                If Len(astrOrder(19)) > 0 Then
                    .PreviousBrokerID = astrOrder(19)
                    If (Len(.PreviousBrokerID) > 0) And (.PreviousBrokerID <> .BrokerID) Then
                        Set PrevOrder = New cPtOrder
                        PrevOrder.LoadForBrokerID .PreviousBrokerID
                        .PreviousOrder = PrevOrder
                    End If
                End If
                
                .Save
                                                
                If (nPrevStatus <> Order.Status) And ((Order.Status = eTT_OrderStatus_Error) Or (Order.Status = eTT_OrderStatus_Rejected)) Then
                    If Len(astrOrder(39)) > 0 Then
                        If Len(astrOrder(1)) > 0 Then
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox "Please contact your broker and report the following problem:||Order ID: " & astrOrder(1) & "|" & "Error: " & astrOrder(39) & "|", m.strBrokerName & " Order Error", vbLeftJustify
                        Else
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox "Please contact your broker and report the following problem:||Error: " & astrOrder(39) & "|", m.strBrokerName & " Order Error", vbLeftJustify
                        End If
                    Else
                        If Len(astrOrder(1)) > 0 Then
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox "Please contact your broker and report the following problem:||Order ID " & astrOrder(1) & " was rejected.|", m.strBrokerName & " Order Error", vbLeftJustify
                        Else
                            Set frm = New frmAlertPopup
                            frm.ShowMessageBox "Please contact your broker and report the following problem:||Order was rejected.|", m.strBrokerName & " Order Error", vbLeftJustify
                        End If
                    End If
                End If
                
                If bFillReceived Then
                    m.BrokerInfo.AddOrder Order, bRefresh
                    m.BrokerInfo.AddFill Fill, bRefresh
                    FillCallback Order, Fill, bFillExisted, True, bFillChanged
                    If bRefresh = False Then
                        g.ActivityLogs.AddFillToActivityLog m.nBroker, Fill, False
                    End If
                    m.dLastFillTime = gdTickCount
                Else
                    m.BrokerInfo.AddOrder Order, bRefresh
                    OrderCallback Order, False
                    If bRefresh = False Then
                        g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                    End If
                End If
                DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                
                If .Status = eTT_OrderStatus_Rejected Then
                    For lPos = 0 To m.astrAmendPending.Size - 1
                        '' Broker ID, Old Access and Genesis ID's, New Access and Genesis ID's
                        If Parse(m.astrAmendPending(lPos), vbTab, 5) = .GenesisOrderID Then
                            Set Order = LoadOrder(astrOrder(4), Parse(m.astrAmendPending(lPos), vbTab, 3), Parse(m.astrAmendPending(lPos), vbTab, 1), "")
                            If Not Order Is Nothing Then
                                Order.GenesisOrderID = Parse(m.astrAmendPending(lPos), vbTab, 3)
                                Order.Status = eTT_OrderStatus_Working
                                Order.Message = ""
                                Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.nBroker, Order.Symbol)
                                Order.Save
                                
                                m.BrokerInfo.AddOrder Order, bRefresh
                                
                                OrderCallback Order
                                If bRefresh = False Then
                                    g.ActivityLogs.AddOrderToActivityLog m.nBroker, Order
                                End If
                            
                                DumpDebug Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & "): " & OrderStatus(Order.Status)
                            End If
                            
                            m.astrAmendPending.Remove lPos
                            
                            Exit For
                        End If
                    Next lPos
                End If
            End With
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.HandleOrderCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleFillCallback
'' Description: Handle a fill from a callback
'' Inputs:      Fill String
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleFillCallback(ByVal strFill As String, Optional ByVal bRefresh As Boolean = False)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    Dim Fill As cPtFill                 ' Fill object
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrFill As New cGdArray        ' Array of order information from callback
    Dim strRanSymbol As String          ' Rolfe & Nolan symbol
    Dim strGenesisSymbol As String      ' Genesis version of Rolfe & Nolan symbol
    Dim dMult As Double                 ' Price multiplier
    Dim bFillExisted As Boolean         ' Did the fill exist?
    Dim Bars As New cGdBars             ' Bars object
    Dim strFillID As String             ' Broker ID for the fill
    
    If IsBroker = False Then
        astrFill.SplitFields strFill, vbTab
        
        If (Len(astrFill(5)) > 0) And (Len(astrFill(6)) > 0) Then
            strRanSymbol = astrFill(2) & "-" & astrFill(3) & "-" & astrFill(17)
            If GenesisSymbol(strRanSymbol, strGenesisSymbol, dMult) Then
                SetBarProperties Bars, strGenesisSymbol
                Set Order = LoadOrder(astrFill(14), astrFill(16), astrFill(7), "")
                If Not Order Is Nothing Then
                    Set Fill = Nothing
                    bFillExisted = False
                    
                    strFillID = astrFill(7) & "-" & astrFill(0)
                    
                    For lIndex = 1 To Order.Fills.Count
                        If Order.Fills(lIndex).BrokerID = strFillID Then
                            Set Fill = Order.Fills(lIndex)
                            bFillExisted = True
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
                        .FillDate = CalendarDate(DateFromString(astrFill(8), astrFill(9)), strGenesisSymbol, Bars)
                        .IsManual = False
                        .IsSnapshot = True
                        .OrderID = Order.OrderID
                        .Price = GenesisPrice(Val(astrFill(6)), dMult)
                        .Quantity = CLng(Val(astrFill(5)))
                        .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.FillDate, m.nBroker, .Symbol, False))
                        .Symbol = Order.Symbol
                        .SymbolID = Order.SymbolID
                        
                        .Save
                    End With
                    
                    Set Order.Fills(Str(Fill.FillID)) = Fill
                    Order.Save
                    
                    m.BrokerInfo.AddOrder Order, bRefresh
                    m.BrokerInfo.AddFill Fill, bRefresh
                    
                    FillCallback Order, Fill, bFillExisted, False
                    g.ActivityLogs.AddFillToActivityLog m.nBroker, Fill, False
                    m.dLastFillTime = gdTickCount
                    DumpDebug Fill.FillText
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.HandleFillCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleAccountStatusCallback
'' Description: Handle an account status from a callback
'' Inputs:      Account Status String
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleAccountStatusCallback(ByVal strAccountStatus As String)
On Error GoTo ErrSection:

    If FormIsLoaded("frmBrokerScreen") Then
        frmBrokerScreen.AccountStatusCallback strAccountStatus
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.HandleAccountStatusCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandlePositionCallback
'' Description: Handle a position from a callback
'' Inputs:      Position String
'' Returns:     None
''
'' Record From: Account, Contract Name, Contract Date, B/S, XrefPersist,
''              Currency, Description, Salesman, P/C Strike, Last Tick,
''              Broker Symbol, P/L, Quantity, Price, Date, Account Name, Reg Code
'' RecordTo:    Account, Symbol, Position, Average Entry, Entry Prices
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandlePositionCallback(ByVal strPosition As String)
On Error GoTo ErrSection:

    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim strRanSymbol As String          ' Rolfe & Nolan symbol
    Dim dMult As Double                 ' Price multiplier
    Dim lPos As Long                    ' Position in the array
    Dim astrPosition As cGdArray        ' Array of position information
    Dim adEntryPrices As cGdArray       ' Array of entry prices
    Dim lRanPosition As Long            ' Rolfe & Nolan position
    Dim lRanPos As Long                 ' Rolfe & Nolan position
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrBrokerPos As cGdArray       ' Broker position information
    Dim dGenesisPrice As Double         ' Genesis price for the position
    Dim strID As String                 ' ID for the position
    Dim astrCarried As cGdArray         ' Array of carried position information
    Dim bFound As Boolean               ' Was the ID found in the array?

    Set astrPosition = New cGdArray
    astrPosition.Create eGDARRAY_Strings
    Set adEntryPrices = New cGdArray
    adEntryPrices.Create eGDARRAY_Doubles
    Set astrBrokerPos = New cGdArray
    astrBrokerPos.Create eGDARRAY_Strings
    Set astrCarried = New cGdArray
    astrCarried.Create eGDARRAY_Strings

    If IsBroker = False Then
        If Len(strPosition) > 0 Then
            astrPosition.SplitFields strPosition, vbTab
            
            strRanSymbol = astrPosition(1) & "-" & astrPosition(2) & "-" & astrPosition(8)
            If GenesisSymbol(strRanSymbol, strGenesisSymbol, dMult) Then
                If UCase(astrPosition(3)) = "B" Then
                    lRanPosition = CLng(Val(astrPosition(12)))
                Else
                    lRanPosition = CLng(Val(astrPosition(12))) * -1&
                End If
                lRanPos = lRanPosition
                
                strID = astrPosition(14) & "-" & astrPosition(10) & "-" & astrPosition(13)
                
                If m.astrBrokerPos.BinarySearch(astrPosition(0) & vbTab & strGenesisSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
                    dGenesisPrice = GenesisPrice(astrPosition(13), dMult)
                    
                    astrBrokerPos.SplitFields m.astrBrokerPos(lPos), vbTab
                    
                    adEntryPrices.SplitFields astrBrokerPos(5), ","
                    For lIndex = 1 To Abs(lRanPosition)
                        adEntryPrices.Add dGenesisPrice
                    Next lIndex
                    lRanPosition = lRanPosition + CLng(Val(astrBrokerPos(3)))
                    
                    bFound = False
                    astrCarried.SplitFields astrBrokerPos(6), ","
                    For lIndex = 0 To astrCarried.Size - 1
                        If Parse(astrCarried(lIndex), ";", 3) = strID Then
                            astrCarried(lIndex) = Str(CLng(Val(Parse(astrCarried(lIndex), ";", 1))) + lRanPos) & ";" & Parse(astrCarried(lIndex), ";", 2) & ";" & Parse(astrCarried(lIndex), ";", 3)
                            bFound = True
                            Exit For
                        End If
                    Next lIndex
                    
                    If bFound = False Then
                        astrCarried.Add Str(lRanPos) & ";" & Str(dGenesisPrice) & ";" & strID
                    End If
                    
                    m.astrBrokerPos(lPos) = astrBrokerPos(0) & vbTab & astrBrokerPos(1) & vbTab & astrBrokerPos(2) & vbTab & Str(lRanPosition) & vbTab & Str(adEntryPrices.CalcStatistic(eGdStat_Average)) & vbTab & adEntryPrices.JoinFields(",") & vbTab & astrCarried.JoinFields(",")
                    
                    DumpDebug "Position Updated: " & m.astrBrokerPos(lPos)
                Else
                    dGenesisPrice = GenesisPrice(astrPosition(13), dMult)
                    
                    adEntryPrices.Clear
                    For lIndex = 1 To Abs(lRanPosition)
                        adEntryPrices.Add dGenesisPrice
                    Next lIndex
                    m.astrBrokerPos.Add astrPosition(0) & vbTab & strGenesisSymbol & vbTab & "" & vbTab & Str(lRanPosition) & vbTab & Str(adEntryPrices.CalcStatistic(eGdStat_Average)) & vbTab & adEntryPrices.JoinFields(",") & vbTab & Str(lRanPos) & ";" & Str(dGenesisPrice) & ";" & strID, lPos
                    
                    DumpDebug "New Position Added: " & m.astrBrokerPos(lPos)
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.HandlePositionCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SynchronizePositions
'' Description: Synchronize the Genesis positions with the Rolfe & Nolan positions
'' Inputs:      None
'' Returns:     None
''
'' FillSummary: Account, Symbol, At ID, Buys, Sells, Net, Total, PriceSum, Entries,
''              ClosedProfit, AvgEntry, Initial Fill Price, Initial Fill Date,
''              Session Date, Last Traded, Overnight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SynchronizePositions(Optional ByVal strSymbol As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPosition As Long               ' Position from the broker
    Dim strPosition As String           ' Position information
    Dim lSymbolID As Long               ' Symbol ID
    Dim frm As frmAlertPopup            ' Alert popup form
    Dim FillSumms As cAccountPositions  ' Collection of fill summaries
    Dim strAccount As String            ' Account number
    Dim lBrokerPos As Long              ' Broker position
    Dim lBrokerCarried As Long          ' Broker carried position
    Dim lPos As Long                    ' Position in an array
    
    m.bSyncInProgress = True
        
    BuildPositions
    
    If Not IsBroker Then
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
    End If
    
    If m.astrToFix.Size > 0 Then
        FixPositions
    End If
    
    m.bSyncInProgress = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.SynchronizePositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RanOrder
'' Description: Convert a Genesis order to a string of Rolfe & Nolan order information
'' Inputs:      Genesis Order
'' Returns:     Rolfe & Nolan Order String (empty if problem)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RanOrder(ByVal Order As cPtOrder) As String
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information
    Dim strRanSymbol As String          ' Rolfe & Nolan version of the Genesis symbol
    Dim strRanExchange As String        ' Rolfe & Nolan exchange for Genesis symbol
    Dim dMult As Double                 ' Multiplier
    
    astrOrder.Create eGDARRAY_Strings

    If (Order.OrderType = eTT_OrderType_MIT) And (g.Broker.SyntheticOrders.Exists(Str(Order.OrderID)) = True) Then
        Order.OrderType = eTT_OrderType_Market
    End If
    
    If RanSymbol(Order.Symbol, strRanSymbol, strRanExchange, dMult) Then
        With Order
            astrOrder(0) = g.Broker.AccountNumberForID(.AccountID)
            astrOrder(1) = RanOrderType(.OrderType)
            astrOrder(2) = strRanExchange
            astrOrder(3) = Parse(strRanSymbol, "-", 1)
            astrOrder(4) = Parse(strRanSymbol, "-", 2)
            'astrOrder(5) = Parse(strRanSymbol, "-", 3)
            If UCase(.Session) = "E" Then astrOrder(5) = "E" Else astrOrder(5) = "P"
            If .Buy Then astrOrder(6) = "B" Else astrOrder(6) = "S"
            Select Case .OrderType
                Case eTT_OrderType_Market
                    astrOrder(7) = ""
                    astrOrder(8) = ""
                Case eTT_OrderType_Stop
                    astrOrder(7) = Str(RanPrice(Order.StopPrice, dMult))
                    astrOrder(8) = ""
                Case eTT_OrderType_Limit
                    astrOrder(7) = Str(RanPrice(Order.LimitPrice, dMult))
                    astrOrder(8) = ""
                Case eTT_OrderType_StopWithLimit
                    astrOrder(7) = Str(RanPrice(Order.StopPrice, dMult))
                    astrOrder(8) = Str(RanPrice(Order.LimitPrice, dMult))
            End Select
            astrOrder(9) = Str(Order.Quantity)
            astrOrder(10) = ""
            astrOrder(11) = Order.GenesisOrderID
            
            If Order.Expiration = 0& Then
                astrOrder(12) = "GTC"
            ElseIf Order.Expiration < 0& Then
                astrOrder(12) = "DAY"
            Else
                astrOrder(12) = Format("MMDDYYYY", Order.Expiration)
            End If
            
            astrOrder(13) = ""
            astrOrder(14) = ""
            astrOrder(15) = ""
            astrOrder(16) = Order.BrokerID
        End With
    
        RanOrder = astrOrder.JoinFields(vbTab)
    Else
        RanOrder = ""
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.RanOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RanOrderType
'' Description: Convert a Genesis Order Type to an Rolfe & Nolan order type
'' Inputs:      Genesis Order Type
'' Returns:     Rolfe & Nolan Order Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RanOrderType(ByVal nOrderType As eTT_OrderType) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case nOrderType
        Case eTT_OrderType_Market
            strReturn = "Market"
        Case eTT_OrderType_Stop
            strReturn = "Stop"
        Case eTT_OrderType_Limit
            strReturn = "Limit"
        Case eTT_OrderType_StopWithLimit
            strReturn = "StopWithLimit"
    End Select
    
    RanOrderType = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.RanOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RanSymbol
'' Description: Convert a Genesis symbol to an Rolfe & Nolan symbol
'' Inputs:      Genesis Symbol
'' Returns:     Rolfe & Nolan Symbol, Rolfe & Nolan Exchange, Price Multiplier
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RanSymbol(ByVal strGenesisSymbol As String, strRanSymbol As String, strRanExchange As String, dMult As Double) As Boolean
On Error GoTo ErrSection:

    Dim strBaseSymbol As String         ' Genesis base symbol
    Dim lPos As Long                    ' Position of symbol in the table
    
    strBaseSymbol = Parse(strGenesisSymbol, "-", 1)
    If m.tblSymbols.SearchAsIndex(m.aGenesisIdx, TblCol(eGDTblCol_GenesisSymbol), strBaseSymbol, lPos) = True Then
        strRanSymbol = m.tblSymbols(TblCol(eGDTblCol_RanSymbol), m.aGenesisIdx(lPos))
        strRanSymbol = strRanSymbol & "-" & Parse(strGenesisSymbol, "-", 2)
        strRanSymbol = strRanSymbol & "-" & m.tblSymbols(TblCol(eGDTblCol_Session), m.aGenesisIdx(lPos))
        strRanExchange = m.tblSymbols(TblCol(eGDTblCol_Exchange), m.aGenesisIdx(lPos))
        dMult = m.tblSymbols.Num(TblCol(eGDTblCol_Multiplier), m.aGenesisIdx(lPos))
        RanSymbol = True
    Else
        RanSymbol = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.RanSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RanPrice
'' Description: Convert a Genesis price to an Rolfe & Nolan price
'' Inputs:      Genesis Price, Multiplier
'' Returns:     Rolfe & Nolan Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RanPrice(ByVal dGenesisPrice As Double, ByVal dMult As Double) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value from the function

    Select Case Abs(dMult)
        Case 8
            dReturn = CDbl(Int(dGenesisPrice) * 10) + CDbl((dGenesisPrice - Int(dGenesisPrice)) * 8)
        Case 16
            dReturn = CDbl(Int(dGenesisPrice) * 100) + CDbl((dGenesisPrice - Int(dGenesisPrice)) * 16)
        Case 32
            dReturn = CDbl(Int(dGenesisPrice) * 100) + CDbl((dGenesisPrice - Int(dGenesisPrice)) * 32)
        Case 32.25
            dReturn = CDbl(Int(dGenesisPrice) * 1000) + CDbl(Int((dGenesisPrice - Int(dGenesisPrice)) * 320))
        Case 32.5
            dReturn = CDbl(Int(dGenesisPrice) * 1000) + CDbl((dGenesisPrice - Int(dGenesisPrice)) * 320)
        Case 64
            dReturn = CDbl(Int(dGenesisPrice) * 100) + CDbl((dGenesisPrice - Int(dGenesisPrice)) * 64)
        Case 100.5
            dReturn = CDbl(Int(dGenesisPrice * 100))
        Case Else
            dReturn = dGenesisPrice * Abs(dMult)
    End Select
    
    If dMult < 0 Then
        RanPrice = Val(Int(dReturn))
    Else
        RanPrice = dReturn
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.RanPrice"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Convert an Rolfe & Nolan symbol and exchange to a Genesis symbol
'' Inputs:      Rolfe & Nolan Symbol
'' Returns:     Genesis Symbol, Multiplier
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbol(ByVal strRanSymbol As String, strGenesisSymbol As String, dMult As Double) As Boolean
On Error GoTo ErrSection:

    Dim strBaseSymbol As String         ' Rolfe & Nolan base symbol
    Dim lPos As Long                    ' Position in the table
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    
    If Len(strRanSymbol) > 0 Then
        If Len(Parse(strRanSymbol, "-", 3)) = 0 Then
            strBaseSymbol = Parse(strRanSymbol, "-", 1)
            If m.tblSymbols.SearchAsIndex(m.aRanIdx, TblCol(eGDTblCol_RanSymbol), strBaseSymbol, lPos) Then
                strGenesisSymbol = m.tblSymbols(TblCol(eGDTblCol_GenesisSymbol), m.aRanIdx(lPos))
                strGenesisSymbol = strGenesisSymbol & "-" & Parse(strRanSymbol, "-", 2)
                dMult = m.tblSymbols.Num(TblCol(eGDTblCol_Multiplier), m.aRanIdx(lPos))
                bReturn = True
            End If
        End If
    End If
    
    GenesisSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.GenesisSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderType
'' Description: Convert an Rolfe & Nolan order type to a Genesis order type
'' Inputs:      Rolfe & Nolan Order Type
'' Returns:     Genesis Order Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderType(ByVal strRanOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Select Case UCase(strRanOrderType)
        Case "MARKET"
            GenesisOrderType = eTT_OrderType_Market
        Case "MARKETONOPEN"
            GenesisOrderType = eTT_OrderType_Market
        Case "MARKETONCLOSE"
            GenesisOrderType = eTT_OrderType_MarketOnClose
        Case "LIMIT"
            GenesisOrderType = eTT_OrderType_Limit
        Case "STOP"
            GenesisOrderType = eTT_OrderType_Stop
        Case "STOPONOPEN"
            GenesisOrderType = eTT_OrderType_Stop
        Case "STOPCLOSEONLY"
            GenesisOrderType = eTT_OrderType_StopCloseOnly
        Case "STOPANDLIMIT"
            GenesisOrderType = eTT_OrderType_Stop
        Case "STOPWITHLIMIT"
            GenesisOrderType = eTT_OrderType_StopWithLimit
        Case "STOPLLIMITCLOSEONLY"
            GenesisOrderType = eTT_OrderType_StopWithLimitCloseOnly
        
        'Case "FILLORKILL"
        'Case "MARKETIFTOUCHED"
        'Case "ORBETTER"
        'Case "ORBETTERONOPEN"
        'Case "ORBETTERCLOSEONLY"
        'Case "UNKNOWN"
        
        Case Else
            GenesisOrderType = -1&
        
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.GenesisOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisPrice
'' Description: Convert an Rolfe & Nolan price to a Genesis price
'' Inputs:      Rolfe & Nolan Price, Multiplier
'' Returns:     Genesis Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisPrice(ByVal dRanPrice As Double, ByVal dMult As Double) As Double
On Error GoTo ErrSection:

    If (dMult < 0) Then
        If ((Int(dRanPrice) Mod 10) = 2) Or ((Int(dRanPrice) Mod 10) = 7) Then
            dRanPrice = dRanPrice + 0.5
        End If
    End If

    Select Case Abs(dMult)
        Case 0
            GenesisPrice = 0#
        Case 8
            GenesisPrice = CDbl(Int(dRanPrice / 10)) + ((dRanPrice Mod 10) / 8)
        Case 16
            GenesisPrice = CDbl(Int(dRanPrice / 100)) + ((dRanPrice Mod 100) / 16)
        Case 32
            GenesisPrice = CDbl(Int(dRanPrice / 100)) + ((dRanPrice Mod 100) / 32)
        Case 32.25
            If ((Int(dRanPrice) Mod 10) = 2) Or ((Int(dRanPrice) Mod 10) = 7) Then
                dRanPrice = dRanPrice + 0.5
            End If
            GenesisPrice = CDbl(Int(dRanPrice / 1000)) + ((dRanPrice Mod 1000) / 320)
        Case 32.5
            GenesisPrice = CDbl(Int(dRanPrice / 1000)) + ((dRanPrice Mod 1000) / 320)
        Case 64
            GenesisPrice = CDbl(Int(dRanPrice / 100)) + ((dRanPrice Mod 100) / 64)
        Case 100.5
            If (dRanPrice Mod 10 = 2) Or (dRanPrice Mod 10 = 7) Then
                GenesisPrice = (dRanPrice + 0.5) / 100
            Else
                GenesisPrice = dRanPrice / 100
            End If
        Case Else
            GenesisPrice = dRanPrice / Abs(dMult)
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.GenesisPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderStatus
'' Description: Convert an Rolfe & Nolan order status to a Genesis order status
'' Inputs:      Rolfe & Nolan Order Status
'' Returns:     Genesis Order Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderStatus(ByVal strRanStatus As String) As eTT_OrderStatus
On Error GoTo ErrSection:

    Select Case UCase(strRanStatus)
        'Case "OPEN"
        '    GenesisOrderStatus = eTT_OrderStatus_Open
        Case "PARTIALLY FILLED"
            GenesisOrderStatus = eTT_OrderStatus_Partial
        Case "FILLED"
            GenesisOrderStatus = eTT_OrderStatus_Filled
        Case "CANCELED"
            GenesisOrderStatus = eTT_OrderStatus_Cancelled
        Case "PENDING NEW", "UNKNOWN STATUS"
            GenesisOrderStatus = eTT_OrderStatus_Queued
        'Case "SENT"
        '    GenesisOrderStatus = eTT_OrderStatus_Sent
        Case "NEW ORDER", "REPLACED"
            GenesisOrderStatus = eTT_OrderStatus_Working
        Case "REJECTED"
            GenesisOrderStatus = eTT_OrderStatus_Rejected
        'Case "BALCANCELLED"
        '    GenesisOrderStatus = eTT_OrderStatus_BalCancelled
        Case "PENDING CANCEL"
            GenesisOrderStatus = eTT_OrderStatus_CancelPending
        Case "PENDING REPLACE"
            GenesisOrderStatus = eTT_OrderStatus_AmendPending
        Case "SUSPENDED"
            GenesisOrderStatus = eTT_OrderStatus_Held
        'Case "CANCELHELDORDER"
        '    GenesisOrderStatus = eTT_OrderStatus_CancelHeld
        Case "EXPIRED"
            GenesisOrderStatus = eTT_OrderStatus_Expired
        Case "ACCEPTED", "ADAPTER ACCEPTED", "ACCEPTED FOR BIDDING"
            GenesisOrderStatus = eTT_OrderStatus_Approved
        Case "PARKED"
            GenesisOrderStatus = eTT_OrderStatus_BrokerParked
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.GenesisOrderStatus"
    
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
    RaiseError "cRan.ExpirationDateForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateFromString
'' Description: Convert a string formatted date to a double
'' Inputs:      Date, Time
'' Returns:     Julian Date/Time
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DateFromString(ByVal strDate As String, ByVal strTime As String) As Double
On Error GoTo ErrSection:

    Dim dTime As Double                 ' Time portion of the julian
    Dim dReturn As Double               ' Return value from the function
    Dim dCurrentTime As Double          ' Current time
    
    If Len(strTime) > 0 Then
        dTime = (Val(Mid(strTime, 1, 2)) * 3600) + (Val(Mid(strTime, 3, 2)) * 60) + Val(Mid(strTime, 5, 2))
        dTime = dTime / 86400#
    End If
    
    If Len(strDate) > 0 Then
        dReturn = JulFromLong(CLng(strDate)) + dTime
    Else
        dReturn = dTime
    End If
    
    DateFromString = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.DateFromString"
    
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
        If (Len(strBrokerID) > 0) Then
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
        If (bFound = False) And Len(strGenesisOrderID) > 0 Then
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
    End If
    
    If bFound Then
        Set LoadOrder = Order
    Else
        Set LoadOrder = Nothing
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.LoadOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MessageTypeToString
'' Description: Convert a message type enumeration to a string
'' Inputs:      Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MessageTypeToString(ByVal nType As eGDRanMessageTypes) As String
On Error GoTo ErrSection:

    Select Case nType
        Case eGDRanMessageType_Connect
            MessageTypeToString = "eGDRanMessageType_Connect"
        Case eGDRanMessageType_Disconnect
            MessageTypeToString = "eGDRanMessageType_Disconnect"
        Case eGDRanMessageType_AddOrder
            MessageTypeToString = "eGDRanMessageType_AddOrder"
        Case eGDRanMessageType_AmendOrder
            MessageTypeToString = "eGDRanMessageType_AmendOrder"
        Case eGDRanMessageType_CancelOrder
            MessageTypeToString = "eGDRanMessageType_CancelOrder"
        Case eGDRanMessageType_UnloadApp
            MessageTypeToString = "eGDRanMessageType_UnloadApp"
        Case eGDRanMessageType_GetAccounts
            MessageTypeToString = "eGDRanMessageType_GetAccounts"
        Case eGDRanMessageType_GetContracts
            MessageTypeToString = "eGDRanMessageType_GetContracts"
        Case eGDRanMessageType_GetFills
            MessageTypeToString = "eGDRanMessageType_GetFills"
        Case eGDRanMessageType_GetOrders
            MessageTypeToString = "eGDRanMessageType_GetOrders"
        Case eGDRanMessageType_GetPositions
            MessageTypeToString = "eGDRanMessageType_GetPositions"
        Case eGDRanMessageType_GetSecurityList
            MessageTypeToString = "eGDRanMessageType_GetSecurityList"
        Case eGDRanMessageType_GetStatement
            MessageTypeToString = "eGDRanMessageType_GetStatement"
        Case eGDRanMessageType_GetSystemMessages
            MessageTypeToString = "eGDRanMessageType_GetSystemMessages"
        Case eGDRanMessageType_GetUserMessages
            MessageTypeToString = "eGDRanMessageType_GetUserMessages"
        Case eGDRanMessageType_GetAccountStatus
            MessageTypeToString = "eGDRanMessageType_GetAccountStatus"
        
        Case eGDRanMessageType_ConnectionInfo
            MessageTypeToString = "eGDRanMessageType_ConnectionInfo"
        Case eGDRanMessageType_AppLoaded
            MessageTypeToString = "eGDRanMessageType_AppLoaded"
        Case eGDRanMessageType_AppUnloaded
            MessageTypeToString = "eGDRanMessageType_AppUnloaded"
        Case eGDRanMessageType_Account
            MessageTypeToString = "eGDRanMessageType_Account"
        Case eGDRanMessageType_Contract
            MessageTypeToString = "eGDRanMessageType_Contract"
        Case eGDRanMessageType_Fill
            MessageTypeToString = "eGDRanMessageType_Fill"
        Case eGDRanMessageType_Order
            MessageTypeToString = "eGDRanMessageType_Order"
        Case eGDRanMessageType_Position
            MessageTypeToString = "eGDRanMessageType_Position"
        Case eGDRanMessageType_SecurityList
            MessageTypeToString = "eGDRanMessageType_SecurityList"
        Case eGDRanMessageType_Statement
            MessageTypeToString = "eGDRanMessageType_Statement"
        Case eGDRanMessageType_SystemMessage
            MessageTypeToString = "eGDRanMessageType_SystemMessage"
        Case eGDRanMessageType_UserMessage
            MessageTypeToString = "eGDRanMessageType_UserMessage"
        Case eGDRanMessageType_AccountStatus
            MessageTypeToString = "eGDRanMessageType_AccountStatus"
            
        Case Else
            MessageTypeToString = Str(nType)
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.MessageTypeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixOrderType
'' Description: Fix the order type if it is unknown (if possible)
'' Inputs:      Old Order string
'' Returns:     New Order string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FixOrderType(ByVal strOrderString As String) As String
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Order split up into an array
    Dim lPos As Long                    ' Position in an array
    Dim nOrderType As eTT_OrderType     ' Genesis Order type
    
    Set astrOrder = New cGdArray
    astrOrder.SplitFields strOrderString, vbTab

    nOrderType = GenesisOrderType(astrOrder(37))
    
    If nOrderType = -1& Then
        If m.astrOrderMap.BinarySearch(astrOrder(1) & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
            astrOrder(37) = Parse(m.astrOrderMap(lPos), vbTab, 2)
        End If
    Else
        If m.astrOrderMap.BinarySearch(astrOrder(1) & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
            m.astrOrderMap(lPos) = astrOrder(1) & vbTab & astrOrder(37)
        Else
            m.astrOrderMap.Add astrOrder(1) & vbTab & astrOrder(37), lPos
        End If
        m.astrOrderMap.Serialize AddSlash(m.strLogPath) & "Order.Map", True
    End If
    
    FixOrderType = astrOrder.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.FixOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CleanOutOrderMap
'' Description: Clean old orders out of the order map
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CleanOutOrderMap()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim strOrderID As String            ' Order ID in the order map
    Dim bFound As Boolean               ' Was the order found in the current orders?
    Dim bChanged As Boolean             ' Did the array change?
    
    bChanged = False
    For lIndex = m.astrOrderMap.Size - 1 To 0 Step -1
        strOrderID = Parse(m.astrOrderMap(lIndex), vbTab, 1)
        
        bFound = False
        For lIndex2 = 1 To m.BrokerInfo.Orders.Count
            If m.BrokerInfo.Orders(lIndex2).BrokerID = strOrderID Then
                bFound = True
                Exit For
            End If
        Next lIndex2
        
        If bFound = False Then
            bChanged = True
            m.astrOrderMap.Remove lIndex
        End If
    Next lIndex
    
    If bChanged Then
        m.astrOrderMap.Serialize AddSlash(m.strLogPath) & "Order.Map", True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.CleanOutOrderMap"
    
End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBrokerInfoOrder
'' Description: Add the order string to the broker info class
'' Inputs:      Order String, Refresh?
'' Returns:     None
''
'' RecordFrom:  Fill Index, Order ID, Exchange Order ID, User Name, Account,
''              Exchange Name, Base Symbol, Contact, B/S, Price, Price2,
''              Quantity, Linked, Qty Filled, Status, Date Exch Recd, Time Exch Recd,
''              XRef, Expiration, Original Order ID, Credit, Currency, Desc,
''              Exchange#, Order Type1, Order Type2, Salesman, Session, P/C,
''              Parked, Last Price, Broker Symbol, P/L, Date Filled, Time Filled,
''              Fill Price, Fill Flag, Order Type String
'' Order Rec:   Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date
'' Fill Rec:    Broker ID, Genesis ID, Fill ID, Account, Symbol, Fill Date,
''              B/S, Fill Quantity, Fill Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateBrokerInfoOrder(ByVal strOrderString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Order information split out from string
    Dim astrToAdd As cGdArray           ' Order information to add to broker info
    Dim strGenesisSymbol As String      ' Genesis symbol from the Rolfe & Nolan symbol
    Dim dMult As Double                 ' Multiplier
    Dim nOrderType As eTT_OrderType     ' Order type for the order
    
    Set astrOrder = New cGdArray
    astrOrder.Create eGDARRAY_Strings
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings
    
    astrOrder.SplitFields strOrderString, vbTab
    
    If Len(astrOrder(6)) = 0 Then
        astrToAdd(0) = astrOrder(1)
        astrToAdd(1) = astrOrder(17)
        astrToAdd(2) = astrOrder(19)
        astrToAdd(3) = astrOrder(4)
        astrToAdd(4) = Str(GenesisOrderStatus(astrOrder(14)))
        astrToAdd(5) = ""
        
        m.BrokerInfo.OrderStatusChange astrToAdd.JoinFields(vbTab), bRefresh
        
    ElseIf GenesisSymbol(astrOrder(6) & "-" & astrOrder(7) & "-" & astrOrder(28), strGenesisSymbol, dMult) Then
        nOrderType = GenesisOrderType(astrOrder(37))
        
        astrToAdd(0) = astrOrder(1)
        astrToAdd(1) = astrOrder(17)
        astrToAdd(2) = astrOrder(19)
        astrToAdd(3) = astrOrder(4)
        astrToAdd(4) = strGenesisSymbol
        astrToAdd(5) = Str(nOrderType)
        astrToAdd(6) = Str(CLng(UCase(astrOrder(8)) = "B"))
        astrToAdd(7) = astrOrder(11)
        Select Case nOrderType
            Case eTT_OrderType_Market, eTT_OrderType_MarketOnClose
                astrToAdd(8) = ""
                astrToAdd(9) = ""
            Case eTT_OrderType_Stop, eTT_OrderType_StopCloseOnly
                astrToAdd(8) = ""
                astrToAdd(9) = GenesisPrice(astrOrder(9), dMult)
            Case eTT_OrderType_Limit, eTT_OrderType_LimitCloseOnly
                astrToAdd(8) = GenesisPrice(astrOrder(9), dMult)
                astrToAdd(9) = ""
            Case eTT_OrderType_StopWithLimit, eTT_OrderType_StopWithLimitCloseOnly
                astrToAdd(8) = GenesisPrice(astrOrder(10), dMult)
                astrToAdd(9) = GenesisPrice(astrOrder(9), dMult)
        End Select
        astrToAdd(10) = astrOrder(18)
        astrToAdd(11) = Str(GenesisOrderStatus(astrOrder(14)))
        astrToAdd(12) = Str(CLng(Val(astrOrder(11))) - CLng(Val(astrOrder(13))))
        astrToAdd(13) = Str(DateFromString(astrOrder(15), astrOrder(16)))
        
        If Len(astrOrder(19)) > 0 Then
            m.BrokerInfo.RemoveOrder astrOrder(19)
        End If
            
        m.BrokerInfo.AddOrder astrToAdd.JoinFields(vbTab), bRefresh
        
        If astrOrder(36) = "F" Then
            astrToAdd.Clear
            astrToAdd(0) = astrOrder(1)
            astrToAdd(1) = astrOrder(17)
            astrToAdd(2) = astrOrder(1) & "-" & astrOrder(0)
            astrToAdd(3) = astrOrder(4)
            astrToAdd(4) = strGenesisSymbol
            astrToAdd(5) = Str(CalendarDate(DateFromString(astrOrder(33), astrOrder(34)), strGenesisSymbol))
            astrToAdd(6) = Str(CLng(UCase(astrOrder(8)) = "B"))
            astrToAdd(7) = astrOrder(13)
            astrToAdd(8) = GenesisPrice(astrOrder(35), dMult)
            
            m.BrokerInfo.AddFill astrToAdd.JoinFields(vbTab), bRefresh
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.UpdateBrokerInfoOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderToBrokerInfo
'' Description: Add the Genesis order string to the broker info class
'' Inputs:      Order String, Status
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
    RaiseError "cRan.GenesisOrderToBrokerInfo"
    
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
'' RecordTo:    Account, Broker ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelBrokerInfoOrder(ByVal strOrder As String)
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Order broken out into an array

    If Len(strOrder) > 0 Then
        Set astrOrder = New cGdArray
        astrOrder.Create eGDARRAY_Strings
    
        astrOrder.SplitFields strOrder, vbTab
        SendRanMessage eGDRanMessageType_CancelOrder, astrOrder(3) & vbTab & astrOrder(0)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.CancelBrokerInfoOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFillToBrokerInfo
'' Description: Add the fill string to the broker info class
'' Inputs:      Fill String, Refresh?
'' Returns:     None
''
'' RecordFrom:  Fill ID, Exchange, Base Symbol, Contract, B/S, Quantity,
''              Price, Order ID, Date Filled, Time Filled, Date Recd,
''              Time Recd, Exch Order ID, Order Status, Account, User Name,
''              XRef, P/C, Broker Symbol
'' RecordTo:    Broker ID, Genesis ID, Fill ID, Account, Symbol, Fill Date,
''              B/S, Fill Quantity, Fill Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddFillToBrokerInfo(ByVal strFillString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim astrFill As cGdArray            ' Fill information split from the string
    Dim astrToAdd As cGdArray           ' Array to add to the broker information
    Dim strGenesisSymbol As String      ' Genesis symbol from the Rolfe & Nolan symbol
    Dim dMult As Double                 ' Multiplier
    
    Set astrFill = New cGdArray
    astrFill.Create eGDARRAY_Strings
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings
    
    astrFill.SplitFields strFillString, vbTab
    
    If (Len(astrFill(5)) > 0) And (Len(astrFill(6)) > 0) Then
        If GenesisSymbol(astrFill(2) & "-" & astrFill(3) & "-" & astrFill(17), strGenesisSymbol, dMult) Then
            astrToAdd(0) = astrFill(7)
            astrToAdd(1) = astrFill(16)
            astrToAdd(2) = astrFill(0)
            astrToAdd(3) = astrFill(14)
            astrToAdd(4) = strGenesisSymbol
            astrToAdd(5) = CalendarDate(DateFromString(astrFill(8), astrFill(9)), strGenesisSymbol)
            astrToAdd(6) = Str(CLng(UCase(astrFill(4)) = "B"))
            astrToAdd(7) = astrFill(5)
            astrToAdd(8) = GenesisPrice(astrFill(6), dMult)
            
            m.BrokerInfo.AddFill astrToAdd.JoinFields(vbTab), bRefresh
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.AddFillToBrokerInfo"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddPositionToBrokerInfo
'' Description: Add the position string to the broker info class
'' Inputs:      Position String, Refresh?
'' Returns:     None
''
'' RecordFrom:  Account, Base Symbol, Contract, B/S, XRef, Currency, Desc,
''              Salesman, P/C, Last Price, Broker Symbol, P/L, Qty, Trade Price,
''              Trade Date, Account Name, Reg Code
'' RecordTo:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddPositionToBrokerInfo(ByVal strPositionString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim astrPosition As cGdArray        ' Array of position information split out from string
    Dim astrToAdd As cGdArray           ' Array of information to pass to broker info class
    Dim strGenesisSymbol As String      ' Genesis symbol from the Rolfe & Nolan symbol
    Dim dMult As Double                 ' Multiplier
    
    Set astrPosition = New cGdArray
    astrPosition.Create eGDARRAY_Strings
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings
    
    astrPosition.SplitFields strPositionString, vbTab
    
    If GenesisSymbol(astrPosition(1) & "-" & astrPosition(2) & "-" & astrPosition(8), strGenesisSymbol, dMult) Then
        astrToAdd(0) = astrPosition(0)
        astrToAdd(1) = strGenesisSymbol
        If UCase(astrPosition(3)) = "B" Then
            astrToAdd(2) = astrPosition(12)
        Else
            astrToAdd(2) = Str(CLng(Val(astrPosition(12))) * -1)
        End If
        astrToAdd(3) = Str(GenesisPrice(astrPosition(13), dMult))
        astrToAdd(4) = astrToAdd(2)
        
        m.BrokerInfo.AddPosition astrToAdd.JoinFields(vbTab), bRefresh
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.AddPositionToBrokerInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddPositionsToBrokerInfo
'' Description: Walk the positions array and send them all to the broker info
'' Inputs:      None
'' Returns:     None
''
'' RecordFrom:  Account, Symbol, Position, Carried Position, Carried Average Entry,
''              Carried Entry Prices
'' RecordTo:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddPositionsToBrokerInfo()
On Error GoTo ErrSection:

    Dim astrPosition As cGdArray        ' Array of position information split out from string
    Dim astrToAdd As cGdArray           ' Array of information to pass to broker info class
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrPosition = New cGdArray
    astrPosition.Create eGDARRAY_Strings
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings
    
    For lIndex = 0 To m.astrBrokerPos.Size - 1
        astrPosition.Clear
        astrToAdd.Clear
        
        astrPosition.SplitFields m.astrBrokerPos(lIndex), vbTab
        
        astrToAdd(0) = astrPosition(0)
        astrToAdd(1) = astrPosition(1)
        astrToAdd(2) = astrPosition(3)
        astrToAdd(3) = astrPosition(4)
        astrToAdd(4) = astrPosition(3)
        
        m.BrokerInfo.AddPosition astrToAdd.JoinFields(vbTab), True
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.AddPositionsToBrokerInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildPositions
'' Description: Build the positions from the position and fill summary info
'' Inputs:      None
'' Returns:     None
''
'' Position:    Account, Symbol, Current Position, Carried Position, Average
''              Entry, Overnight Position
'' FillSummary: Account, Symbol, At ID, Buys, Sells, Net, Total, PriceSum, Entries,
''              ClosedProfit, AvgEntry, Initial Fill Price, Initial Fill Date,
''              Session Date, Last Traded, Overnight
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
    Dim FillSumms As cAccountPositions  ' Fill summaries from broker info
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
                            
                            DumpDebug "Position Updated: " & m.astrBrokerPos(lPos)
                        Else
                            lPosition = 0&
                            strOvernight = "0"
                            lCurrentPos = .NumBuysSnapshot - .NumSellsSnapshot
                            m.astrBrokerPos.Add strAccount & vbTab & .Symbol & vbTab & Str(lCurrentPos) & vbTab & Str(lPosition) & vbTab & "" & vbTab & "", lPos
                        
                            DumpDebug "New Position Added: " & m.astrBrokerPos(lPos)
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
    RaiseError "cRan.BuildPositions"
    
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
    RaiseError "cRan.CreateFillAtPreviousClose"
    
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
    
    If FillSummary Is Nothing Then
        Set FillSummary = g.Broker.FillSummary(vAccountNumberOrID, vSymbolOrSymbolID, -1&)
    End If
    
    If Not FillSummary Is Nothing Then
        Set astrInfo = New cGdArray
        Set Bars = New cGdBars
        SetBarProperties Bars, FillSummary.SymbolOrSymbolID
        
        strAccount = g.Broker.GetAccountNumber(vAccountNumberOrID)
        If m.astrBrokerPos.BinarySearch(strAccount & vbTab & FillSummary.Symbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
            astrInfo.SplitFields Parse(m.astrBrokerPos(lPos), vbTab, 7), ","
        End If
        
        For lIndex = 0 To astrInfo.Size - 1
            Set astrPos = New cGdArray
            astrPos.SplitFields astrInfo(lIndex), ";"
        
            Set Fill = New cPtFill
            With Fill
                .AccountID = FillSummary.AccountID
                .AutoTradingItemID = 0&
                .SymbolOrSymbolID = FillSummary.SymbolOrSymbolID
                .Quantity = Abs(CLng(Val(astrPos(0))))
                .Price = Val(astrPos(1))
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
    RaiseError "cRan.CreateFillsForCarriedInfo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalendarDate
'' Description: Make sure to make the given date the calendar date
'' Inputs:      Date, Symbol, Bars
'' Returns:     Calendar Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CalendarDate(ByVal dDate As Double, ByVal vSymbolOrSymbolID As Variant, Optional ByVal Bars As cGdBars = Nothing) As Double
On Error GoTo ErrSection:

    Dim lSessionDate As Long            ' Session date for the date passed in
    Dim dReturn As Double               ' Return value for the function

    If Bars Is Nothing Then
        Set Bars = New cGdBars
        SetBarProperties Bars, vSymbolOrSymbolID
    End If
    
    dReturn = dDate
    lSessionDate = Bars.SessionDateForTradeTime(ConvertTimeZone(dDate, "CHI", Bars.Prop(eBARS_ExchangeTimeZoneInf)))
    If lSessionDate = Int(dDate) + 1 Then
        dReturn = dReturn - 1#
    End If
    
    CalendarDate = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRan.CalendarDate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when object is created
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    ConnectionStatus = eGDConnectionStatus_Disconnected
    m.bVerifyPositions = True
    m.dConnectTime = 0#
    
    Set m.astrToFix = New cGdArray
    m.astrToFix.Create eGDARRAY_Strings
    
    m.bWasConnected = False
    m.bSyncInProgress = False
    m.bFixCalledOnce = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.tblSymbols = Nothing
    Set m.aRanIdx = Nothing
    Set m.aGenesisIdx = Nothing
    Set m.astrAccounts = Nothing
    Set m.astrBrokerPos = Nothing
    Set m.astrToFix = Nothing
    
    Set m.astrAmendPending = Nothing
    Set m.BrokerInfo = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRan.Class_Terminate"
    
End Sub
