VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerTradier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerTradier.cls
'' Description: Class to manange communications with Tradier servers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 09/14/2015   DAJ         Created
'' 10/06/2015   DAJ         Grab AverageEntry and Clearing date on positions; strike price
''                          conversion fix; fake a position record when flat, but have fills
'' 12/02/2015   DAJ         Fix for Invalid Procedure call error in PositionMessageForGenesisSymbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Broker As cBroker                   ' Broker base class object
    BrokerKv As cBrokerKeyValue         ' Routines for key-value broker messages
    Symbols As cBrokerSymbols           ' Broker symbol conversion object
    
    strLoginUrl As String               ' Login URL
    strConsumerKey As String            ' Consumer key
    strConsumerSecret As String         ' Consumer secret
    strApiUrl As String                 ' API URL
    strState As String                  ' Unique state string
    strCode As String                   ' Code back from the website
    dLoginExpire As Double              ' Time the login is set to expire
    
    astrAccounts As cGdArray            ' Array of account numbers
    astrFillSymbols As cGdArray         ' Array of symbols for which we got fills
End Type
Private m As mPrivate

Public Property Get Broker() As cBroker
    Set Broker = m.Broker
End Property

Private Property Get ConnectString() As String
    ConnectString = ""
End Property

Private Property Get ConnectStringLog() As String
    ConnectStringLog = ""
End Property

Public Property Get EnabledSymbols() As cGdArray
    Set EnabledSymbols = GenesisSymbolList
End Property

Public Property Get State() As String
    If Len(m.strState) = 0 Then
        m.strState = StripStr(UCase(RI_GetMachineID), "- ")
    End If
    
    State = m.strState
End Property

Public Property Get LoginExpire() As Double
    LoginExpire = m.dLoginExpire
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
    Dim strBrokerPath As String         ' Broker path
    
    Select Case nBroker
        Case eTT_AccountType_Tradier
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "Tradier", True
            m.Broker.IniFile = AddSlash(App.Path) & "Tradier.INI"
            m.Broker.BrokerName = "Tradier"
            m.Broker.ProcessName = "GenTradier"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\Tradier\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenTradier.EXE"
            m.Broker.ControlID = "Tradier"
            strSymbolFile = AddSlash(App.Path) & "Provided\TradierToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\TradierCnct.INI"
            
    End Select

    m.Broker.IsBroker = False

    Set m.Symbols = New cBrokerSymbols
    m.Symbols.FromFile strSymbolFile
    
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Connect) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Disconnect) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccounts) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccountsMode) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ConnectInfo) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ChangePassword) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Refresh) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewActivity) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_BrokerView) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewOnline) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_VerifyPositions) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_AccountDetails) = False
    
    m.strApiUrl = GetIniFileProperty("Api", "", "URL", m.Broker.ConnectIni)
    m.strConsumerKey = DecryptFromHex(GetIniFileProperty("Key", "", "Info", m.Broker.ConnectIni))
    m.strConsumerSecret = DecryptFromHex(GetIniFileProperty("Secret", "", "Info", m.Broker.ConnectIni))
    m.strLoginUrl = GetIniFileProperty("Login", "", "URL", m.Broker.ConnectIni)
    If Len(m.strLoginUrl) > 0 Then
        m.strLoginUrl = m.strLoginUrl & "&state=" & State
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.Init"

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

    ShowConfig = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.ShowConfig"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowBrokerConnectionInfo
'' Description: Allow the user to view their broker connection information
'' Inputs:      New?, User Name
'' Returns:     True if dialog OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowBrokerConnectionInfo(Optional ByVal bNew As Boolean = False, Optional ByVal strUserName As String = "") As Boolean
On Error GoTo ErrSection:

    ShowBrokerConnectionInfo = ShowLogin

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.ShowBrokerConnectionInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowLogin
'' Description: Show the user the login form
'' Inputs:      New User Name, Are we switching?
'' Returns:     True if connect, False othwerwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowLogin(Optional ByVal strNewUserName As String = "", Optional ByVal bSwitching As Boolean = False)
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
        
    bReturn = frmLoginTradier.ShowMe(Me, m.strLoginUrl, State)
    If bReturn = True Then
        m.strCode = frmLoginTradier.Code
    End If

    ShowLogin = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.ShowLogin"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendConnect
'' Description: Send a connect to the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendConnect()
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker Message object

    Set brokerMsg = New cBrokerMessage
    brokerMsg.Add "ApiUrl", m.strApiUrl
    brokerMsg.Add "Code", m.strCode
    brokerMsg.Add "ConsumerKey", m.strConsumerKey
    brokerMsg.Add "ConsumerSecret", m.strConsumerSecret
    
    m.Broker.SendBrokerMessage eGDBrokerMessageType_Connect, brokerMsg.ToString(False), , brokerMsg.ToString(True)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.SendConnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendDisconnect
'' Description: Send a disconnect to the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendDisconnect()
On Error GoTo ErrSection:

    m.Broker.SendBrokerMessage eGDBrokerMessageType_Disconnect, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.SendDisconnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnloadApp
'' Description: Unload the stand-alone application
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UnloadApp()
On Error GoTo ErrSection:

    m.Broker.SendBrokerMessage eGDBrokerMessageType_UnloadApp, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.UnloadApp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerSymbol
'' Description: Attempt to convert the Genesis symbol to the broker symbology
'' Inputs:      Genesis Symbol
'' Returns:     Broker Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerSymbol(ByVal strGenesisSymbol As String, Optional BrokerSym As cBrokerSymbol) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strSecType As String            ' Security type of the symbol passed in
    Dim OptionSym As cOptionSymbol      ' Option symbol object
    
    strReturn = ""
    strSecType = SecurityType(strGenesisSymbol, True)
    
    Select Case UCase(strSecType)
        Case "S"
            strReturn = strGenesisSymbol
            Set BrokerSym = GetBrokerSym(strReturn, strGenesisSymbol)
        
        Case "SO"
            Set OptionSym = New cOptionSymbol
            OptionSym.FromGenesis strGenesisSymbol
            strReturn = OptionSym.ToTradier(OptionSym.BaseSymbol, 0.001)
            Set BrokerSym = GetBrokerSym(OptionSym.BaseSymbol, OptionSym.BaseSymbol)
        
    End Select
    
    BrokerSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.BrokerSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Attempt to convert the broker symbol to the Genesis symbology
'' Inputs:      Broker Symbol, Broker Sym object
'' Returns:     Genesis Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbol(ByVal strBrokerSymbol As String, Optional BrokerSym As cBrokerSymbol) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim OptionSym As cOptionSymbol      ' Option symbol object
    
    strReturn = ""
    
    If Len(strBrokerSymbol) < 10 Then
        strReturn = strBrokerSymbol
        Set BrokerSym = GetBrokerSym(strBrokerSymbol, strBrokerSymbol)
        
    Else
        Set OptionSym = New cOptionSymbol
        OptionSym.FromTradier strBrokerSymbol
        OptionSym.ConvertBrokerInfo OptionSym.BrokerBase, 0.001
        strReturn = OptionSym.ToGenesis
        
        Set BrokerSym = GetBrokerSym(OptionSym.BrokerBase, OptionSym.BrokerBase)
    End If
    
    GenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.GenesisSymbol"
    
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

    BrokerOrderType = m.BrokerKv.BrokerOrderType(nGenesisOrderType)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.BrokerOrderType"
    
End Function

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
    Dim strOrderString As String        ' Broker order string
    
    bReturn = False
    strOrderString = BrokerOrderString(Order, True)
    If Len(strOrderString) > 0 Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AddOrder, strOrderString
        bReturn = True
    End If
    
    AddOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.AddOrder"
    
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

    Dim bReturn As Boolean              ' Return value for the function
    Dim strOrderString As String        ' Broker order string
    
    bReturn = False
    strOrderString = BrokerOrderString(Order, False)
    If Len(strOrderString) > 0 Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AmendOrder, strOrderString
        bReturn = True
    End If
    
    AmendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.AmendOrder"
    
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

    Dim bReturn As Boolean              ' Return value for the function
    Dim strOrderString As String        ' Broker order string
    
    bReturn = False
    strOrderString = BrokerOrderString(Order, False)
    If Len(strOrderString) > 0 Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_CancelOrder, strOrderString
        bReturn = True
    End If
    
    CancelOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.CancelOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Attempt to park the order
'' Inputs:      Order to park
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ParkOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strOrderString As String        ' Broker order string
    
    bReturn = False
    strOrderString = BrokerOrderString(Order, False)
    If Len(strOrderString) > 0 Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_CancelOrder, strOrderString
        bReturn = True
    End If
    
    ParkOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.ParkOrder"
    
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

    GetAccounts

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.Refresh"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Request accounts from the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAccounts()
On Error GoTo ErrSection:

    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetAccounts, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.GetAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Request orders from the server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetOrders(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    
    Set brokerMsg = New cBrokerMessage
    If Len(strAccount) > 0 Then
        brokerMsg.Add "Account", strAccount
    Else
        brokerMsg.Add "Accounts", m.astrAccounts.JoinFields(",")
    End If
    
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetOrders, brokerMsg.ToString
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.GetOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFills
'' Description: Request fills from the server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetFills(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    
    Set brokerMsg = New cBrokerMessage
    If Len(strAccount) > 0 Then
        brokerMsg.Add "Account", strAccount
    Else
        brokerMsg.Add "Accounts", m.astrAccounts.JoinFields(",")
    End If
    
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetFills, brokerMsg.ToString
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.GetFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositions
'' Description: Request positions from the server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPositions(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    
    Set brokerMsg = New cBrokerMessage
    If Len(strAccount) > 0 Then
        brokerMsg.Add "Account", strAccount
    Else
        brokerMsg.Add "Accounts", m.astrAccounts.JoinFields(",")
    End If
    
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetPositions, brokerMsg.ToString
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.GetPositions"
    
End Sub

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
    Dim strIniFile As String            ' Ini File
    
    strIniFile = m.Broker.IniFile
    lDate = GetIniFileProperty("Date", 0&, "ID", strIniFile)
    lCounter = GetIniFileProperty("Counter", 0&, "ID", strIniFile)
    lCustomerID = (RI_GetLastDataServiceID \ 1000) Mod 1000000
    
    If lDate <> Date Then
        lDate = Date
        lCounter = 0&
    Else
        lCounter = lCounter + 1
    End If

    SetIniFileProperty "Date", lDate, "ID", strIniFile
    SetIniFileProperty "Counter", lCounter, "ID", strIniFile
    
    NextGenesisID = Format(lCustomerID, "000000") & Format(lDate, "00000") & Format(lCounter, "00000")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.NextGenesisID"
    
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
    
    StartingGenesisIdForOptNav = Format(lCustomerID, "000000") & Format(Date, "00000") & "90000"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.StartingGenesisIdForOptNav"
    
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

    IsTradeableSymbol = (Len(BrokerSymbol(strGenesisSymbol)) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.IsTradeableSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsEnabledSymbol
'' Description: Is the given symbol enabled for trading for the user?
'' Inputs:      Symbol
'' Returns:     True if enabled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEnabledSymbol(ByVal strGenesisSymbol As String, Optional strBrokerBase As String, Optional strBrokerExchange As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strSecurityType As String       ' Security type for the symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = (Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0)
    If bReturn Then
        strBrokerBase = BrokerSym.BrokerBase
        strBrokerExchange = BrokerSym.BrokerExchange
        
        strSecurityType = SecurityType(strGenesisSymbol, True)
        
        Select Case UCase(strSecurityType)
            Case "S"
                bReturn = HasModule("E_US")
            
            Case "SO"
                bReturn = HasModule("E_OPRA")
            
        End Select
    End If
    
    IsEnabledSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.IsEnabledSymbol"
    
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
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = False
    If Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0 Then
        bReturn = BrokerSym.OrderTypeAllowed(nOrderType)
    End If
    
    OrderTypeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.OrderTypeAllowed"
    
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
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = False
    If Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0 Then
        bReturn = BrokerSym.TifAllowed(nTimeInForce)
    End If
    
    TimeInForceAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.TimeInForceAllowed"
    
End Function

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
    
    Select Case nType
        Case eGDBrokerMessageType_LoginUrl
            HandleLoginUrl strMessage
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.HandleMessage"
    
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

    Dim dReturn As Double               ' Return value of the function
    
    dReturn = 0#
    
    DateToDouble = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.DateToDouble"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StatusFromInfo
'' Description: Get connection status information from the given message
'' Inputs:      Message, Status, Error, User Name, Allow Reconnect?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StatusFromInfo(ByVal strMessage As String, nStatus As eGDConnectionStatus, strError As String, strUserName As String, bAllowReconnect As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    
    Set brokerMsg = New cBrokerMessage
    brokerMsg.FromString strMessage
    
    nStatus = CLng(Val(brokerMsg("Status")))
    strError = brokerMsg("Error")
    strUserName = ""
    
    If Len(brokerMsg("AllowReconnect")) = 0 Then
        bAllowReconnect = True
    Else
        bAllowReconnect = (CLng(Val(brokerMsg("AllowReconnect"))) <> 0)
    End If
    
    StatusFromInfo = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.StatusFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountFromInfo
'' Description: Fill an account object from the given message
'' Inputs:      Message, Account
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountFromInfo(ByVal strMessage As String, Acct As cPtAccount) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    Dim strFirstField As String         ' First field in the string
    Dim lPos As Long                    ' Position in an array

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
            m.astrAccounts.Clear
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
            
            Set Acct = m.BrokerKv.AccountFromMessage(brokerMsg, m.Broker, "", "")
            Acct.Name = brokerMsg("Account")
            
            If m.astrAccounts.BinarySearch(Acct.AccountNumber, lPos) = False Then
                m.astrAccounts.Add Acct.AccountNumber, lPos
            End If
            
            bReturn = True
        End If
    End If
    
    AccountFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.AccountFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromInfo
'' Description: Fill an order object from the given message
'' Inputs:      Message, Order
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderFromInfo(ByVal strMessage As String, Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    Dim strFirstField As String         ' First field in the string
    Dim astrGenesisSymbols As cGdArray  ' Genesis symbols
    Dim lNumLegs As Long                ' Number of legs
    Dim lIndex As Long                  ' Index into a for loop
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            bReturn = True
            
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
            
            lNumLegs = CLng(Val(brokerMsg("NumLegs")))
            Set astrGenesisSymbols = New cGdArray
            astrGenesisSymbols.Create eGDARRAY_Strings, lNumLegs
            
            astrGenesisSymbols(0) = GenesisSymbol(brokerMsg("Symbol1"), BrokerSym)
            If Len(astrGenesisSymbols(0)) > 0 Then
                For lIndex = 2 To lNumLegs
                    astrGenesisSymbols(lIndex - 1) = GenesisSymbol(brokerMsg("Symbol" & Str(lIndex)))
                    If Len(astrGenesisSymbols(lIndex - 1)) = 0 Then
                        bReturn = False
                    End If
                Next lIndex
            Else
                bReturn = False
            End If
            
            Set Order = m.BrokerKv.MultiLegOrderFromMessage(brokerMsg, m.Broker, astrGenesisSymbols, BrokerSym)
        End If
    End If
    
    OrderFromInfo = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.OrderFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillFromInfo
'' Description: Fill a fill object from the given message
'' Inputs:      Message, Fill, Genesis Order ID
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillFromInfo(ByVal strMessage As String, Fill As cPtFill, Optional strGenesisOrderID As String = "") As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    Dim strFirstField As String         ' First field in the string
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim strKey As String                ' Key into the fill symbols array
    Dim lPos As Long                    ' Position of the symbol in the array
    
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
            m.astrFillSymbols.Clear
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
            
            strGenesisSymbol = GenesisSymbol(brokerMsg("Symbol"), BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                Set Fill = m.BrokerKv.FillFromMessage(brokerMsg, m.Broker, strGenesisSymbol, BrokerSym)
            
                strKey = brokerMsg("Account") & vbTab & strGenesisSymbol
                If m.astrFillSymbols.BinarySearch(strKey, lPos) = False Then
                    m.astrFillSymbols.Add strKey, lPos
                    m.Broker.DumpDebug vbTab & vbTab & "Added to Fill Symbols: '" & strKey & "'"
                End If
            End If
            
            bReturn = (Len(strGenesisSymbol) > 0)
        End If
    End If
    
    FillFromInfo = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.FillFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionFromInfo
'' Description: Fill a position object from the given message
'' Inputs:      Message, Position
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionFromInfo(ByVal strMessage As String, Position As cPtPosition) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    Dim strFirstField As String         ' First field in the string
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim lPos As Long                    ' Position of the symbol in the array
    Dim lIndex As Long                  ' Index into a for loop
    Dim strKey As String                ' Key into the fill symbols array
    Dim strPositionMessage As String    ' Position message
    
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
            For lIndex = m.astrFillSymbols.Size - 1 To 0 Step -1
                strPositionMessage = PositionMessageForGenesisSymbol(Parse(m.astrFillSymbols(lIndex), vbTab, 2), Parse(m.astrFillSymbols(lIndex), vbTab, 1))
                If Len(strPositionMessage) > 0 Then
                    m.Broker.PositionReceived strPositionMessage, True
                End If
            Next lIndex
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
            
            strGenesisSymbol = GenesisSymbol(brokerMsg("Symbol"), BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                Set Position = m.BrokerKv.PositionFromMessage(brokerMsg, m.Broker, strGenesisSymbol)
                
                Position.ClearingDate = m.BrokerKv.DateToDouble(brokerMsg("ClearingBusinessDate"))
                Position.AverageEntry = m.Broker.GenesisPrice(brokerMsg("AvgPx"), BrokerSym.PriceMult)
            
                strKey = brokerMsg("Account") & vbTab & strGenesisSymbol
                If m.astrFillSymbols.BinarySearch(strKey, lPos) = True Then
                    m.astrFillSymbols.Remove lPos
                    m.Broker.DumpDebug vbTab & vbTab & "Removed from Fill Symbols: '" & strKey & "'"
                End If
            End If
            
            bReturn = (Len(strGenesisSymbol) > 0)
        End If
    End If
    
    PositionFromInfo = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.PositionFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbolList
'' Description: Return a list of the Genesis symbols that are in the file
'' Inputs:      Only if stream?
'' Returns:     List of symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbolList(Optional ByVal bOnlyIfStream As Boolean = True) As cGdArray
On Error GoTo ErrSection:

    Set GenesisSymbolList = m.Symbols.GenesisSymbolList

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.GenesisSymbolList"
    
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

    TimeZone = "GMT"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.TimeZone"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    JustConnected
'' Description: Handle the fact that we have just gone connected with the broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub JustConnected()
On Error GoTo ErrSection:

    GetAccounts

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.JustConnected"
    
End Sub

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
    
    strReturn = ""
    
    SymbolInformation = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.SymbolInformation"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBrokerSym
'' Description: Get a broker symbol object based on the security type
'' Inputs:      Broker Base Symbol, Genesis Base Symbol
'' Returns:     Broker Symbol object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetBrokerSym(ByVal strBrokerBase As String, ByVal strGenesisBase As String) As cBrokerSymbol
On Error GoTo ErrSection:

    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object to return from the function
    
    Set BrokerSym = New cBrokerSymbol
            
    BrokerSym.BrokerBase = strBrokerBase
    BrokerSym.BrokerPositionBase = strBrokerBase
    BrokerSym.GenesisBase = strGenesisBase
    BrokerSym.BrokerExchange = ""
    BrokerSym.OrderTypeMask = "11110000"
    BrokerSym.PriceFormat = "0.####"
    BrokerSym.PriceMult = 1
    BrokerSym.RtOverride = True
    BrokerSym.StrikeMult = 1
    BrokerSym.StrikePriceFormat = ""
    BrokerSym.TifMask = "111"
    
    Set GetBrokerSym = BrokerSym

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.GetBrokerSym"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerPrice
'' Description: Convert a Genesis price to the broker price
'' Inputs:      Genesis Price, Broker Symbol Object
'' Returns:     Broker Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerPrice(ByVal dGenesisPrice As Double, ByVal BrokerSym As cBrokerSymbol) As Double
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = Str(dGenesisPrice)
    If Not BrokerSym Is Nothing Then
        strReturn = m.Broker.BrokerPrice(dGenesisPrice, BrokerSym.PriceMult, BrokerSym.PriceFormat)
    End If
    
    BrokerPrice = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.BrokerPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisPrice
'' Description: Attempt to convert the broker price to a Genesis price
'' Inputs:      Broker Price, Broker Symbol Object
'' Returns:     Genesis Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisPrice(ByVal strBrokerPrice As String, ByVal BrokerSym As cBrokerSymbol) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    
    dReturn = 0#
    If Not BrokerSym Is Nothing Then
        dReturn = m.Broker.GenesisPrice(strBrokerPrice, BrokerSym.PriceMult)
    End If
    
    GenesisPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.GenesisPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerOrderString
'' Description: Build an order string to send to the broker
'' Inputs:      Order, New Order?, Previous Genesis ID
'' Returns:     Broker Order String
''
'' Fields:      GenesisID, BrokerID, Account, Symbol, Exchange, Order Type,
''              Buy/Sell, Quantity, Stop, Limit, TIF, Expiration
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerOrderString(ByVal Order As cPtOrder, ByVal bNewOrder As Boolean, Optional ByVal strPreviousGenesisID As String = "") As String
On Error GoTo ErrSection:

    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim astrBrokerSymbols As cGdArray   ' Array of broker symbols for the legs
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrBrokerSymbols = New cGdArray
    astrBrokerSymbols.Create eGDARRAY_Strings, Order.NumberOfLegs
    
    astrBrokerSymbols(0) = BrokerSymbol(Order.OrderLegs(1).Symbol, BrokerSym)
    For lIndex = 2 To Order.NumberOfLegs
        astrBrokerSymbols(lIndex - 1) = BrokerSymbol(Order.OrderLegs(lIndex).Symbol)
    Next lIndex
    
    BrokerOrderString = m.BrokerKv.MultiLegOrderToMessage(Order, bNewOrder, astrBrokerSymbols, BrokerSym, m.Broker, strPreviousGenesisID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.BrokerOrderString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleLoginUrl
'' Description: Handle the login URL message back from the stand-alone
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleLoginUrl(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    
    Set brokerMsg = New cBrokerMessage
    brokerMsg.FromString strMessage
    
    m.dLoginExpire = Val(brokerMsg("ExpireTime"))
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.HandleLoginUrl"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionMessageForGenesisSymbol
'' Description: Build a position message for the given Genesis symbol
'' Inputs:      Genesis Symbol, Account
'' Returns:     Position Message
''
'' Fields:      Account, IB Base Symbol, Symbol, Contract, Strike, P/C, Exchange,
''              Security Type, Currency, Position, Realized, Unrealized,
''              Avg Entry
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PositionMessageForGenesisSymbol(ByVal strGenesisSymbol As String, ByVal strAccount As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim strBrokerSymbol As String       ' Broker symbol
    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    
    strReturn = ""
    strBrokerSymbol = BrokerSymbol(strGenesisSymbol, BrokerSym)
    If Len(strBrokerSymbol) > 0 Then
        Set brokerMsg = New cBrokerMessage
        
        brokerMsg.Add "Account", strAccount
        brokerMsg.Add "Current", "0"
        brokerMsg.Add "Symbol", strBrokerSymbol
        
        strReturn = brokerMsg.ToString
    End If
    
    PositionMessageForGenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTradier.PositionMessageForGenesisSymbol"
    
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

    Set m.Broker = New cBroker
    Set m.BrokerKv = New cBrokerKeyValue
    
    Set m.astrAccounts = New cGdArray
    m.astrAccounts.Create eGDARRAY_Strings
    Set m.astrFillSymbols = New cGdArray
    m.astrFillSymbols.Create eGDARRAY_Strings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.Class_Initialize"
    
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

    Set m.Broker = Nothing
    Set m.BrokerKv = Nothing
    Set m.astrAccounts = Nothing
    Set m.astrFillSymbols = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTradier.Class_Terminate"
    
End Sub
