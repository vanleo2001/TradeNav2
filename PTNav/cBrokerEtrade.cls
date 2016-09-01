VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerEtrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerEtrade.cls
'' Description: Class to manange communications with E-Trade servers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Broker As cBroker                   ' Broker base class object
    BrokerKv As cBrokerKeyValue         ' Routines for key-value broker messages
    Symbols As cBrokerSymbols           ' Broker symbol conversion object
    
    strConsumerKey As String            ' Consumer key assigned to Trade Navigator by E-Trade
    strConsumerSecret As String         ' Consumer secret assigned to Trade Navigator by E-Trade
    strAuthorizeUrl As String           ' Authorize URL
    strAuthorizeUrlFormat As String     ' Authorize URL format
    lLastRefresh As Long                ' Date of the last refresh
    
    astrAccounts As cGdArray            ' Array of account numbers
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
        Case eTT_AccountType_Etrade
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "ETrade", True
            m.Broker.IniFile = AddSlash(App.Path) & "ETrade.INI"
            m.Broker.BrokerName = "E-Trade"
            m.Broker.ProcessName = "GenEtrade"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\ETrade\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenEtrade.EXE"
            m.Broker.ControlID = "ETrade"
            strSymbolFile = AddSlash(App.Path) & "Provided\ETradeToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\ETradeCnct.INI"
            
    End Select

    m.Broker.IsBroker = False

    'm.Broker.CopyProcess AddSlash(App.Path) & "..\Brokers\GenCqg.EXE"
    'FileCopy AddSlash(App.Path) & "..\Brokers\FixDefCqg.xml", AddSlash(strBrokerPath) & "FixDef.xml", True
    
    Set m.Symbols = New cBrokerSymbols
    m.Symbols.FromFile strSymbolFile
    
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Connect) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Disconnect) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccounts) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccountsMode) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ConnectInfo) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ChangePassword) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Refresh) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewActivity) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_BrokerView) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewOnline) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_VerifyPositions) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_AccountDetails) = False
    
    m.strConsumerKey = DecryptFromHex(GetIniFileProperty("Key1", "", "Vendor", m.Broker.ConnectIni))
    m.strConsumerSecret = DecryptFromHex(GetIniFileProperty("Key2", "", "Vendor", m.Broker.ConnectIni))
    m.strAuthorizeUrl = GetIniFileProperty("AuthorizeUrl", "", "URL", m.Broker.ConnectIni)
    m.strAuthorizeUrlFormat = GetIniFileProperty("AuthorizeUrlFormat", "", "URL", m.Broker.ConnectIni)
    m.lLastRefresh = GetIniFileProperty("Last", Int(CurrentTime) - 730, "Refresh", m.Broker.IniFile)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerEtrade.Init"

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
    RaiseError "cBrokerEtrade.ShowConfig"
    
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

    ShowBrokerConnectionInfo = False

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerEtrade.ShowBrokerConnectionInfo"
    
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
    Dim brokerMsg As cBrokerMessage     ' Broker Message object
        
    bReturn = frmLoginEtrade.ShowMe(Me)
    If bReturn = True Then
        Set brokerMsg = New cBrokerMessage
        brokerMsg.Add "OAuthVerifier", frmLoginEtrade.Verifier
        
        m.Broker.SendBrokerMessage eGDBrokerMessageType_Connect, brokerMsg.ToString(False)
    End If

    ShowLogin = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerEtrade.ShowLogin"
    
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

    GetUrl

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerEtrade.SendConnect"
    
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
    RaiseError "cBrokerEtrade.SendDisconnect"
    
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
    RaiseError "cBrokerEtrade.UnloadApp"
    
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
    Set BrokerSym = New cBrokerSymbol
    
    Select Case UCase(strSecType)
        Case "S"
            If InStr(strGenesisSymbol, "_") = 0 Then
                strReturn = strGenesisSymbol
            ElseIf Right(strGenesisSymbol, 2) = "_P" Then
                strReturn = Replace(strGenesisSymbol, "_P", ".PR")
            ElseIf InStr(strGenesisSymbol, "_P") <> 0 Then
                strReturn = Replace(strGenesisSymbol, "_P", ".PR.")
            ElseIf Right(strGenesisSymbol, 2) = "_W" Then
                strReturn = Replace(strGenesisSymbol, "_P", ".WS")
            ElseIf InStr(strGenesisSymbol, "_W") <> 0 Then
                strReturn = Replace(strGenesisSymbol, "_P", ".WS.")
            Else
                strReturn = Replace(strGenesisSymbol, "_", ".")
            End If
            
            BrokerSym.BrokerBase = strReturn
            BrokerSym.BrokerPositionBase = BrokerSym.BrokerBase
            BrokerSym.GenesisBase = strGenesisSymbol
        
        Case "SO"
            Set OptionSym = New cOptionSymbol
            OptionSym.FromGenesis strGenesisSymbol
            strReturn = OptionSym.ToEtrade(OptionSym.BaseSymbol)
        
            BrokerSym.BrokerBase = OptionSym.BaseSymbol
            BrokerSym.BrokerPositionBase = BrokerSym.BrokerBase
            BrokerSym.GenesisBase = OptionSym.BaseSymbol
        
    End Select
    
    BrokerSym.BrokerExchange = ""
    BrokerSym.OrderTypeMask = "11110000"
    BrokerSym.PriceFormat = ""
    BrokerSym.PriceMult = 1
    BrokerSym.RtOverride = True
    BrokerSym.StrikeMult = 1
    BrokerSym.StrikePriceFormat = ""
    BrokerSym.TifMask = "111"
    
    BrokerSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerEtrade.BrokerSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Attempt to convert the broker symbol to the Genesis symbology
'' Inputs:      Broker Symbol
'' Returns:     Genesis Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbol(ByVal strBrokerSymbol As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim OptionSym As cOptionSymbol      ' Option symbol object
    
    strReturn = ""
    If InStr(strBrokerSymbol, "|") = 0 Then
        If InStr(strBrokerSymbol, ".PR.") <> 0 Then
            strReturn = Replace(strBrokerSymbol, ".PR.", "_P")
        ElseIf InStr(strBrokerSymbol, ".PR") <> 0 Then
            strReturn = Replace(strBrokerSymbol, ".PR", "_P")
        ElseIf InStr(strBrokerSymbol, ".WS.") <> 0 Then
            strReturn = Replace(strBrokerSymbol, ".WS.", "_W")
        ElseIf InStr(strBrokerSymbol, ".WS") <> 0 Then
            strReturn = Replace(strBrokerSymbol, ".WS", "_W")
        ElseIf InStr(strBrokerSymbol, ".") <> 0 Then
            strReturn = Replace(strBrokerSymbol, ".", "_")
        Else
            strReturn = strBrokerSymbol
        End If
    Else
        Set OptionSym = New cOptionSymbol
        OptionSym.FromEtrade strBrokerSymbol
        OptionSym.ConvertBrokerInfo OptionSym.BrokerBase, 1#
        strReturn = OptionSym.ToGenesis
    End If
    
    GenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerEtrade.GenesisSymbol"
    
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
    RaiseError "cBrokerEtrade.BrokerOrderType"
    
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
    RaiseError "cBrokerEtrade.AddOrder"
    
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
    
    bReturn = False
    
    AmendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerEtrade.AmendOrder"
    
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
    
    bReturn = False
    
    CancelOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerEtrade.CancelOrder"
    
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
    
    bReturn = False
    
    ParkOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerEtrade.ParkOrder"
    
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
    RaiseError "cBrokerEtrade.Refresh"
    
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
    RaiseError "cBrokerEtrade.GetAccounts"
    
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
    RaiseError "cBrokerEtrade.GetOrders"
    
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
    
    brokerMsg.Add "StartDate", Str(m.lLastRefresh)
    
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetFills, brokerMsg.ToString
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerEtrade.GetFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTransactions
'' Description: Request transactions from the server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetTransactions(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    
    Set brokerMsg = New cBrokerMessage
    If Len(strAccount) > 0 Then
        brokerMsg.Add "Account", strAccount
    Else
        brokerMsg.Add "Accounts", m.astrAccounts.JoinFields(",")
    End If
    
    brokerMsg.Add "StartDate", Str(m.lLastRefresh)
    
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetTransactions, brokerMsg.ToString
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerEtrade.GetTransactions"
    
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
    RaiseError "cBrokerEtrade.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetUrl
'' Description: Request the URL to get the authorization code
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetUrl()
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    
    Set brokerMsg = New cBrokerMessage
    brokerMsg.Add "ConsumerKey", m.strConsumerKey
    brokerMsg.Add "ConsumerSecret", m.strConsumerSecret
    brokerMsg.Add "OAuthAuthorizeUrl", m.strAuthorizeUrl
    brokerMsg.Add "AuthorizeUrlFormat", m.strAuthorizeUrlFormat

    m.Broker.SendBrokerMessage eGDBrokerMessageType_ConsumerInfo, brokerMsg.ToString(False), , brokerMsg.ToString(True)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerEtrade.GetUrl"
    
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
    RaiseError "cBrokerEtrade.NextGenesisID"
    
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
    RaiseError "cBrokerEtrade.StartingGenesisIdForOptNav"
    
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
    RaiseError "cBrokerEtrade.IsTradeableSymbol"
    
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
    RaiseError "cBrokerEtrade.IsEnabledSymbol"
    
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
    RaiseError "cBrokerEtrade.OrderTypeAllowed"
    
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
    RaiseError "cBrokerEtrade.TimeInForceAllowed"
    
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
            
        Case eGDBrokerMessageType_Transactions
            TransactionReceived strMessage
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerEtrade.HandleMessage"
    
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
    RaiseError "cBrokerEtrade.DateToDouble"
    
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
    RaiseError "cBrokerEtrade.StatusFromInfo"
    
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
            Acct.Name = brokerMsg("AccountName")
            
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
    RaiseError "cBrokerEtrade.AccountFromInfo"
    
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

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    Dim strFirstField As String         ' First field in the string
    
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerEtrade.OrderFromInfo"
    
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

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    Dim strFirstField As String         ' First field in the string
    
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerEtrade.FillFromInfo"
    
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

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    Dim strFirstField As String         ' First field in the string
    Dim strSecondField As String        ' Second field in the string
    
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        strSecondField = Parse(strMessage, vbTab, 2)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
            GetTransactions strSecondField
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerEtrade.PositionFromInfo"
    
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
    RaiseError "cBrokerEtrade.GenesisSymbolList"
    
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
    RaiseError "cBrokerEtrade.TimeZone"
    
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
    RaiseError "cBrokerEtrade.JustConnected"
    
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
    RaiseError "cBrokerEtrade.SymbolInformation"
    
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
    RaiseError "cBrokerEtrade.BrokerPrice"
    
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
    RaiseError "cBrokerEtrade.GenesisPrice"
    
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
    Dim strBrokerSymbol As String       ' Broker symbol
    
    strBrokerSymbol = BrokerSymbol(Order.Symbol, BrokerSym)
    BrokerOrderString = m.BrokerKv.OrderToMessage(Order, bNewOrder, strBrokerSymbol, BrokerSym, m.Broker, strPreviousGenesisID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerEtrade.BrokerOrderString"
    
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
    
    If FormIsLoaded("frmLoginEtrade") Then
        frmLoginEtrade.Broker_Url Replace(brokerMsg("VerificationUrl"), "~", "=")
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerEtrade.HandleLoginUrl"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactionReceived
'' Description: Handle the transaction message back from the stand-alone
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactionReceived(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    Dim strFirstField As String         ' First field in the string
    
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
            m.lLastRefresh = Int(CurrentTime)
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerEtrade.TransactionReceived"
    
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

    Set m.Broker = New cBroker
    Set m.BrokerKv = New cBrokerKeyValue
    
    Set m.astrAccounts = New cGdArray
    m.astrAccounts.Create eGDARRAY_Strings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerEtrade.Class_Initialize"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerEtrade.Class_Terminate"
    
End Sub

