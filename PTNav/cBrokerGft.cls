VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerGft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerGft.cls
'' Description: Class to manage communications with GFT servers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/30/2012   DAJ         Allow Reconnect flag on Connection Status message
'' 08/28/2012   DAJ         CarriedFillFromInfo, First Field Begin/End
'' 12/11/2012   DAJ         Broker enabled symbols for trading
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Broker As cBroker                   ' Broker base class object
    BrokerKv As cBrokerKeyValue         ' Routines for key-value broker messages
    Symbols As cBrokerSymbols           ' Broker symbol conversion object
    
    strUserName As String               ' User name to login to the server with
    strPassword As String               ' Password for logging into the server
    strUrl As String                    ' URL for the server
    
    AmendOrders As cGdTree              ' Collection of orders that have been sent for amend
    astrFillSymbols As cGdArray         ' Array of symbols for which we got fills
End Type
Private m As mPrivate

Public Property Get Broker() As cBroker
    Set Broker = m.Broker
End Property

Public Property Get AmendOrders() As cGdTree
    Set AmendOrders = m.AmendOrders
End Property

Public Property Get EnabledSymbols() As cGdArray
    Set EnabledSymbols = Nothing
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
        Case eTT_AccountType_Gft
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "GFT", True
            m.Broker.IniFile = AddSlash(App.Path) & "GFT.INI"
            m.Broker.BrokerName = "GFT Forex"
            m.Broker.ProcessName = "GenGft"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\GFT\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenGft.EXE"
            m.Broker.ControlID = "GFT"
            strSymbolFile = AddSlash(App.Path) & "Provided\GftToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\GftCnct.INI"
            
    End Select

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
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_BrokerView) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewOnline) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_VerifyPositions) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_AccountDetails) = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerGft.Init"

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

    ShowConfig = ((Len(m.strUserName) = 0) Or (Len(m.strPassword) = 0) Or (Len(m.strUrl) = 0))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.ShowConfig"
    
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

    ShowBrokerConnectionInfo = GetLoginInfo(strUserName, , True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.ShowBrokerConnectionInfo"
    
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
    Dim lTimeOut As Long                ' Timeout variable

    If (m.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected) Then
        If (m.Broker.AskedPassword = True) And ((strNewUserName = m.strUserName) Or (Len(strNewUserName) = 0)) Then
            bReturn = True
        Else
            bReturn = GetLoginInfo(strNewUserName)
        End If
    Else
        If ((strNewUserName <> m.strUserName) And (Len(strNewUserName) > 0)) Or (bSwitching = True) Then
            bReturn = GetLoginInfo(strNewUserName, bSwitching)
            If bReturn = True Then
                m.Broker.Disconnect False, "Changing user names to '" & strNewUserName & "'"
                
                lTimeOut = 0&
                Do While (m.Broker.ConnectionStatus <> eGDConnectionStatus_Disconnected) And (lTimeOut < 30&)
                    Sleep 1
                    lTimeOut = lTimeOut + 1&
                Loop
                
                bReturn = (m.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected)
            End If
        End If
    End If
    
    ShowLogin = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.ShowLogin"
    
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

    m.Broker.DumpDebug "Connecting to " & m.Broker.BrokerName & " as '" & m.strUserName & "'"
    m.Broker.SendBrokerMessage eGDBrokerMessageType_Connect, ConnectString, , ConnectString(True)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerGft.SendConnect"
    
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
    RaiseError "cBrokerGft.SendDisconnect"
    
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
    RaiseError "cBrokerGft.UnloadApp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerSymbol
'' Description: Attempt to convert the Genesis symbol to the broker symbology
'' Inputs:      Genesis Symbol, Broker Symbol object
'' Returns:     Broker Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerSymbol(ByVal strGenesisSymbol As String, Optional BrokerSym As cBrokerSymbol) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Len(strGenesisSymbol) > 0 Then
        If IsForex(strGenesisSymbol) Then
            If InStr(strGenesisSymbol, "@") = 0 Then
                Set BrokerSym = m.Symbols.GetByGenesisSymbol("$")
                
                strReturn = Mid(Replace(strGenesisSymbol, "-", "/"), 2)
            End If
        End If
    End If
    
    BrokerSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.BrokerSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Attempt to convert the broker symbol to the Genesis symbology
'' Inputs:      Broker Symbol, Broker Exchange, Broker Symbol object, For Position?
'' Returns:     Genesis Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbol(ByVal strBrokerSymbol As String, ByVal strBrokerExchange As String, Optional BrokerSym As cBrokerSymbol, Optional ByVal bPosition As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Len(strBrokerSymbol) = 7 Then
        Set BrokerSym = m.Symbols.GetByBrokerSymbol("$")
        strReturn = "$" & Replace(strBrokerSymbol, "/", "-")
        If GetSymbolID(strReturn) = 0 Then
            strReturn = ""
        End If
    End If
    
    GenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.GenesisSymbol"
    
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
    RaiseError "cBrokerGft.BrokerOrderType"
    
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
    
    If (Len(strOrderString) > 0) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AddOrder, strOrderString
        bReturn = True
    End If
    
    AddOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.AddOrder"
    
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
    Dim NewOrder As cPtOrder            ' New Order object
    
    bReturn = False
    
    Set NewOrder = Order.MakeCopy
    NewOrder.GenesisOrderID = NextGenesisID(g.Broker.AccountNumberForID(Order.AccountID))
    strOrderString = BrokerOrderString(NewOrder, False, Order.GenesisOrderID)
        
    If (Len(strOrderString) > 0) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AmendOrder, strOrderString
        bReturn = True
    End If
    
    AmendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.AmendOrder"
    
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
    Dim NewOrder As cPtOrder            ' New Order object
    
    bReturn = False
    
    Set NewOrder = Order.MakeCopy
    NewOrder.GenesisOrderID = NextGenesisID(g.Broker.AccountNumberForID(Order.AccountID))
    strOrderString = BrokerOrderString(NewOrder, False, Order.GenesisOrderID)
    
    If (Len(strOrderString) > 0) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_CancelOrder, strOrderString
        bReturn = True
    End If
    
    CancelOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.CancelOrder"
    
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
    Dim NewOrder As cPtOrder            ' New Order object
    
    bReturn = False
    
    Set NewOrder = Order.MakeCopy
    NewOrder.GenesisOrderID = NextGenesisID(g.Broker.AccountNumberForID(Order.AccountID))
    strOrderString = BrokerOrderString(NewOrder, False, Order.GenesisOrderID)
        
    If (Len(strOrderString) > 0) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_CancelOrder, strOrderString
        bReturn = True
    End If
    
    ParkOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.ParkOrder"
    
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
    RaiseError "cBrokerGft.Refresh"
    
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
    RaiseError "cBrokerGft.GetAccounts"
    
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

    Dim brokerMsg As cBrokerMessage     ' Broker message
    
    If Len(strAccount) = 0 Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_GetOrders, ""
    Else
        Set brokerMsg = New cBrokerMessage
        brokerMsg.Add "Account", strAccount
        m.Broker.SendBrokerMessage eGDBrokerMessageType_GetOrders, brokerMsg.ToString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerGft.GetOrders"
    
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

    Dim brokerMsg As cBrokerMessage     ' Broker message
    
    If Len(strAccount) = 0 Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_GetFills, ""
    Else
        Set brokerMsg = New cBrokerMessage
        brokerMsg.Add "Account", strAccount
        m.Broker.SendBrokerMessage eGDBrokerMessageType_GetFills, brokerMsg.ToString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerGft.GetFills"
    
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

    Dim brokerMsg As cBrokerMessage     ' Broker message
    
    If Len(strAccount) = 0 Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_GetPositions, ""
    Else
        Set brokerMsg = New cBrokerMessage
        brokerMsg.Add "Account", strAccount
        m.Broker.SendBrokerMessage eGDBrokerMessageType_GetPositions, brokerMsg.ToString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerGft.GetPositions"
    
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
    Dim strMachineID As String          ' Machine ID
    Dim strIniFile As String            ' Ini File
    
    strIniFile = m.Broker.IniFile
    lDate = GetIniFileProperty("Date", 0&, "ID", strIniFile)
    lCounter = GetIniFileProperty("Counter", 0&, "ID", strIniFile)
    strMachineID = StripStr(UCase(RI_GetMachineID), "- ")
    
    If lDate <> Date Then
        lDate = Date
        lCounter = 0&
    Else
        lCounter = lCounter + 1
    End If

    SetIniFileProperty "Date", lDate, "ID", strIniFile
    SetIniFileProperty "Counter", lCounter, "ID", strIniFile
    
    NextGenesisID = "TN_" & strMachineID & "_" & Format(lDate, "00000") & Format(lCounter, "00000")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.NextGenesisID"
    
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
    RaiseError "cBrokerGft.StartingGenesisIdForOptNav"
    
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
    RaiseError "cBrokerGft.IsTradeableSymbol"
    
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
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = (Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0)
    If bReturn Then
        strBrokerBase = BrokerSym.BrokerBase
        strBrokerExchange = BrokerSym.BrokerExchange
    End If
    
    IsEnabledSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.IsEnabledSymbol"
    
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
    RaiseError "cBrokerGft.OrderTypeAllowed"
    
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
    RaiseError "cBrokerGft.TimeInForceAllowed"
    
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
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTemplate.HandleMessage"
    
End Sub

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
    strUserName = m.strUserName
    
    If Len(brokerMsg("AllowReconnect")) = 0 Then
        bAllowReconnect = True
    Else
        bAllowReconnect = (CLng(Val(brokerMsg("AllowReconnect"))) <> 0)
    End If
    
    StatusFromInfo = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.StatusFromInfo"
    
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

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
            
            Set Acct = m.BrokerKv.AccountFromMessage(brokerMsg, m.Broker, m.strUserName, m.strPassword)
            bReturn = True
        End If
    End If
    
    AccountFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.AccountFromInfo"
    
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
    Dim brokerMsg As cBrokerMessage     ' Broker message object
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim lQuantityFilled As Long         ' Quantity of the order that has filled
    Dim Bars As New cGdBars             ' Bars object
    Dim AmendedOrder As cPtOrder        ' Order object
    Dim strPrevGenesisID As String      ' Previous Genesis ID
    Dim strError As String              ' Error Message
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            
            brokerMsg.FromString strMessage
            
            strGenesisSymbol = GenesisSymbol(brokerMsg("Symbol"), "", BrokerSym)
            strError = brokerMsg("Message")
            strPrevGenesisID = brokerMsg("PreviousGenesisID")
            
            If Len(strPrevGenesisID) > 0 Then
                If (UCase(brokerMsg("Status")) = "REPLACED") Or (UCase(brokerMsg("Status")) = "NEW") Then
                    If m.AmendOrders.Exists(strPrevGenesisID) Then
                        If Len(strError) > 0 Then
                            Set AmendedOrder = m.AmendOrders(strPrevGenesisID)
                            AmendedOrder.ChangeOrderStatus eTT_OrderStatus_Working
                            
                            InfBox strError, , , m.Broker.BrokerName & " Order Error"
                            strGenesisSymbol = ""
                        Else
                            Set AmendedOrder = m.AmendOrders(strPrevGenesisID)
                            AmendedOrder.ChangeOrderStatus eTT_OrderStatus_Amended
                        End If
                        
                        m.AmendOrders.Remove strPrevGenesisID
                    End If
                End If
            End If
            
            If Len(strGenesisSymbol) > 0 Then
                Set Order = m.BrokerKv.OrderFromMessage(brokerMsg, m.Broker, strGenesisSymbol, BrokerSym)
                bReturn = True
            End If
        End If
    End If
    
    OrderFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.OrderFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillFromInfo
'' Description: Fill a fill object from the given message
'' Inputs:      Message, Fill, Genesis Order ID, Carried?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillFromInfo(ByVal strMessage As String, Fill As cPtFill, Optional strGenesisOrderID As String = "", Optional ByVal bCarried As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker message object
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim Bars As New cGdBars             ' Bars object
    Dim lPos As Long                    ' Position of the symbol in the array
    Dim strKey As String                ' Key into the fills array
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
            m.astrFillSymbols.Clear
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            
            brokerMsg.FromString strMessage
            
            strGenesisSymbol = GenesisSymbol(brokerMsg("Symbol"), "", BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                strKey = brokerMsg("Account") & vbTab & brokerMsg("Symbol")
                If m.astrFillSymbols.BinarySearch(strKey, lPos) = False Then
                    m.astrFillSymbols.Add strKey, lPos
                End If
                
                brokerMsg("BrokerID") = ""
                Set Fill = m.BrokerKv.FillFromMessage(brokerMsg, m.Broker, strGenesisSymbol, BrokerSym)
                strGenesisOrderID = brokerMsg("GenesisID")
                bReturn = True
            End If
        End If
    End If
    
    FillFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.FillFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CarriedFillFromInfo
'' Description: Fill a fill object from the given message
'' Inputs:      Message, Fill, Genesis Order ID
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CarriedFillFromInfo(ByVal strMessage As String, Fill As cPtFill, Optional strGenesisOrderID As String = "") As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            bReturn = FillFromInfo(strMessage, Fill, strGenesisOrderID, True)
        End If
    End If
    
    CarriedFillFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.CarriedFillFromInfo"
    
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
    Dim brokerMsg As cBrokerMessage     ' Broker message object
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim lIndex As Long                  ' Index into a for loop
    Dim carriedFill As cPtFill          ' Carried fill
    Dim Fills As cPtFills               ' Fills
    Dim lPos As Long                    ' Position of the symbol in the array
    Dim strKey As String                ' Key into the fills array
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
            For lIndex = 0 To m.astrFillSymbols.Size - 1
                Set brokerMsg = New cBrokerMessage
                brokerMsg.Add "Account", Parse(m.astrFillSymbols(lIndex), vbTab, 1)
                brokerMsg.Add "Symbol", Parse(m.astrFillSymbols(lIndex), vbTab, 2)
                
                m.Broker.PositionReceived brokerMsg.ToString, True
            Next lIndex
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
            
            strGenesisSymbol = GenesisSymbol(brokerMsg("Symbol"), "", BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                strKey = brokerMsg("Account") & vbTab & brokerMsg("Symbol")
                If m.astrFillSymbols.BinarySearch(strKey, lPos) = True Then
                    m.astrFillSymbols.Remove lPos
                End If
                
                Set Position = m.BrokerKv.PositionFromMessage(brokerMsg, m.Broker, strGenesisSymbol)
                
                Position.CurrentPosition = 0
                If Not m.Broker Is Nothing Then
                    For lIndex = 1 To m.Broker.CarriedFills.Count
                        Set carriedFill = m.Broker.CarriedFills(lIndex)
                        If carriedFill.Symbol = strGenesisSymbol Then
                            Position.CarriedFills.Add carriedFill
                            If carriedFill.Buy Then
                                Position.CurrentPosition = Position.CurrentPosition + carriedFill.Quantity
                            Else
                                Position.CurrentPosition = Position.CurrentPosition - carriedFill.Quantity
                            End If
                        End If
                    Next lIndex
                    
                    Set Fills = m.Broker.BrokerInfo.Fills
                    For lIndex = 1 To Fills.Count
                        If Fills(lIndex).AccountID = Position.AccountID Then
                            If Fills(lIndex).SymbolOrSymbolID = Position.SymbolOrSymbolID Then
                                If Fills(lIndex).Buy Then
                                    Position.CurrentPosition = Position.CurrentPosition + Fills(lIndex).Quantity
                                Else
                                    Position.CurrentPosition = Position.CurrentPosition - Fills(lIndex).Quantity
                                End If
                            End If
                        End If
                    Next lIndex
                End If
                
                bReturn = True
            End If
        End If
    End If
    
    PositionFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.PositionFromInfo"
    
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
    RaiseError "cBrokerGft.GenesisSymbolList"
    
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
    RaiseError "cBrokerGft.TimeZone"
    
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
    RaiseError "cBrokerGft.JustConnected"
    
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
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol information
    
    strReturn = ""
    If Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0 Then
        strReturn = BrokerSym.OrderTypeMask & vbTab & BrokerSym.TifMask
    End If
    
    SymbolInformation = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.SymbolInformation"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectString
'' Description: Build a string to send for a Connect call
'' Inputs:      For Log?
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ConnectString(Optional ByVal bForLog As Boolean = False) As String
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    
    Set brokerMsg = New cBrokerMessage
    
    brokerMsg.Add "User", m.strUserName
    brokerMsg.Add "Pass", m.strPassword
    brokerMsg.Add "URL", m.strUrl
    
    ConnectString = brokerMsg.ToString(bForLog)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.ConnectString"
    
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
    RaiseError "cBrokerGft.BrokerOrderString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLoginInfo
'' Description: Get login information from the user
'' Inputs:      User Name, Switching, Show IP?
'' Returns:     True if user OKed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetLoginInfo(Optional ByVal strUserName As String = "", Optional ByVal bSwitching As Boolean = False, Optional ByVal bShowIP As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = frmLoginBrokerUrl.ShowMe(m.Broker, strUserName, bSwitching, bShowIP)
    If bReturn = True Then
        m.strUserName = frmLoginBrokerUrl.UserName
        m.strPassword = frmLoginBrokerUrl.Password
        m.strUrl = frmLoginBrokerUrl.Url
        
        m.Broker.AskedPassword = True
    End If
    
    GetLoginInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerGft.GetLoginInfo"
    
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
    Set m.AmendOrders = New cGdTree
    
    Set m.astrFillSymbols = New cGdArray
    m.astrFillSymbols.Create eGDARRAY_Strings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerGft.Class_Initialize"
    
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
    Set m.AmendOrders = Nothing
    Set m.astrFillSymbols = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerGft.Class_Terminate"
    
End Sub
