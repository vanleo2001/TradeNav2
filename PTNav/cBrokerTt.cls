VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerTt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerTt.cls
'' Description: Class to manange communications with Trading Technologies servers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 09/07/2011   DAJ         On a refresh, call GetOrders instead of GetAccounts
'' 09/08/2011   DAJ         Use GetAccounts, but generate them here
'' 09/20/2011   DAJ         Numerous fixes
'' 09/26/2011   DAJ         Fix for calculating carried position
'' 10/25/2011   DAJ         Handle carried fills
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
    strIP As String                     ' IP address for the server
    strPort As String                   ' Port for the server
    strTargetID As String               ' Target computer ID
    strPriceIP As String                ' IP address for the price server
    strPricePort As String              ' Port for the price server
    strPriceTargetID As String          ' Target computer ID for the price server
    
    strSubscribed As String             ' Symbol that is subscribed to
    
    AmendOrders As cGdTree              ' Collection of orders that have been sent for amend
End Type
Private m As mPrivate

Public Property Get Broker() As cBroker
    Set Broker = m.Broker
End Property

Public Property Get Symbols() As cBrokerSymbols
    Set Symbols = m.Symbols
End Property

Public Property Get AmendOrders() As cGdTree
    Set AmendOrders = m.AmendOrders
End Property

Public Property Get ConnectToPrice() As Boolean
    ConnectToPrice = FileExist(AddSlash(App.Path) & "TTPrice.Flg")
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

    Dim strSymbolFile As String         ' Symbol file
    
    Select Case nBroker
        Case eTT_AccountType_AdvFut
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "AdvantageFutures"
            m.Broker.IniFile = AddSlash(App.Path) & "AdvantageFutures.INI"
            m.Broker.BrokerName = "Advantage Futures"
            m.Broker.ProcessName = "GenAdvantageFutures"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\AdvantageFutures\GenAdvantageFutures.EXE"
            m.Broker.ControlID = "AdvantageFutures"
            strSymbolFile = AddSlash(App.Path) & "Provided\AfToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\AdvFutIps.INI"
        
        Case eTT_AccountType_TT
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "TradingTechnologies"
            m.Broker.IniFile = AddSlash(App.Path) & "TradingTechnologies.INI"
            m.Broker.BrokerName = "Trading Technologies"
            m.Broker.ProcessName = "GenTradingTechnologies"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\TradingTechnologies\GenTradingTechnologies.EXE"
            m.Broker.ControlID = "TradingTechnologies"
            strSymbolFile = AddSlash(App.Path) & "Provided\TtToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\TtIps.INI"
    
    End Select

    m.Broker.CopyProcess AddSlash(App.Path) & "..\Brokers\GenTradingTechnologies.EXE"
    
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
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_BrokerView) = ConnectToPrice
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewOnline) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_VerifyPositions) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_AccountDetails) = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTt.Init"

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

    ShowConfig = ((Len(m.strUserName) = 0) Or (Len(m.strPassword) = 0) Or (Len(m.strIP) = 0) Or (Len(m.strPort) = 0) Or (Len(m.strTargetID) = 0))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTt.ShowConfig"
    
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
    RaiseError "cBrokerTt.ShowBrokerConnectionInfo"
    
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
    RaiseError "cBrokerTt.ShowLogin"
    
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
    RaiseError "cBrokerTt.SendConnect"
    
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

    If Len(m.strSubscribed) > 0 Then
        Unsubscribe m.strSubscribed
    End If

    m.Broker.SendBrokerMessage eGDBrokerMessageType_Disconnect, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTt.SendDisconnect"
    
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
    RaiseError "cBrokerTt.UnloadApp"
    
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
    Dim astrFields As cGdArray          ' Fields out of the symbol
    
    strReturn = ""
    If Len(strGenesisSymbol) > 0 Then
        Select Case UCase(SecurityType(strGenesisSymbol, True))
            Case "F"
                Set BrokerSym = m.Symbols.GetByGenesisSymbol(Parse(strGenesisSymbol, "-", 1))
                If Not BrokerSym Is Nothing Then
                    Set astrFields = New cGdArray
                    
                    astrFields(0) = BrokerSym.BrokerBase
                    astrFields(1) = Parse(strGenesisSymbol, "-", 2)
                    astrFields(2) = BrokerSym.BrokerExchange
                    astrFields(3) = "FUT"
                    
                    'strReturn = brokerSym.BrokerBase & "01" & Parse(strGenesisSymbol, "-", 2) & "0000003232"
                    strReturn = astrFields.JoinFields("|")
                End If
                
        End Select
    End If
    
    BrokerSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTt.BrokerSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Attempt to convert the broker symbol to the Genesis symbology
'' Inputs:      Broker Symbol, Broker Exchange, Broker Symbol object
'' Returns:     Genesis Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbol(ByVal strBrokerSymbol As String, ByVal strBrokerExchange As String, Optional BrokerSym As cBrokerSymbol) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrFields As cGdArray          ' Fields out of the symbol
    
    strReturn = ""
    If Len(strBrokerSymbol) > 0 Then
        Set astrFields = New cGdArray
        astrFields.SplitFields strBrokerSymbol, "|"
        
        Select Case UCase(astrFields(3))
            Case "FUT"
                Set BrokerSym = m.Symbols.GetByBrokerSymbol(astrFields(0), astrFields(2))
                If Not BrokerSym Is Nothing Then
                    strReturn = BrokerSym.GenesisBase & "-" & astrFields(1)
                End If
        
        End Select
    End If
    
    GenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTt.GenesisSymbol"
    
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
    RaiseError "cBrokerTt.BrokerOrderType"
    
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
    RaiseError "cBrokerTt.AddOrder"
    
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
    RaiseError "cBrokerTt.AmendOrder"
    
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
    RaiseError "cBrokerTt.CancelOrder"
    
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
    RaiseError "cBrokerTt.ParkOrder"
    
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
    RaiseError "cBrokerTt.Refresh"
    
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

    Dim rs As Recordset                 ' Recordset into the database
    
    m.Broker.HandleMessage eGDBrokerMessageType_AccountRefresh, "Begin"
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountType]=" & Str(m.Broker.Broker) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        m.Broker.HandleMessage eGDBrokerMessageType_AccountRefresh, "Account=" & rs!AccountNumber
        rs.MoveNext
    Loop
    m.Broker.HandleMessage eGDBrokerMessageType_AccountRefresh, "End"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTt.GetAccounts"
    
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
    RaiseError "cBrokerTt.GetOrders"
    
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
    RaiseError "cBrokerTt.GetFills"
    
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
    RaiseError "cBrokerTt.GetPositions"
    
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
    RaiseError "cBrokerTt.NextGenesisID"
    
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
    RaiseError "cBrokerTt.StartingGenesisIdForOptNav"
    
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
    RaiseError "cBrokerTt.IsTradeableSymbol"
    
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
    RaiseError "cBrokerTt.IsEnabledSymbol"
    
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
    RaiseError "cBrokerTt.OrderTypeAllowed"
    
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
    RaiseError "cBrokerTt.TimeInForceAllowed"
    
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
        Case eGDBrokerMessageType_PriceUpdate
            HandlePriceUpdate strMessage
            
    End Select
    
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
    RaiseError "cBrokerTt.StatusFromInfo"
    
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
    RaiseError "cBrokerTt.AccountFromInfo"
    
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
            
            strGenesisSymbol = GenesisSymbol(brokerMsg.Symbol, brokerMsg("Exchange"), BrokerSym)
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
    RaiseError "cBrokerTt.OrderFromInfo"
    
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
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            
            brokerMsg.FromString strMessage
            
            strGenesisSymbol = GenesisSymbol(brokerMsg.Symbol, brokerMsg("Exchange"), BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
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
    RaiseError "cBrokerTt.FillFromInfo"
    
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
    RaiseError "cBrokerTt.CarriedFillFromInfo"
    
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
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
            
            strGenesisSymbol = GenesisSymbol(brokerMsg.Symbol, brokerMsg("Exchange"), BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                Set Position = m.BrokerKv.PositionFromMessage(brokerMsg, m.Broker, strGenesisSymbol)
                
                Position.CarriedPosition = 0
                If Not m.Broker Is Nothing Then
                    For lIndex = 1 To m.Broker.CarriedFills.Count
                        Set carriedFill = m.Broker.CarriedFills(lIndex)
                        If carriedFill.Symbol = strGenesisSymbol Then
                            Position.CarriedFills.Add carriedFill
                            If carriedFill.Buy Then
                                Position.CarriedPosition = Position.CarriedPosition + carriedFill.Quantity
                            Else
                                Position.CarriedPosition = Position.CarriedPosition - carriedFill.Quantity
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
    RaiseError "cBrokerTt.PositionFromInfo"
    
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
    RaiseError "cBrokerTt.GenesisSymbolList"
    
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
    RaiseError "cBrokerTt.TimeZone"
    
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

    If FormIsLoaded("frmBrokerSnapshot") Then
        If frmBrokerSnapshot.Broker = m.Broker.Broker Then
            If Len(frmBrokerSnapshot.Symbol) > 0 Then
                Subscribe frmBrokerSnapshot.Symbol
            End If
        End If
    End If

    GetAccounts

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTt.JustConnected"
    
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
    RaiseError "cBrokerTt.SymbolInformation"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Subscribe
'' Description: Subscribe to data for the given symbol
'' Inputs:      Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Subscribe(ByVal strGenesisSymbol As String)
On Error GoTo ErrSection:

    Dim brokerMessage As cBrokerMessage ' Broker message to send to the broker
    Dim strSymbol As String             ' Broker symbol
    
    strSymbol = BrokerSymbol(strGenesisSymbol)
    If Len(strSymbol) > 0 Then
        Set brokerMessage = New cBrokerMessage
        brokerMessage.Symbol = strSymbol
        m.Broker.SendBrokerMessage eGDBrokerMessageType_Subscribe, brokerMessage.ToString
        
        m.strSubscribed = strGenesisSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTt.Subscribe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Unsubscribe
'' Description: Unsubscribe from data for the given symbol
'' Inputs:      Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Unsubscribe(ByVal strGenesisSymbol As String)
On Error GoTo ErrSection:

    Dim brokerMessage As cBrokerMessage ' Broker message to send to the broker
    Dim strSymbol As String             ' Broker symbol
    
    strSymbol = BrokerSymbol(strGenesisSymbol)
    If Len(strSymbol) > 0 Then
        Set brokerMessage = New cBrokerMessage
        brokerMessage.Symbol = strSymbol
        m.Broker.SendBrokerMessage eGDBrokerMessageType_Unsubscribe, brokerMessage.ToString
        
        m.strSubscribed = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTt.Unsubscribe"
    
End Sub

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
    brokerMsg.Add "IP", m.strIP
    brokerMsg.Add "Port", m.strPort
    brokerMsg.Add "Target", m.strTargetID
    
    If ConnectToPrice Then
        brokerMsg.Add "PriceIP", m.strPriceIP
        brokerMsg.Add "PricePort", m.strPricePort
        brokerMsg.Add "PriceTarget", m.strPriceTargetID
    End If
    
    ConnectString = brokerMsg.ToString(bForLog)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTt.ConnectString"
    
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
    RaiseError "cBrokerTt.BrokerOrderString"
    
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
    Dim bShowPrice As Boolean           ' Show price server information?
    
    bShowPrice = ConnectToPrice

    bReturn = frmLoginFix.ShowMe(m.Broker, strUserName, bSwitching, bShowIP, bShowPrice)
    If bReturn = True Then
        m.strUserName = frmLoginFix.UserName
        m.strPassword = frmLoginFix.Password
        m.strIP = frmLoginFix.IP
        m.strPort = frmLoginFix.Port
        m.strTargetID = frmLoginFix.TargetID
        
        If bShowPrice Then
            m.strPriceIP = frmLoginFix.PriceIP
            m.strPricePort = frmLoginFix.PricePort
            m.strPriceTargetID = frmLoginFix.PriceTargetID
        End If
        
        m.Broker.AskedPassword = True
    End If
    
    GetLoginInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerTt.GetLoginInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandlePriceUpdate
'' Description: Handle a price update from the TT servers
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandlePriceUpdate(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim brokerMessage As cBrokerMessage ' Broker message
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim dLastPrice As Double            ' Last price for the symbol
    Dim dBidPrice As Double             ' Bid price for the symbol
    Dim dAskPrice As Double             ' Ask price for the symbol
    
    If Len(strMessage) > 0 Then
        Set brokerMessage = New cBrokerMessage
        brokerMessage.FromString strMessage
        
        strGenesisSymbol = GenesisSymbol(brokerMessage.Symbol, brokerMessage("Exchange"), BrokerSym)
        If Len(strGenesisSymbol) > 0 Then
            If FormIsLoaded("frmBrokerSnapshot") Then
                If frmBrokerSnapshot.Broker = m.Broker.Broker Then
                    dLastPrice = kNullData
                    dBidPrice = kNullData
                    dAskPrice = kNullData
                    
                    If Len(brokerMessage("LastPrice")) > 0 Then
                        dLastPrice = m.Broker.GenesisPrice(brokerMessage("LastPrice"), BrokerSym.PriceMult)
                    End If
                    If Len(brokerMessage("BidPrice")) > 0 Then
                        dBidPrice = m.Broker.GenesisPrice(brokerMessage("BidPrice"), BrokerSym.PriceMult)
                    End If
                    If Len(brokerMessage("AskPrice")) > 0 Then
                        dAskPrice = m.Broker.GenesisPrice(brokerMessage("AskPrice"), BrokerSym.PriceMult)
                    End If
                    
                    frmBrokerSnapshot.Broker_DataUpdate strGenesisSymbol, dLastPrice, dBidPrice, dAskPrice
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTt.HandlePriceUpdate"
    
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
    Set m.AmendOrders = New cGdTree
    
    m.strSubscribed = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTt.Class_Initialize"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTt.Class_Terminate"
    
End Sub

