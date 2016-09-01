VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerCurrenex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerCurrenex.cls
'' Description: Class for handling communication with Currenex
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/21/2012   DAJ         Changed the GetAccounts call to pull from database
'' 07/13/2012   DAJ         Fix for converting to broker symbol
'' 07/16/2012   DAJ         KnightCnx
'' 08/17/2012   DAJ         Design changes / fixes
'' 08/28/2012   DAJ         CarriedFillFromInfo, First Field Begin/End
'' 08/29/2012   DAJ         Zaner (Currenex)
'' 09/12/2012   DAJ         Added Currenex, FXDD (Currenex), and VanKar (Currenex)
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 08/30/2013   DAJ         Create fake position records for symbols for which we got a fill
'' 09/05/2013   DAJ         Fix for creating fake position records
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Broker As cBroker                   ' Broker base class object
    BrokerKv As cBrokerKeyValue         ' Routines for key-value broker messages
    Symbols As cBrokerSymbols           ' Broker symbol conversion object
    
    strUserName As String               ' User name to login to the server with
    strPassword As String               ' Password for logging into the server
    strAccount As String                ' Account the user is logged into
    strIP As String                     ' IP address for the server
    strPort As String                   ' Port for the server
    strTargetID As String               ' Target computer ID
    strRefreshServer As String          ' Refresh server information
    
    astrFillSymbols As cGdArray         ' Array of symbols for which we got fills
    SentOrders As cGdTree               ' Collection of orders sent to Currenex
    AmendOrders As cGdTree              ' Collection of orders that have been sent for amend
    CancelOrders As cGdTree             ' Collection of orders that have been sent for cancel
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
        Case eTT_AccountType_Currenex
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "Currenex", True
            m.Broker.IniFile = AddSlash(App.Path) & "Currenex.INI"
            m.Broker.BrokerName = "Currenex"
            m.Broker.ProcessName = "GenCurrenex"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\Currenex\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenCurrenex.EXE"
            m.Broker.ControlID = "Currenex"
            strSymbolFile = AddSlash(App.Path) & "Provided\CnxToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\CnxCnct.INI"
            
        Case eTT_AccountType_AlpariCurrenex
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "AlpariCnx", True
            m.Broker.IniFile = AddSlash(App.Path) & "AlpariCnx.INI"
            m.Broker.BrokerName = "Alpari (Currenex)"
            m.Broker.ProcessName = "GenAlpariCnx"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\AlpariCnx\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenAlpariCnx.EXE"
            m.Broker.ControlID = "AlpariCnx"
            strSymbolFile = AddSlash(App.Path) & "Provided\AlpCnxToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\AlpCnxCnct.INI"
            
        Case eTT_AccountType_FxddCurrenex
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "FxddCnx", True
            m.Broker.IniFile = AddSlash(App.Path) & "FxddCnx.INI"
            m.Broker.BrokerName = "FXDD (Currenex)"
            m.Broker.ProcessName = "GenFxddCnx"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\FxddCnx\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenFxddCnx.EXE"
            m.Broker.ControlID = "FxddCnx"
            strSymbolFile = AddSlash(App.Path) & "Provided\FxddCnxToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\FxddCnxCnct.INI"
            
        Case eTT_AccountType_KnightCurrenex
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "KnightCnx", True
            m.Broker.IniFile = AddSlash(App.Path) & "KnightCnx.INI"
            m.Broker.BrokerName = "Knight (Currenex)"
            m.Broker.ProcessName = "GenKnightCnx"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\KnightCnx\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenKnightCnx.EXE"
            m.Broker.ControlID = "KnightCnx"
            strSymbolFile = AddSlash(App.Path) & "Provided\KnightCnxToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\KnightCnxCnct.INI"
    
        Case eTT_AccountType_VanKarCurrenex
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "VanKarCnx", True
            m.Broker.IniFile = AddSlash(App.Path) & "VanKarCnx.INI"
            m.Broker.BrokerName = "VanKar (Currenex)"
            m.Broker.ProcessName = "GenVanKarCnx"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\VanKarCnx\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenVanKarCnx.EXE"
            m.Broker.ControlID = "VanKarCnx"
            strSymbolFile = AddSlash(App.Path) & "Provided\VanKarCnxToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\VanKarCnxCnct.INI"
            
        Case eTT_AccountType_ZanerCurrenex
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "ZanerCnx", True
            m.Broker.IniFile = AddSlash(App.Path) & "ZanerCnx.INI"
            m.Broker.BrokerName = "Zaner (Currenex)"
            m.Broker.ProcessName = "GenZanerCnx"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\ZanerCnx\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenZanerCnx.EXE"
            m.Broker.ControlID = "ZanerCnx"
            strSymbolFile = AddSlash(App.Path) & "Provided\ZanerCnxToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\ZanerCnxCnct.INI"
    
    End Select

    m.Broker.CopyProcess AddSlash(App.Path) & "..\Brokers\GenCurrenex.EXE"
    FileCopy AddSlash(App.Path) & "..\Brokers\FixDefCnx.xml", AddSlash(strBrokerPath) & "FixDef.xml", True
    
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
    RaiseError "cBrokerCurrenex.Init"

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

    ShowConfig = ((Len(m.strUserName) = 0) Or (Len(m.strPassword) = 0) Or (Len(m.strIP) = 0) Or (Len(m.strPort) = 0) Or (Len(m.strTargetID) = 0) Or (Len(m.strRefreshServer) = 0))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCurrenex.ShowConfig"
    
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
    RaiseError "cBrokerCurrenex.ShowBrokerConnectionInfo"
    
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
    RaiseError "cBrokerCurrenex.ShowLogin"
    
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
    RaiseError "cBrokerCurrenex.SendConnect"
    
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
    RaiseError "cBrokerCurrenex.SendDisconnect"
    
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
    RaiseError "cBrokerCurrenex.UnloadApp"
    
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
        Select Case UCase(SecurityType(strGenesisSymbol, True))
            Case "I"
                If IsForex(strGenesisSymbol) Then
                    Set BrokerSym = m.Symbols.GetByGenesisSymbol(strGenesisSymbol)
                    If Not BrokerSym Is Nothing Then
                        strReturn = BrokerSym.BrokerBase
                    End If
                End If
                
        End Select
    End If
    
    BrokerSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCurrenex.BrokerSymbol"
    
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
    If Len(strBrokerSymbol) > 0 Then
        If bPosition Then
            Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol(strBrokerSymbol, strBrokerExchange)
        Else
            Set BrokerSym = m.Symbols.GetByBrokerSymbol(strBrokerSymbol, strBrokerExchange)
        End If
        If Not BrokerSym Is Nothing Then
            strReturn = BrokerSym.GenesisBase
        End If
    End If
    
    GenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCurrenex.GenesisSymbol"
    
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
    RaiseError "cBrokerCurrenex.BrokerOrderType"
    
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
        
        UpdateSentOrder strOrderString
        
        bReturn = True
    End If
    
    AddOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCurrenex.AddOrder"
    
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
    'Dim NewOrder As cPtOrder            ' New Order object
    
    bReturn = False
    
    ' DAJ 08/15/2012: Send the original Genesis ID from now on ( otherwise refreshes
    ' will come back with both the old and new Genesis ID's as working )...
    'Set NewOrder = Order.MakeCopy
    'NewOrder.GenesisOrderID = NextGenesisID(g.Broker.AccountNumberForID(Order.AccountID))
    'strOrderString = BrokerOrderString(NewOrder, False, Order.GenesisOrderID)
    
    strOrderString = BrokerOrderString(Order, False, Order.GenesisOrderID)
    If (Len(strOrderString) > 0) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AmendOrder, strOrderString
        
        UpdateSentOrder strOrderString
        
        bReturn = True
    End If
    
    AmendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCurrenex.AmendOrder"
    
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
    'Dim NewOrder As cPtOrder            ' New Order object
    
    bReturn = False
    
    ' DAJ 08/15/2012: Send the original Genesis ID from now on...
    'NewOrder.GenesisOrderID = NextGenesisID(g.Broker.AccountNumberForID(Order.AccountID))
    'Set NewOrder = Order.MakeCopy
    'strOrderString = BrokerOrderString(NewOrder, False, Order.GenesisOrderID)
        
    strOrderString = BrokerOrderString(Order, False, Order.GenesisOrderID)
    If (Len(strOrderString) > 0) Then
        m.CancelOrders.Add Order, Order.GenesisOrderID
        m.Broker.SendBrokerMessage eGDBrokerMessageType_CancelOrder, strOrderString
        
        UpdateSentOrder strOrderString
        
        bReturn = True
    End If
    
    CancelOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCurrenex.CancelOrder"
    
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
    RaiseError "cBrokerCurrenex.ParkOrder"
    
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
    RaiseError "cBrokerCurrenex.Refresh"
    
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

    'Dim rs As Recordset                 ' Recordset into the database
    
    m.Broker.HandleMessage eGDBrokerMessageType_AccountRefresh, "Begin"
    'Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountType]=" & Str(m.Broker.Broker) & ";", dbOpenDynaset)
    'Do While Not rs.EOF
    '    m.Broker.HandleMessage eGDBrokerMessageType_AccountRefresh, "Account=" & rs!AccountNumber
    '    rs.MoveNext
    'Loop
    
    m.Broker.HandleMessage eGDBrokerMessageType_AccountRefresh, "Account=" & m.strAccount
    m.Broker.HandleMessage eGDBrokerMessageType_AccountRefresh, "End"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCurrenex.GetAccounts"
    
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

    Set brokerMsg = New cBrokerMessage
    
    brokerMsg.Add "User", m.strUserName
    brokerMsg.Add "Account", m.strAccount
    brokerMsg.Add "StartDate", RefreshStartDate
    brokerMsg.Add "EndDate", RefreshEndDate
    
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetOrders, brokerMsg.ToString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCurrenex.GetOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSingleOrder
'' Description: Request a single order from the server
'' Inputs:      Order, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetSingleOrder(ByVal Order As cPtOrder, ByVal strReason As String)
On Error GoTo ErrSection:

    Dim strOrderInfo As String          ' Order information in a string
    
    m.Broker.DumpDebug "Calling for single order refresh for '" & Order.GenesisOrderID & ";" & Order.BrokerID & "' because '" & strReason & "'"
    
    strOrderInfo = BrokerOrderString(Order, False)
    If Len(strOrderInfo) > 0 Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_CnxGetSingleOrder, strOrderInfo
        UpdateSentOrder strOrderInfo
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCurrenex.GetSingleOrder"
    
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

    Set brokerMsg = New cBrokerMessage
    
    brokerMsg.Add "User", m.strUserName
    brokerMsg.Add "Account", m.strAccount
    brokerMsg.Add "StartDate", RefreshStartDate
    brokerMsg.Add "EndDate", RefreshEndDate
    
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetFills, brokerMsg.ToString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCurrenex.GetFills"
    
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
    
    Set brokerMsg = New cBrokerMessage
    
    brokerMsg.Add "Account", m.strAccount
    
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetPositions, brokerMsg.ToString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCurrenex.GetPositions"
    
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
    RaiseError "cBrokerCurrenex.NextGenesisID"
    
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
    RaiseError "cBrokerCurrenex.StartingGenesisIdForOptNav"
    
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
    RaiseError "cBrokerCurrenex.IsTradeableSymbol"
    
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
    RaiseError "cBrokerCurrenex.IsEnabledSymbol"
    
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
    RaiseError "cBrokerCurrenex.OrderTypeAllowed"
    
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
    RaiseError "cBrokerCurrenex.TimeInForceAllowed"
    
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
        Case eGDBrokerMessageType_CnxSingleOrderRefresh
            SingleOrderReceived strMessage
        
        Case eGDBrokerMessageType_CnxWorkingOrderRefresh
            WorkingOrderReceived strMessage
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCurrenex.HandleMessage"
    
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
    RaiseError "cBrokerCurrenex.StatusFromInfo"
    
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
    RaiseError "cBrokerCurrenex.AccountFromInfo"
    
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
    Dim strPrevGenesisID As String      ' Previous Genesis ID
    Dim strError As String              ' Error Message
    Dim SentOrder As cBrokerMessage     ' Sent order message
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            
            brokerMsg.FromString strMessage
            
            ' Currenex does not send back the account or the order price in an order
            ' message, so we have to use the last one we sent...
            If m.SentOrders.Exists(brokerMsg("GenesisID")) Then
                Set SentOrder = m.SentOrders(brokerMsg("GenesisID"))
                If Len(brokerMsg("Account")) = 0 Then
                    brokerMsg.Add "Account", SentOrder("Account")
                End If
                If Len(brokerMsg("Symbol")) = 0 Then
                    brokerMsg.Add "Symbol", SentOrder("Symbol")
                End If
                If (Len(brokerMsg("StopPrice")) = 0) And (Len(SentOrder("StopPrice")) > 0) Then
                    brokerMsg.Add "StopPrice", SentOrder("StopPrice")
                End If
                If (Len(brokerMsg("LimitPrice")) = 0) And (Len(SentOrder("LimitPrice")) > 0) Then
                    brokerMsg.Add "LimitPrice", SentOrder("LimitPrice")
                End If
                If Len(brokerMsg("Type")) = 0 Then
                    brokerMsg.Add "Type", SentOrder("Type")
                End If
                If Len(brokerMsg("Quantity")) = 0 Then
                    brokerMsg.Add "Quantity", SentOrder("Quantity")
                End If
                If Len(brokerMsg("TIF")) = 0 Then
                    brokerMsg.Add "TIF", SentOrder("TIF")
                End If
            ElseIf Len(brokerMsg("Account")) = 0 Then
                brokerMsg.Add "Account", m.strAccount
            End If
                
            strGenesisSymbol = GenesisSymbol(brokerMsg("Symbol"), "", BrokerSym)
            strError = brokerMsg("Message")
            strPrevGenesisID = brokerMsg("PreviousGenesisID")
            
            If Len(strGenesisSymbol) > 0 Then
                If UCase(brokerMsg("BrokerID")) = "UNKNOWN" Then
                    brokerMsg.Add "BrokerID", ""
                End If
            
                Set Order = m.BrokerKv.OrderFromMessage(brokerMsg, m.Broker, strGenesisSymbol, BrokerSym)
                
                If (UCase(Order.Message) = "SYSTEM CANCEL") Or (UCase(Order.Message) = "USER CANCEL") Then
                    Order.Message = ""
                End If
                
                If m.AmendOrders.Exists(Order.GenesisOrderID) Then
                    If Order.Status = eTT_OrderStatus_Rejected Then
                        GetSingleOrder m.AmendOrders(Order.GenesisOrderID), "Amend Rejected"
                    Else
                        m.AmendOrders.Remove Order.GenesisOrderID
                    End If
                ElseIf m.CancelOrders.Exists(Order.GenesisOrderID) Then
                    If Order.Status = eTT_OrderStatus_Rejected Then
                        GetSingleOrder m.CancelOrders(Order.GenesisOrderID), "Cancel Rejected"
                    Else
                        m.CancelOrders.Remove Order.GenesisOrderID
                    End If
                End If
                
                bReturn = True
            End If
        End If
    End If
    
    OrderFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCurrenex.OrderFromInfo"
    
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
    Dim SentOrder As cBrokerMessage     ' Sent order message
    Dim strFirstField As String         ' First field in the string
    Dim lPos As Long                    ' Position of the symbol in the array
    Dim strKey As String                ' Key into the fills array

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
                brokerMsg("BrokerID") = ""
                            
                ' Currenex does not send back the account in a fill message, so we have
                ' to use the last one we sent...
                If m.SentOrders.Exists(brokerMsg("GenesisID")) Then
                    Set SentOrder = m.SentOrders(brokerMsg("GenesisID"))
                    If Len(brokerMsg("Account")) = 0 Then
                        brokerMsg.Add "Account", SentOrder("Account")
                    End If
                ElseIf Len(brokerMsg("Account")) = 0 Then
                    brokerMsg.Add "Account", m.strAccount
                End If
                    
                strKey = brokerMsg("Account") & vbTab & brokerMsg("Symbol")
                If m.astrFillSymbols.BinarySearch(strKey, lPos) = False Then
                    m.astrFillSymbols.Add strKey, lPos
                End If
                    
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
    RaiseError "cBrokerCurrenex.FillFromInfo"
    
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
    RaiseError "cBrokerCurrenex.CarriedFillFromInfo"
    
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
    Dim strFirstField As String         ' First field in the string
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position of the symbol in the array
    Dim strKey As String                ' Key into the fills array

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
            For lIndex = m.astrFillSymbols.Size - 1 To 0 Step -1
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
                bReturn = True
            End If
        End If
    End If
    
    PositionFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCurrenex.PositionFromInfo"
    
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
    RaiseError "cBrokerCurrenex.GenesisSymbolList"
    
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
    RaiseError "cBrokerCurrenex.TimeZone"
    
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
    RaiseError "cBrokerCurrenex.JustConnected"
    
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
    RaiseError "cBrokerCurrenex.SymbolInformation"
    
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
    Dim astrWeb As cGdArray             ' Web information
    
    Set brokerMsg = New cBrokerMessage
    
    brokerMsg.Add "User", m.strUserName
    brokerMsg.Add "Pass", m.strPassword
    brokerMsg.Add "IP", m.strIP
    brokerMsg.Add "Port", m.strPort
    brokerMsg.Add "Target", m.strTargetID
    brokerMsg.Add "WebServer", m.strRefreshServer
    
    ConnectString = brokerMsg.ToString(bForLog)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCurrenex.ConnectString"
    
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
    RaiseError "cBrokerCurrenex.BrokerOrderString"
    
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

    bReturn = frmLoginFix.ShowMe(m.Broker, strUserName, bSwitching, bShowIP, False, False, False, True, True)
    If bReturn = True Then
        m.strUserName = frmLoginFix.userName
        m.strPassword = frmLoginFix.password
        m.strAccount = frmLoginFix.Account
        m.strIP = frmLoginFix.IP
        m.strPort = frmLoginFix.Port
        m.strTargetID = frmLoginFix.TargetID
        m.strRefreshServer = frmLoginFix.RefreshServer
        
        m.Broker.AskedPassword = True
    End If
    
    GetLoginInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCurrenex.GetLoginInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetWorkingOrders
'' Description: Request a working orders from the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetWorkingOrders()
On Error GoTo ErrSection:

    m.Broker.SendBrokerMessage eGDBrokerMessageType_CnxGetWorkingOrders, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCurrenex.GetWorkingOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    WorkingOrderReceived
'' Description: Handle a working order coming back from the Currenex servers
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub WorkingOrderReceived(ByVal strMessage As String)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCurrenex.WorkingOrderReceived"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SingleOrderReceived
'' Description: Handle a single order coming back from the Currenex servers
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SingleOrderReceived(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the message
    Dim brokerMessage As cBrokerMessage
    
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMessage = New cBrokerMessage
            brokerMessage.FromString strMessage
            
            If UCase(brokerMessage("ItemType")) = "ORDER" Then
                m.Broker.OrderReceived strMessage, False
            ElseIf UCase(brokerMessage("ItemType")) = "FILL" Then
                m.Broker.FillReceived strMessage, False
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCurrenex.SingleOrderReceived"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateSentOrder
'' Description: Update a sent order in the collection
'' Inputs:      Order Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateSentOrder(ByVal strOrderString As String)
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker message object

    Set brokerMsg = New cBrokerMessage
    
    brokerMsg.FromString strOrderString
    If m.SentOrders.Exists(brokerMsg("GenesisID")) Then
        Set m.SentOrders(brokerMsg("GenesisID")) = brokerMsg
    Else
        m.SentOrders.Add brokerMsg, brokerMsg("GenesisID")
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCurrenex.UpdateSentOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshStartDate
'' Description: Determine the refresh start date
'' Inputs:      None
'' Returns:     Refresh Start Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RefreshStartDate() As String
On Error GoTo ErrSection:

    Dim lStartDate As Long              ' Starting date for the refresh
    
    lStartDate = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", m.Broker.IniFile))))
    If lStartDate = 0 Then
        lStartDate = Int(CurrentTime)
    End If
    
    Do While Weekday(lStartDate) <> vbSunday
        lStartDate = lStartDate - 1&
    Loop
    
    RefreshStartDate = Format(lStartDate, "YYYYMMDD")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCurrenex.RefreshStartDate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshEndDate
'' Description: Determine the refresh end date
'' Inputs:      None
'' Returns:     Refresh End Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RefreshEndDate() As String
On Error GoTo ErrSection:

    Dim lEndDate As Long                ' Ending date for the refresh
    
    lEndDate = Int(CurrentTime)
    Do While Weekday(lEndDate) <> vbFriday
        lEndDate = lEndDate + 1&
    Loop

    RefreshEndDate = Format(lEndDate, "YYYYMMDD")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCurrenex.RefreshEndDate"
    
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
    
    Set m.SentOrders = New cGdTree
    Set m.AmendOrders = New cGdTree
    Set m.CancelOrders = New cGdTree

    Set m.astrFillSymbols = New cGdArray
    m.astrFillSymbols.Create eGDARRAY_Strings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCurrenex.Class_Initialize"
    
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
    Set m.SentOrders = Nothing
    Set m.AmendOrders = Nothing
    Set m.CancelOrders = Nothing
    Set m.astrFillSymbols = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCurrenex.Class_Terminate"
    
End Sub
