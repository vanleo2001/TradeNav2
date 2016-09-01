VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerPats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerPats.cls
'' Description: Class to manange communications with PATS servers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 12/06/2011   DAJ         Added data calls, changed "Broker Name"
'' 12/13/2011   DAJ         Added Capital Trading Group for PATS
'' 01/30/2012   DAJ         Allow Reconnect flag on Connection Status message
'' 03/14/2012   DAJ         Added Alpari(PATS)
'' 07/16/2012   DAJ         ZanerPats
'' 07/18/2012   DAJ         RCG (New PATS)
'' 07/27/2012   DAJ         Demo (PATS)
'' 08/23/2012   DAJ         Born (PATS), RJO Hong Kong (PATS), Get symbols from PATS server
'' 08/28/2012   DAJ         CarriedFillFromInfo, First Field Begin/End
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 07/22/2014   DAJ         Enhancements for amending orders with PATS
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Private Const kAppID = "APPID_2284"
'Private Const kAppLicense = "BETA"
Private Const kAppID = "GFTTradeNav2.0"
Private Const kAppLicense = "KAD104V001"
Private Const kAppVersion = "3.05"

Private Type mPrivate
    Broker As cBroker                   ' Broker base class object
    BrokerKv As cBrokerKeyValue         ' Routines for key-value broker messages
    Symbols As cBrokerSymbols           ' Broker symbol conversion object
    
    strUserName As String               ' User name to login to the server with
    strPassword As String               ' Password for logging into the server
    strIP As String                     ' IP address for the server
    strPort As String                   ' Port for the server
    strPriceIP As String                ' IP address for the price server
    strPricePort As String              ' Port for the price server
    strEnvironment As String            ' Environment to log into
    bSuperTAS As Boolean                ' Use SuperTAS?
    
    strSubscribed As String             ' Symbol that is subscribed to
    
    AmendOrders As cGdTree              ' Collection of orders that have been sent for amend
    astrFillSymbols As cGdArray         ' Array of symbols for which we got fills
    astrEnabledSymbols As cGdArray      ' Array of Genesis symbols the user is enabled for
    GenesisIdMap As cGdTree             ' Map of new to old Genesis Order ID's
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

Public Property Get EnabledSymbols() As cGdArray
    Set EnabledSymbols = m.astrEnabledSymbols
End Property

Private Property Get AppID() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = GetIniFileProperty("ID", "", "Application", m.Broker.IniFile)
    If Len(strReturn) = 0 Then
        strReturn = kAppID
    End If
    
    AppID = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerPats.AppID.Get"

End Property

Private Property Get AppLicense() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = GetIniFileProperty("License", "", "Application", m.Broker.IniFile)
    If Len(strReturn) = 0 Then
        strReturn = kAppLicense
    End If
    
    AppLicense = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerPats.AppLicense.Get"

End Property

Private Property Get AppVersion() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = GetIniFileProperty("Version", "", "Application", m.Broker.IniFile)
    If Len(strReturn) = 0 Then
        strReturn = kAppVersion
    End If
    
    AppVersion = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerPats.AppVersion.Get"

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
        Case eTT_AccountType_PATS
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "Pats", True
            m.Broker.IniFile = AddSlash(App.Path) & "Pats.INI"
            m.Broker.BrokerName = "PATS"
            m.Broker.ProcessName = "GenPats"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\Pats\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenPats.EXE"
            m.Broker.ControlID = "Pats"
            strSymbolFile = AddSlash(App.Path) & "Provided\PatsToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\PatsCnct.INI"
            
        Case eTT_AccountType_AlpariPats
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "AlpariPats", True
            m.Broker.IniFile = AddSlash(App.Path) & "AlpariPats.INI"
            m.Broker.BrokerName = "Alpari (PATS)"
            m.Broker.ProcessName = "GenAlpariPats"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\AlpariPats\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenAlpariPats.EXE"
            m.Broker.ControlID = "AlpariPats"
            strSymbolFile = AddSlash(App.Path) & "Provided\AlpPatsToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\AlpPatsCnct.INI"
            
        Case eTT_AccountType_BornPats
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "BornPats", True
            m.Broker.IniFile = AddSlash(App.Path) & "BornPats.INI"
            m.Broker.BrokerName = "Born Capital (PATS)"
            m.Broker.ProcessName = "GenBornPats"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\BornPats\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenBornPats.EXE"
            m.Broker.ControlID = "BornPats"
            strSymbolFile = AddSlash(App.Path) & "Provided\BornPatsToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\BornPatsCnct.INI"
            
        Case eTT_AccountType_CtgPats
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "CtgPats", True
            m.Broker.IniFile = AddSlash(App.Path) & "CtgPats.INI"
            m.Broker.BrokerName = "CTG1"
            m.Broker.ProcessName = "GenCtgPats"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\CtgPats\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenCtgPats.EXE"
            m.Broker.ControlID = "CtgPats"
            strSymbolFile = AddSlash(App.Path) & "Provided\CtgPatsToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\CtgPatsCnct.INI"
            
        Case eTT_AccountType_DemoPats
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "DemoPats", True
            m.Broker.IniFile = AddSlash(App.Path) & "DemoPats.INI"
            m.Broker.BrokerName = "PATS (Demo)"
            m.Broker.ProcessName = "GenDemoPats"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\DemoPats\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenDemoPats.EXE"
            m.Broker.ControlID = "DemoPats"
            strSymbolFile = AddSlash(App.Path) & "Provided\DemoPatsToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\DemoPatsCnct.INI"
            
        Case eTT_AccountType_RcgPats
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "RcgPats", True
            m.Broker.IniFile = AddSlash(App.Path) & "RcgPats.INI"
            m.Broker.BrokerName = "RCG (PATS)"
            m.Broker.ProcessName = "GenRcgPats"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\RcgPats\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenRcgPats.EXE"
            m.Broker.ControlID = "RcgPats"
            strSymbolFile = AddSlash(App.Path) & "Provided\RcgPatsToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\RcgPatsCnct.INI"
                        
        Case eTT_AccountType_RjoHkPats
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "RjoHkPats", True
            m.Broker.IniFile = AddSlash(App.Path) & "RjoHkPats.INI"
            m.Broker.BrokerName = "RJO Hong Kong (PATS)"
            m.Broker.ProcessName = "GenRjoHkPats"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\RjoHkPats\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenRjoHkPats.EXE"
            m.Broker.ControlID = "RjoHkPats"
            strSymbolFile = AddSlash(App.Path) & "Provided\RjoHkPatsToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\RjoHkPatsCnct.INI"
            
        Case eTT_AccountType_RjoPats
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "RjoPats", True
            m.Broker.IniFile = AddSlash(App.Path) & "RjoPats.INI"
            m.Broker.BrokerName = "RJ O'Brien (PATS)"
            m.Broker.ProcessName = "GenRjoPats"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\RjoPats\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenRjoPats.EXE"
            m.Broker.ControlID = "RjoPats"
            strSymbolFile = AddSlash(App.Path) & "Provided\RjoPatsToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\RjoPatsCnct.INI"
            
        Case eTT_AccountType_ZanerPats
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "ZanerPats", True
            m.Broker.IniFile = AddSlash(App.Path) & "ZanerPats.INI"
            m.Broker.BrokerName = "Zaner (PATS)"
            m.Broker.ProcessName = "GenZanerPats"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\ZanerPats\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenZanerPats.EXE"
            m.Broker.ControlID = "ZanerPats"
            strSymbolFile = AddSlash(App.Path) & "Provided\ZanerPatsToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\ZanerPatsCnct.INI"
                        
    End Select

    m.Broker.CopyProcess AddSlash(App.Path) & "..\Brokers\GenPats2.EXE"
    FileCopy AddSlash(App.Path) & "..\Brokers\PatsAPI2.DLL", AddSlash(strBrokerPath) & "PatsAPI.DLL", True
    FileCopy AddSlash(App.Path) & "..\Brokers\SSLSocketLib.DLL", AddSlash(strBrokerPath) & "SSLSocketLib.DLL", True
    
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
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_BrokerView) = FileExist("C:\Common\Files.EXE")
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewOnline) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_VerifyPositions) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_AccountDetails) = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerPats.Init"

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

    ShowConfig = ((Len(m.strUserName) = 0) Or (Len(m.strPassword) = 0) Or (Len(m.strIP) = 0) Or (Len(m.strPort) = 0) Or (Len(m.strPriceIP) = 0) Or (Len(m.strPricePort) = 0) Or (Len(m.strEnvironment) = 0))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerPats.ShowConfig"
    
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
    RaiseError "cBrokerPats.ShowBrokerConnectionInfo"
    
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
    RaiseError "cBrokerPats.ShowLogin"
    
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
    RaiseError "cBrokerPats.SendConnect"
    
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
    RaiseError "cBrokerPats.SendDisconnect"
    
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
    RaiseError "cBrokerPats.UnloadApp"
    
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
                    astrFields(1) = BrokerSym.ContractToMMMYY(Parse(strGenesisSymbol, "-", 2))
                    astrFields(2) = BrokerSym.BrokerExchange
                    astrFields(3) = "FUT"
                    
                    strReturn = astrFields.JoinFields("|")
                End If
                
        End Select
    End If
    
    BrokerSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerPats.BrokerSymbol"
    
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
    Dim astrFields As cGdArray          ' Fields out of the symbol
    
    strReturn = ""
    If Len(strBrokerSymbol) > 0 Then
        Set astrFields = New cGdArray
        astrFields.SplitFields strBrokerSymbol, "|"
        
        If Len(astrFields(4)) = 0 Then
            Set BrokerSym = m.Symbols.GetByBrokerSymbol(astrFields(0), astrFields(2))
            If Not BrokerSym Is Nothing Then
                strReturn = BrokerSym.GenesisBase & "-" & BrokerSym.ContractFromMMMYY(astrFields(1))
            End If
        End If
    End If
    
    GenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerPats.GenesisSymbol"
    
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
    RaiseError "cBrokerPats.BrokerOrderType"
    
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
    RaiseError "cBrokerPats.AddOrder"
    
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
        If m.GenesisIdMap.Exists(NewOrder.GenesisOrderID) Then
            m.GenesisIdMap(NewOrder.GenesisOrderID) = Order.GenesisOrderID
            m.Broker.DumpDebug vbTab & "Map for '" & NewOrder.GenesisOrderID & "' updated: '" & Order.GenesisOrderID & "'"
        Else
            m.GenesisIdMap.Add Order.GenesisOrderID, NewOrder.GenesisOrderID
            m.Broker.DumpDebug vbTab & "Map for '" & NewOrder.GenesisOrderID & "' added: '" & Order.GenesisOrderID & "'"
        End If
        If m.AmendOrders.Exists(Order.GenesisOrderID) Then
            Set m.AmendOrders(Order.GenesisOrderID) = Order
            m.Broker.DumpDebug vbTab & Order.OrderText(True, True, True) & " updated in amend orders collection ( " & Order.GenesisOrderID & " )"
        Else
            m.AmendOrders.Add Order, Order.GenesisOrderID
            m.Broker.DumpDebug vbTab & Order.OrderText(True, True, True) & " added to amend orders collection ( " & Order.GenesisOrderID & " )"
        End If
        
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AmendOrder, strOrderString
        bReturn = True
    End If
    
    AmendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerPats.AmendOrder"
    
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
    RaiseError "cBrokerPats.CancelOrder"
    
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
    RaiseError "cBrokerPats.ParkOrder"
    
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
    RaiseError "cBrokerPats.Refresh"
    
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
    RaiseError "cBrokerPats.GetAccounts"
    
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
    RaiseError "cBrokerPats.GetOrders"
    
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
    RaiseError "cBrokerPats.GetFills"
    
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
    If Len(strAccount) = 0 Then
        brokerMsg.Add "Account", m.Broker.Accounts.JoinFields(";")
    Else
        brokerMsg.Add "Account", strAccount
    End If
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetPositions, brokerMsg.ToString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerPats.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSymbols
'' Description: Request symbols from the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetSymbols()
On Error GoTo ErrSection:

    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetSymbols, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerPats.GetSymbols"
    
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
    RaiseError "cBrokerPats.NextGenesisID"
    
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
    RaiseError "cBrokerPats.StartingGenesisIdForOptNav"
    
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
    RaiseError "cBrokerPats.IsTradeableSymbol"
    
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
    Dim strSecurityType As String       ' Security type for the symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = (Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0)
    If bReturn Then
        strBrokerBase = BrokerSym.BrokerBase
        strBrokerExchange = BrokerSym.BrokerExchange
        strSecurityType = SecurityType(strGenesisSymbol, True)
        
        Select Case UCase(strSecurityType)
            Case "F"
                bReturn = m.astrEnabledSymbols.BinarySearch(Parse(strGenesisSymbol, "-", 1))
            
            Case "FO"
                bReturn = m.astrEnabledSymbols.BinarySearch("O:" & Parse(strGenesisSymbol, "-", 1))
            
            Case Else
                bReturn = m.astrEnabledSymbols.BinarySearch(strGenesisSymbol)
                
        End Select
    End If
    
    IsEnabledSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerPats.IsEnabledSymbol"
    
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
    RaiseError "cBrokerPats.OrderTypeAllowed"
    
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
    RaiseError "cBrokerPats.TimeInForceAllowed"
    
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
            
        Case eGDBrokerMessageType_Symbols
            HandleSymbol strMessage
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerPats.HandleMessage"
    
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
    RaiseError "cBrokerPats.StatusFromInfo"
    
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
    RaiseError "cBrokerPats.AccountFromInfo"
    
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
            
            strGenesisSymbol = GenesisSymbol(brokerMsg.Symbol, brokerMsg("Exchange"), BrokerSym)
            strError = brokerMsg("Message")
            
            strPrevGenesisID = ""
            If m.GenesisIdMap.Exists(brokerMsg("GenesisID")) Then
                strPrevGenesisID = m.GenesisIdMap(brokerMsg("GenesisID"))
                m.Broker.DumpDebug vbTab & "Previous Genesis Order ID for '" & brokerMsg("GenesisID") & "' set to '" & strPrevGenesisID & "'"
            End If
            
            If Len(strPrevGenesisID) > 0 Then
                If (UCase(brokerMsg("Status")) = "REPLACED") Or (UCase(brokerMsg("Status")) = "NEW") Then
                    If m.AmendOrders.Exists(strPrevGenesisID) Then
                        Set AmendedOrder = m.AmendOrders(strPrevGenesisID)
                        
                        If Len(strError) > 0 Then
                            AmendedOrder.ChangeOrderStatus eTT_OrderStatus_Working
                            
                            InfBox strError, , , m.Broker.BrokerName & " Order Error"
                            strGenesisSymbol = ""
                        Else
                            ' 07/22/2014 DAJ: PATS orders seem to come back with the same BrokerID
                            ' after they are amended, so we probably want to continue to use the
                            ' same order instead of calling the old one amended.  I don't think that
                            ' we want to have multiple orders in the database with the same Broker ID...
                            'Set AmendedOrder = m.AmendOrders(strPrevGenesisID)
                            'AmendedOrder.ChangeOrderStatus eTT_OrderStatus_Amended
                        End If
                        
                        m.Broker.DumpDebug vbTab & AmendedOrder.OrderText(True, True, True) & " removed from amend orders collection ( " & strPrevGenesisID & " )"
                        m.AmendOrders.Remove strPrevGenesisID
                    End If
                End If
            Else
                Set AmendedOrder = Nothing
            End If
            
            If Len(strGenesisSymbol) > 0 Then
                Set Order = m.BrokerKv.OrderFromMessage(brokerMsg, m.Broker, strGenesisSymbol, BrokerSym)
                If AmendedOrder Is Nothing Then
                    Order.PreviousOrder = Nothing
                Else
                    Order.PreviousOrder = AmendedOrder
                End If
                
                bReturn = True
            End If
        End If
    End If
    
    OrderFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerPats.OrderFromInfo"
    
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
            
            strGenesisSymbol = GenesisSymbol(brokerMsg.Symbol, brokerMsg("Exchange"), BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                strKey = brokerMsg("Account") & vbTab & brokerMsg.Symbol
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
    RaiseError "cBrokerPats.FillFromInfo"
    
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
    RaiseError "cBrokerPats.CarriedFillFromInfo"
    
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
            GetSymbols
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
    RaiseError "cBrokerPats.PositionFromInfo"
    
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
    RaiseError "cBrokerPats.GenesisSymbolList"
    
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
    RaiseError "cBrokerPats.TimeZone"
    
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
    RaiseError "cBrokerPats.JustConnected"
    
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
    RaiseError "cBrokerPats.SymbolInformation"
    
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
    RaiseError "cBrokerPats.Subscribe"
    
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
    RaiseError "cBrokerPats.Unsubscribe"
    
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
    brokerMsg.Add "PriceIP", m.strPriceIP
    brokerMsg.Add "PricePort", m.strPricePort
    brokerMsg.Add "ApplicationID", AppID
    brokerMsg.Add "ApplicationLicense", AppLicense
    brokerMsg.Add "NewPassword", ""
    brokerMsg.Add "Environment", m.strEnvironment
    brokerMsg.Add "Version", AppVersion
    If m.bSuperTAS Then
        brokerMsg.Add "SuperTAS", "Y"
    Else
        brokerMsg.Add "SuperTAS", "N"
    End If
    
    ConnectString = brokerMsg.ToString(bForLog)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerPats.ConnectString"
    
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
    RaiseError "cBrokerPats.BrokerOrderString"
    
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

    bReturn = frmLoginPats.ShowMe(m.Broker, strUserName, bSwitching, bShowIP)
    If bReturn = True Then
        m.strUserName = frmLoginPats.UserName
        m.strPassword = frmLoginPats.Password
        m.strIP = frmLoginPats.IP
        m.strPort = frmLoginPats.Port
        m.strPriceIP = frmLoginPats.PriceIP
        m.strPricePort = frmLoginPats.PricePort
        m.strEnvironment = frmLoginPats.Environment
        m.bSuperTAS = frmLoginPats.SuperTAS
        
        m.Broker.AskedPassword = True
    End If
    
    GetLoginInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerPats.GetLoginInfo"
    
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
    RaiseError "cBrokerPats.HandlePriceUpdate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleSymbol
'' Description: Handle a symbol update from the PATS servers
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleSymbol(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    Dim brokerMessage As cBrokerMessage ' Broker message
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim lPos As Long                    ' Position in the array
    
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        If UCase(strFirstField) = "BEGIN" Then
            m.astrEnabledSymbols.Clear
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMessage = New cBrokerMessage
            brokerMessage.FromString strMessage
            
            Set BrokerSym = m.Symbols.GetByBrokerSymbol(brokerMessage("ContractName"), brokerMessage("ExchangeName"))
            If Not BrokerSym Is Nothing Then
                If m.astrEnabledSymbols.BinarySearch(BrokerSym.GenesisBase, lPos) = False Then
                    m.astrEnabledSymbols.Add BrokerSym.GenesisBase, lPos
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerPats.HandleSymbol"
    
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
    Set m.GenesisIdMap = New cGdTree
    
    Set m.astrFillSymbols = New cGdArray
    m.astrFillSymbols.Create eGDARRAY_Strings
    
    Set m.astrEnabledSymbols = New cGdArray
    m.astrEnabledSymbols.Create eGDARRAY_Strings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerPats.Class_Initialize"
    
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
    Set m.GenesisIdMap = Nothing
    Set m.astrFillSymbols = Nothing
    Set m.astrEnabledSymbols = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerPats.Class_Terminate"
    
End Sub
