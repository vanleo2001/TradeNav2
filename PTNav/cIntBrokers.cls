VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIntBrokers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cIntBrokers.cls
'' Description: Class that handles interaction with Interactive Brokers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 01/26/2009   DAJ         Did an overhaul to make it work with the new TWS
'' 01/28/2009   DAJ         Convert Genesis symbol to Trade Symbol first in IbSymbol
'' 02/02/2009   DAJ         Create dummy order if receive fill for non-existing order
'' 03/26/2009   DAJ         Use CommonFilesPath function instead of hard-coded
'' 06/09/2009   DAJ         Set the BrokerInfo connection status
'' 09/01/2009   DAJ         Use new Parked order status
'' 03/11/2010   DAJ         Use global collections
'' 03/07/2011   DAJ         Moved over to broker base class
'' 04/05/2011   DAJ         Added SymbolInformation call
'' 04/20/2011   DAJ         Added PreSubmitted order status
'' 04/29/2011   DAJ         Added Inactive order status
'' 05/02/2011   DAJ         When getting contract info, convert symbol to trade symbol
'' 05/04/2011   DAJ         Added NYBOT as new exchange for symbol conversion
'' 05/09/2011   DAJ         Fix for trading future options
'' 12/09/2011   DAJ         No longer prepend date on broker order ID
'' 02/02/2012   DAJ         Ask for different symbol than YI2 to check for SFE availability
'' 02/06/2012   DAJ         Change name for 'I-Deal' to 'TWS Pro'
'' 02/14/2012   DAJ         Added multi-leg order support
'' 04/17/2012   DAJ         Set Carried Position to Null so that cBroker will handle it
'' 07/13/2012   DAJ         Added symbol conversion routines for Forex, history calls
'' 08/03/2012   DAJ         Send history responses to the IB Data Pull form
'' 08/10/2012   DAJ         Send exchange, contract ID, and primary exchange on order
'' 08/28/2012   DAJ         CarriedFillFromInfo, First Field Begin/End
'' 09/26/2012   DAJ         Only set GenesisOrderID to NextID at submission time
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 01/18/2013   DAJ         Broker held OCO for Interactive Brokers
'' 05/06/2013   DAJ         Message message from Interactive Brokers stand-alone
''                          Position message for symbol we received fill but not position for
'' 07/24/2013   DAJ         Try not to ask for contract details for an expired contract
'' 10/15/2013   DAJ         Try not to ask for contract details for a spread chart symbol
'' 12/13/2013   DAJ         Fix for getting contract details for future options
'' 05/20/2014   DAJ         Made MessageTypeToString public; Created MessageTypeFromString
'' 05/30/2014   DAJ         Hold single leg order until we have contract information
'' 06/16/2014   DAJ         Added NYSELIFFE exchange to symbol converting functions
'' 07/15/2014   DAJ         Added LIFFE exchange to symbol converting functions
'' 08/13/2014   DAJ         Force Order ID >= 1000; Symbol Availability; Added CFE,SOFFEX to symbol conversions
'' 08/20/2014   DAJ         Added 'ICE' exchange when converting symbols
'' 08/28/2014   DAJ         Add IB Forex symbols into enabled symbols list so they get a broker override
'' 09/26/2014   DAJ         Keep asking for symbol availability on -1's for as much as seven days
'' 10/07/2014   DAJ         Set IsEnabled to true if get symbol real-time
'' 11/14/2014   DAJ         Added support for MIT, On-Close, and On-Open orders
'' 11/20/2014   DAJ         Added 'ICEEUSOFT' exchange to symbol converting functions
'' 11/24/2014   DAJ         Don't request symbol availability if we don't recognize the exchange
'' 11/24/2014   DAJ         Show message when done with symbol availability check
'' 12/22/2014   DAJ         Update main status bar while verifying symbols
'' 02/05/2015   DAJ         Added support for preffered/warrant/class stocks
'' 02/06/2014   DAJ         Added 'ICEEU' exchange to symbol converting functions
'' 05/13/2015   DAJ         Hand over maturity, p/c, and strike in GetContractInfo for stock options
'' 07/28/2015   DAJ         Allow real-time streaming for symbols that come back with 2121 error code
'' 08/24/2015   DAJ         Upon receiving an "open" order, ask for contract details if we don't have them
'' 09/30/2015   DAJ         Don't ask for symbol enablements on the weekend
'' 10/27/2015   DAJ         Pass TradingClass and Multiplier on GetSymbolAvailability and GetExchangeAvailability calls
'' 12/02/2015   DAJ         Added support for trading ASX stocks
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kRegKey = "Software\Genesis Financial Data Services\Navigator Suite\General"

Private Type mPrivate
    Broker As cBroker                   ' Broker base class object
    Symbols As cBrokerSymbols           ' Broker symbol conversion object

    strClientID As String               ' Client ID to log onto the Interactive Brokers servers
    strHostIP As String                 ' Host IP for the Interactive Brokers server
    strHostPort As String               ' Host port for the Interactive Brokers server

    strExchFile As String               ' Last known exchange file
    astrExchanges As cGdArray           ' Array of available exchanges
    strNewExchanges As String           ' New exchanges that the user is enabled for
    astrContracts As cGdArray           ' Array of contract information
    lNextID As Long                     ' Next available ID from Interactive Brokers
    astrFillSymbols As cGdArray         ' Array of symbols for which we got fills
    astrEnabledSymbols As cGdArray      ' Array of symbols that the user is enabled for market data through IB
    astrPrevEnabled As cGdArray         ' Array of previously enabled symbols
    astrGenesisSymbols As cGdArray      ' Array of Genesis symbols from the symbol list
    lSymbolIndex As Long                ' Index into the Genesis symbols array
    
    ContractIds As cGdTree              ' Collection of IB contract ID's for Genesis symbols
    HeldOrdersOut As cGdTree            ' Collection of outgoing orders being held until we have contract ID's
    HeldOrdersIn As cGdTree             ' Collection of incoming orders being held until we have contract ID's
    
    bRequestingSymbols As Boolean       ' Are we currently requesting symbols?
    dLastSymbolRequest As Double        ' Tick count of the last symbol request
    strLastSymbol As String             ' Last symbol requested
    lNumSymbolTries As Long             ' Number of times we have tried the last symbol request
    
    astrMarketClosed As cGdArray        ' List of symbols for which the market was closed during symbol request
    dLastMarketClosedRequest As Double  ' Tick count of the last market closed request
    strLastMarketClosed As String       ' Last symbol requested out of the market closed queue
    astrEnabledDates As cGdArray        ' Enabled symbols with date of last verification
End Type
Private m As mPrivate

Public Property Get Broker() As cBroker
    Set Broker = m.Broker
End Property

Public Property Get ClientID() As String
    ClientID = m.strClientID
End Property
Public Property Let ClientID(ByVal strClientID As String)
    m.strClientID = strClientID
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

Private Property Get ConnectString() As String
    ConnectString = m.strClientID & vbTab & m.strHostIP & vbTab & m.strHostPort
End Property

Public Property Get EnabledSymbols() As cGdArray
    If m.bRequestingSymbols Then
        Set EnabledSymbols = m.astrPrevEnabled
    Else
        Set EnabledSymbols = m.astrEnabledSymbols
    End If
End Property

Private Property Get NextID() As Long
    If m.lNextID < 1000 Then
        NextID = 1000
        m.lNextID = 1001
    Else
        NextID = m.lNextID
        m.lNextID = m.lNextID + 1&
    End If
End Property
Private Property Let NextID(ByVal lNextID As Long)
    m.lNextID = lNextID
End Property

Private Property Get LKSCRC() As Long
    LKSCRC = GetRegistryValue(rkLocalMachine, kRegKey, "LKSC" & Str(m.Broker.Broker), 0&)
End Property
Private Property Let LKSCRC(ByVal lLksCrc As Long)
    SetRegistryValue rkLocalMachine, kRegKey, "LKSC" & Str(m.Broker.Broker), lLksCrc, True
End Property

Private Property Get LKSDCRC() As Long
    LKSDCRC = GetRegistryValue(rkLocalMachine, kRegKey, "LKSDC" & Str(m.Broker.Broker), 0&)
End Property
Private Property Let LKSDCRC(ByVal lLksdCrc As Long)
    SetRegistryValue rkLocalMachine, kRegKey, "LKSDC" & Str(m.Broker.Broker), lLksdCrc, True
End Property

Private Property Get LksFile() As String
    LksFile = AddSlash(App.Path) & "Provided\LKS." & Str(m.Broker.Broker)
End Property

Private Property Get LksdFile() As String
    LksdFile = AddSlash(App.Path) & "Provided\LKSD." & Str(m.Broker.Broker)
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

    Select Case nBroker
        Case eTT_AccountType_Ideal
            m.Broker.Init Me, eTT_AccountType_Ideal, AddSlash(App.Path) & "IDeal"
            m.Broker.IniFile = AddSlash(App.Path) & "IDeal.INI"
            m.Broker.BrokerName = "TWS Pro" ' "I-Deal"
            m.Broker.ProcessName = "GenIdeal"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\IDeal\GenIdeal.EXE"
            m.Broker.ControlID = "Ideal"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\IbIps.INI"
            strSymbolFile = AddSlash(App.Path) & "Provided\IbToGen.TXT"
            m.strExchFile = AddSlash(App.Path) & "Provided\LkExch.IDL"
        
        Case eTT_AccountType_IntBrokers
            m.Broker.Init Me, eTT_AccountType_IntBrokers, AddSlash(App.Path) & "IntBrokers"
            m.Broker.IniFile = AddSlash(App.Path) & "IntBrokers2.INI"
            m.Broker.BrokerName = "Interactive Brokers"
            m.Broker.ProcessName = "GenInteractiveBrokers"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\InteractiveBrokers\GenInteractiveBrokers.EXE"
            m.Broker.ControlID = "InteractiveBrokers"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\IbIps.INI"
            strSymbolFile = AddSlash(App.Path) & "Provided\IbToGen.TXT"
            m.strExchFile = AddSlash(App.Path) & "Provided\LkExch.IB"
    
    End Select
    
    m.Broker.CopyProcess AddSlash(App.Path) & "..\Brokers\GenInteractiveBrokers.EXE"
    
    Set m.Symbols = New cBrokerSymbols
    m.Symbols.FromIbFile strSymbolFile

    Set m.astrExchanges = New cGdArray
    If FileExist(m.strExchFile) Then
        m.astrExchanges.Serialize m.strExchFile, False
    Else
        m.astrExchanges.Add "DTB"
        m.astrExchanges.Add "ECBOT"
        m.astrExchanges.Add "GLOBEX"
        m.astrExchanges.Add "NYMEX"
        
        m.astrExchanges.Serialize m.strExchFile, True
    End If

    Set m.astrContracts = New cGdArray
    m.astrContracts.Create eGDARRAY_Strings
    
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
    
    m.Broker.AllowBrokerOco = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.Init"

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

    ShowConfig = (GetIniFileProperty("Asked", 0&, "User", m.Broker.IniFile) = 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.ShowConfig"
    
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

    ShowBrokerConnectionInfo = frmIbCfg.ShowMe(m.Broker)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.ShowBrokerConnectionInfo"
    
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
    
    bReturn = True
    If Len(m.strClientID) = 0 Then
        bReturn = frmIbCfg.ShowMe(m.Broker)
    End If

    ShowLogin = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.ShowLogin"
    
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

    m.Broker.DumpDebug "Connecting to " & m.Broker.BrokerName & " as '" & m.strClientID & "'"
    SendIbMessage eGDIbMessageType_Connect, ConnectString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.SendConnect"
    
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

    SendIbMessage eGDIbMessageType_Disconnect, ""
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.SendDisconnect"
    
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

    SendIbMessage eGDIbMessageType_UnloadApp, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.UnloadApp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerSymbol
'' Description: Attempt to convert the Genesis symbol to the broker symbology
'' Inputs:      Genesis Symbol, IB Symbol Object, IB Security Type
'' Returns:     Broker Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerSymbol(ByVal strGenesisSymbol As String, Optional IbSymbol As cIbSymbol, Optional strIbSecType As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return from the function
    Dim strGenesisBase As String        ' Genesis Base symbol
    Dim strGenesisContract As String    ' Genesis Contract
    Dim OptSym As cOptionSymbol         ' Option symbol object
        
    Set IbSymbol = Nothing
    strReturn = ""
    If (Len(strGenesisSymbol) > 0) And (InStr(strGenesisSymbol, ",") = 0) Then
        Select Case SecurityType(strGenesisSymbol, True)
            Case "F"
                ' GLOBEX/CME:  6BH7
                ' ECBOT:       ZI   MAR 07
                ' DTB:         FSTX MAR 07
                strGenesisSymbol = ConvertToTradeSymbol(strGenesisSymbol, CurrentTime)
                strGenesisBase = Parse(strGenesisSymbol, "-", 1)
                strGenesisContract = Parse(strGenesisSymbol, "-", 2)
                
                Set IbSymbol = m.Symbols.GetByGenesisSymbol(strGenesisBase)
                If Not IbSymbol Is Nothing Then
                    strIbSecType = "FUT"
                    
                    Select Case UCase(IbSymbol.BrokerExchange)
                        Case "GLOBEX", "CME", "NYBOT", "NYMEX", "SNFE", "LIFFE", "CFE", "IPE", "ICEEU", "ICEEUSOFT"
                            strReturn = IbSymbol.BrokerBase & IbSymbol.ContractToMY(strGenesisContract)
                            
                        Case "ECBOT", "CBOT", "NYSELIFFE"
                            strReturn = IbSymbol.BrokerBase & "   " & IbSymbol.ContractToMMMYY(strGenesisContract)
                            
                        Case "DTB", "SOFFEX"
                            strReturn = IbSymbol.BrokerBase & " " & IbSymbol.ContractToMMMYY(strGenesisContract)
                                                    
                    End Select
                End If
            
            Case "S"
                If (InStr(strGenesisSymbol, "@") = 0) Then
                    Set IbSymbol = m.Symbols.GetByGenesisSymbol("!")
                ElseIf (InStr(strGenesisSymbol, "@ASX") <> 0) Then
                    Set IbSymbol = m.Symbols.GetByGenesisSymbol("!ASX")
                Else
                    Set IbSymbol = Nothing
                End If
                
                If Not IbSymbol Is Nothing Then
                    If InStr(strGenesisSymbol, "_P") <> 0 Then
                        IbSymbol.BrokerBase = Replace(strGenesisSymbol, "_P", " PR")
                        IbSymbol.IbBaseSymbol = Replace(strGenesisSymbol, "_P", " PR")
                    ElseIf InStr(strGenesisSymbol, "_") <> 0 Then
                        IbSymbol.BrokerBase = Replace(strGenesisSymbol, "_", " ")
                        IbSymbol.IbBaseSymbol = Replace(strGenesisSymbol, "_", " ")
                    Else
                        IbSymbol.BrokerBase = strGenesisSymbol
                        IbSymbol.IbBaseSymbol = strGenesisSymbol
                    End If
                    
                    strIbSecType = "STK"
                    strReturn = strGenesisSymbol
                End If
                
            Case "FO"
                ' CME: ESG1 C1330
                ' ECBOT: C OYM  MAR 11  12250
                Set OptSym = New cOptionSymbol
                OptSym.FromGenesis strGenesisSymbol
                Set IbSymbol = m.Symbols.GetByGenesisSymbol("O:" & OptSym.BaseSymbol)
                If Not IbSymbol Is Nothing Then
                    strIbSecType = "FOP"
                    strReturn = OptSym.ToIb(Mid(IbSymbol.BrokerBase, 3), IbSymbol.BrokerExchange, IbSymbol.StrikeMult, IbSymbol.StrikePriceFormat)
                End If
                
            Case "SO"
                ' F     110319P00017000
                Set OptSym = New cOptionSymbol
                OptSym.FromGenesis strGenesisSymbol
                Set IbSymbol = m.Symbols.GetByGenesisSymbol("@")
                If Not IbSymbol Is Nothing Then
                    IbSymbol.BrokerBase = OptSym.BaseSymbol
                    IbSymbol.IbBaseSymbol = OptSym.BaseSymbol
                    
                    strIbSecType = "OPT"
                    strReturn = OptSym.ToIb(OptSym.BaseSymbol, IbSymbol.BrokerExchange, IbSymbol.StrikeMult, IbSymbol.StrikePriceFormat)
                End If
                
            Case "I"
                If IsForex(strGenesisSymbol) Then
                    Set IbSymbol = m.Symbols.GetByGenesisSymbol(strGenesisSymbol)
                    If Not IbSymbol Is Nothing Then
                        strIbSecType = "CASH"
                        strReturn = IbSymbol.BrokerBase
                    End If
                End If
                
        End Select
    End If

    BrokerSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.BrokerSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Attempt to convert the broker symbol to the Genesis symbology
'' Inputs:      Broker Symbol, Broker Exchange, Security Type, IB Symbol Object
'' Returns:     Genesis Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbol(ByVal strBrokerSymbol As String, ByVal strBrokerExchange As String, ByVal strBrokerSecType As String, Optional IbSymbol As cIbSymbol, Optional ByVal strBrokerExpiry As String = "") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return from the function
    Dim strBrokerBase As String         ' Base symbol
    Dim OptSym As cOptionSymbol         ' Option symbol object
                
    strReturn = ""
    If Len(strBrokerSymbol) > 0 Then
        Select Case strBrokerSecType
            Case "FUT"
                ' GLOBEX/CME:  6BH7
                ' ECBOT:       ZI   MAR 07
                ' DTB:         FSTX MAR 07
                Select Case UCase(strBrokerExchange)
                    Case "GLOBEX", "CME", "NYBOT", "NYMEX", "SNFE", "LIFFE", "CFE", "IPE", "ICEEU", "ICEEUSOFT"
                        strBrokerBase = Left(strBrokerSymbol, Len(strBrokerSymbol) - 2)
                        Set IbSymbol = m.Symbols.GetByBrokerSymbol(strBrokerBase, strBrokerExchange)
                        If Not IbSymbol Is Nothing Then
                            If Len(strBrokerExpiry) > 0 Then
                                strReturn = IbSymbol.GenesisBase & "-" & strBrokerExpiry
                            Else
                                strReturn = IbSymbol.GenesisBase & "-" & IbSymbol.ContractFromMY(Right(strBrokerSymbol, 2))
                            End If
                        End If
                        
                    Case "ECBOT", "CBOT", "DTB", "NYSELIFFE", "SOFFEX"
                        strBrokerBase = Parse(strBrokerSymbol, " ", 1)
                        Set IbSymbol = m.Symbols.GetByBrokerSymbol(strBrokerBase, strBrokerExchange)
                        If Not IbSymbol Is Nothing Then
                            If Len(strBrokerExpiry) > 0 Then
                                strReturn = IbSymbol.GenesisBase & "-" & strBrokerExpiry
                            Else
                                strReturn = IbSymbol.GenesisBase & "-" & IbSymbol.ContractFromMMMYY(Parse(strBrokerSymbol, " ", 2) & " " & Parse(strBrokerSymbol, " ", 3))
                            End If
                        End If
                
                End Select
                
            Case "STK"
                If strBrokerExchange = "ASX" Then
                    Set IbSymbol = m.Symbols.GetByBrokerSymbol("!ASX", "SMART")
                Else
                    Set IbSymbol = m.Symbols.GetByBrokerSymbol("!", "SMART")
                End If
                If Not IbSymbol Is Nothing Then
                    If InStr(strBrokerSymbol, " PR") <> 0 Then
                        strReturn = Replace(strBrokerSymbol, " PR", "_P")
                    ElseIf InStr(strBrokerSymbol, " ") <> 0 Then
                        strReturn = Replace(strBrokerSymbol, " ", "_")
                    Else
                        strReturn = strBrokerSymbol
                    End If
                End If
                
            Case "FOP"
                ' CME: ESG1 C1330
                ' ECBOT: C OYM  MAR 11  12250
                Set OptSym = New cOptionSymbol
                OptSym.FromIb strBrokerSymbol, strBrokerExchange, True
                Set IbSymbol = m.Symbols.GetByBrokerSymbol("O:" & OptSym.BrokerBase, strBrokerExchange)
                If Not IbSymbol Is Nothing Then
                    OptSym.Strike = OptSym.BrokerStrike * IbSymbol.StrikeMult
                    OptSym.BaseSymbol = Mid(IbSymbol.GenesisBase, 3)
                    strReturn = OptSym.ToGenesis
                End If
            
            Case "OPT"
                ' F     110319P00017000
                Set OptSym = New cOptionSymbol
                OptSym.FromIb strBrokerSymbol, strBrokerExchange, False
                Set IbSymbol = m.Symbols.GetByBrokerSymbol("@", "SMART")
                If Not IbSymbol Is Nothing Then
                    OptSym.Strike = OptSym.BrokerStrike * IbSymbol.StrikeMult
                    strReturn = OptSym.ToGenesis
                End If
                
            Case "CASH"
                Set IbSymbol = m.Symbols.GetByBrokerSymbol(strBrokerSymbol, strBrokerExchange)
                If Not IbSymbol Is Nothing Then
                    If GetSymbolID(IbSymbol.GenesisBase) > 0 Then
                        strReturn = IbSymbol.GenesisBase
                    End If
                End If
                
        End Select
    End If

    GenesisSymbol = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.GenesisSymbol"
    
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

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    Select Case nGenesisOrderType
        Case eTT_OrderType_Market
            strReturn = "MKT"
        Case eTT_OrderType_Stop
            strReturn = "STP"
        Case eTT_OrderType_Limit
            strReturn = "LMT"
        Case eTT_OrderType_StopWithLimit
            strReturn = "STPLMT"
        Case eTT_OrderType_MarketOnClose
            strReturn = "MOC"
        Case eTT_OrderType_LimitCloseOnly
            strReturn = "LOC"
        Case eTT_OrderType_MIT
            strReturn = "MIT"
        Case eTT_OrderType_MarketOnOpen
            strReturn = "MOO"
    End Select
    
    BrokerOrderType = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.BrokerOrderType"
    
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
    Dim OtherOrder As cPtOrder          ' Other order in an OCO
    
    bReturn = False
    
    strOrderString = BrokerOrderString(Order, True)
    
    If (Len(strOrderString) > 0) Then
        If OrderReadyToSubmit(Order) Then
            ' 09/26/2012 DAJ: TWS seems to generate a duplicate order ID error even
            ' it the ID is unique but out of sequence ( we ran into this with an OTO
            ' order that got an ID but didn't submit until after some other orders).  To
            ' fix that, I will assign a "normal" GenesisOrderID and then change it to
            ' the next available ID when the order get submitted...
            m.Broker.BrokerInfo.ChangeGenesisIdForOrder Order, Str(NextID)
            
            If (Order.NumberOfLegs = 1) And (Len(Order.OcaGroup) = 0) Then
                If Order.BrokerCancelOrderID = 0 Then
                    Order.OcaGroup = OcaGroupForOrder(Order)
                Else
                    Set OtherOrder = New cPtOrder
                    If OtherOrder.Load(Abs(Order.BrokerCancelOrderID)) Then
                        Order.OcaGroup = OtherOrder.OcaGroup
                    Else
                        Order.OcaGroup = OcaGroupForOrder(Order)
                    End If
                End If
            End If
            
            strOrderString = BrokerOrderString(Order, True)
            
            If Order.NumberOfLegs = 1 Then
                SendIbMessage eGDIbMessageType_AddOrder, strOrderString
            ElseIf Order.NumberOfLegs > 1 Then
                SendIbMessage eGDIbMessageType_AddComboOrder, strOrderString
            End If
            bReturn = True
        Else
            m.HeldOrdersOut.Add Order
            RequestContractInfoForOrderLegs Order
        End If
    End If
    
    AddOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.AddOrder"
    
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
    
    If (Len(strOrderString) > 0) Then
        If Order.NumberOfLegs = 1 Then
            SendIbMessage eGDIbMessageType_AmendOrder, strOrderString
        ElseIf Order.NumberOfLegs > 1 Then
            SendIbMessage eGDIbMessageType_AmendComboOrder, strOrderString
        End If
        bReturn = True
    End If
    
    AmendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.AmendOrder"
    
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
    
    If (Len(strOrderString) > 0) Then
        SendIbMessage eGDIbMessageType_CancelOrder, strOrderString
        bReturn = True
    End If
    
    CancelOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.CancelOrder"
    
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
    
    If (Len(strOrderString) > 0) Then
        SendIbMessage eGDIbMessageType_CancelOrder, strOrderString
        bReturn = True
    End If
    
    ParkOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.ParkOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrdersAsOco
'' Description: Submit the given orders as an Order-Cancels-Order
'' Inputs:      Order1, Order2
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SubmitOrdersAsOco(Order1 As cPtOrder, Order2 As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strOrderString1 As String       ' Broker order string for Order 1
    Dim strOrderString2 As String       ' Broker order string for Order 2
    
    bReturn = False
    If (Order1.NumberOfLegs = 1) And (Order2.NumberOfLegs = 1) Then
        m.Broker.BrokerInfo.ChangeGenesisIdForOrder Order1, Str(NextID)
        Order1.OcaGroup = OcaGroupForOrder(Order1)
        
        m.Broker.BrokerInfo.ChangeGenesisIdForOrder Order2, Str(NextID)
        Order2.OcaGroup = Order1.OcaGroup
        
        strOrderString1 = BrokerOrderString(Order1, True)
        strOrderString2 = BrokerOrderString(Order2, True)
        
        If (Len(strOrderString1) > 0) And (Len(strOrderString2) > 0) Then
            SendIbMessage eGDIbMessageType_AddOrder, strOrderString1
            Order1.ChangeOrderStatus eTT_OrderStatus_Sent, m.Broker.BrokerInfo
            
            SendIbMessage eGDIbMessageType_AddOrder, strOrderString2
            Order2.ChangeOrderStatus eTT_OrderStatus_Sent, m.Broker.BrokerInfo
            
            bReturn = True
        End If
    End If

    SubmitOrdersAsOco = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.SubmitOrdersAsOco"
    
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
    RaiseError "cIntBrokers.Refresh"
    
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

    SendIbMessage eGDIbMessageType_GetAccounts, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.GetAccounts"
    
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

    SendIbMessage eGDIbMessageType_GetOrders, strAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.GetOrders"
    
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

    SendIbMessage eGDIbMessageType_GetFills, strAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.GetFills"
    
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

    SendIbMessage eGDIbMessageType_GetPositions, strAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.GetPositions"
    
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

' 09/26/2012 DAJ: TWS seems to generate a duplicate order ID error even
' it the ID is unique but out of sequence ( we ran into this with an OTO
' order that got an ID but didn't submit until after some other orders).  To
' fix that, I will assign a "normal" GenesisOrderID and then change it to
' the next available ID when the order get submitted...
#If 0 Then
    NextGenesisID = Str(NextID)
#Else
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
#End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.NextGenesisID"
    
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
    RaiseError "cIntBrokers.StartingGenesisIdForOptNav"
    
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
    RaiseError "cIntBrokers.IsTradeableSymbol"
    
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
    Dim BrokerSym As cIbSymbol          ' Broker symbol object
    Dim lPos As Long                    ' Position in the array
    Dim strIbSecType As String          ' IB Security type
   
    bReturn = (Len(BrokerSymbol(strGenesisSymbol, BrokerSym, strIbSecType)) > 0)
    If bReturn Then
        strBrokerBase = BrokerSym.BrokerBase
        strBrokerExchange = BrokerSym.BrokerExchange
        
        ' 10/07/2014 DAJ: If the customer is getting the symbol real-time through any means, let
        ' them trade it through Interactive Brokers...
        If g.RealTime.SymbolDelay(strGenesisSymbol) = 0 Then
            bReturn = True
        Else
            Select Case UCase(strIbSecType)
                Case "FUT", "FOP"
                    bReturn = EnabledSymbols.BinarySearch(Parse(strGenesisSymbol, "-", 1))
                    
                Case "STK"
                    bReturn = HasModule("E_US")
                
                Case "OPT"
                    bReturn = HasModule("E_OPRA")
                
                Case "CASH"
                    bReturn = True
                
            End Select
        End If
    End If
    
    IsEnabledSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.IsEnabledSymbol"
    
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
    Dim IbSym As cIbSymbol              ' Broker symbol object
    
    bReturn = False
    If Len(BrokerSymbol(strGenesisSymbol, IbSym)) > 0 Then
        bReturn = IbSym.OrderTypeAllowed(nOrderType)
    End If
    
    OrderTypeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.OrderTypeAllowed"
    
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
    Dim IbSym As cIbSymbol              ' Broker symbol object
    
    bReturn = False
    If Len(BrokerSymbol(strGenesisSymbol, IbSym)) > 0 Then
        bReturn = IbSym.TifAllowed(nTimeInForce)
    End If
    
    TimeInForceAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.TimeInForceAllowed"
    
End Function

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

    Dim lDate As Long                   ' Date portion of the date/time
    Dim dTime As Double                 ' Time portion of the date/time
    Dim strTime As String               ' Time portion of the date/time
    Dim lHour As Long                   ' Hour portion of the time
    Dim lMin As Long                    ' Minute portion of the time
    Dim dSec As Double                  ' Second portion of the time
    
    If InStr(strDateTime, "-") <> 0 Then
        lDate = CLng(Val(Parse(strDateTime, "-", 1)))
        strTime = Parse(strDateTime, "-", 2)
    Else
        lDate = CLng(Val(Parse(strDateTime, " ", 1)))
        strTime = Parse(strDateTime, " ", 2)
    End If
    
    lDate = JulFromLong(lDate)
    lHour = CLng(Val(Parse(strTime, ":", 1)))
    lMin = CLng(Val(Parse(strTime, ":", 2)))
    dSec = CLng(Val(Parse(strTime, ":", 3)))
    dTime = (lHour / 24#) + (lMin / 1440#) + (dSec / 86400#)
    
    DateToDouble = CDbl(lDate) + dTime

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.DateToDouble"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendIbMessage
'' Description: Send an App Mail message to the Interactive Brokers program
'' Inputs:      Type of Message, Message, Send Now?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendIbMessage(ByVal nType As eGDIbMessageTypes, ByVal strMessage As String, Optional ByVal bSendNow As Boolean = True)
On Error GoTo ErrSection:

    m.Broker.DumpDebug vbTab & "Sending " & m.Broker.BrokerName & " Message (" & MessageTypeToString(nType) & "): " & strMessage
    frmOnlineBroker.gdBroker.CreateMessage m.Broker.ControlID, nType, strMessage, , bSendNow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.SendIbMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming App Mail message from the stand-alone program
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDIbMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    If Not m.Broker Is Nothing Then
        If (nType <> eGDIbMessageType_ConnectionInfo) And (nType <> eGDIbMessageType_Heartbeat) And (nType <> eGDIbMessageType_History) Then
            m.Broker.DumpDebug vbTab & m.Broker.BrokerName & " Message Received (" & MessageTypeToString(nType) & "): " & strMessage
        End If
        
        Select Case nType
            Case eGDIbMessageType_ConnectionInfo
                If Len(Parse(strMessage, vbTab, 2)) > 0 Then
                    m.Broker.HandleConnectionInfo CLng(Val(Parse(strMessage, vbTab, 2))), Parse(strMessage, vbTab, 3), m.strClientID
                End If
                
            Case eGDIbMessageType_AppLoaded
                m.Broker.AppLoaded = True
                
            Case eGDIbMessageType_AppUnloaded
                m.Broker.AppLoaded = False
                
            Case eGDIbMessageType_Heartbeat
                m.Broker.HeartBeat = gdTickCount
                
            Case eGDIbMessageType_AccountR
                m.Broker.AccountReceived strMessage
                
            Case eGDIbMessageType_Order
                m.Broker.OrderReceived strMessage, False
            
            Case eGDIbMessageType_OrderR
                m.Broker.OrderReceived strMessage, True
                
            Case eGDIbMessageType_Fill
                m.Broker.FillReceived strMessage, False
            
            Case eGDIbMessageType_FillR
                m.Broker.FillReceived strMessage, True
                
            Case eGDIbMessageType_Position
                m.Broker.PositionReceived strMessage, False
            
            Case eGDIbMessageType_PositionR
                m.Broker.PositionReceived strMessage, True
                
            Case eGDIbMessageType_ExchangeAvailable
                m.Broker.HeartBeat = gdTickCount
                HandleExchangeInfo strMessage
                
            Case eGDIbMessageType_NextValidID
                m.Broker.HeartBeat = gdTickCount
                NextID = CLng(Val(strMessage))
                
            Case eGDIbMessageType_ContractDetails
                m.Broker.HeartBeat = gdTickCount
                HandleContractInfo strMessage
                
            Case eGDIbMessageType_History
                m.Broker.HeartBeat = gdTickCount
                If FormIsLoaded("frmIbDataPull") Then
                    frmIbDataPull.IB_Data strMessage
                End If
                
            Case eGDIbMessageType_Message
                m.Broker.HeartBeat = gdTickCount
                ShowMessage strMessage
                
            Case eGDIbMessageType_SymbolAvailable
                m.Broker.HeartBeat = gdTickCount
                HandleSymbolAvailability strMessage
    
        End Select
    End If
    
    If m.dLastSymbolRequest > 0# Then
        If gdTickCount > m.dLastSymbolRequest + 15000 Then
            If m.lNumSymbolTries > 5& Then
                If GetNextSymbol = False Then
                    FinishedLastSymbol
                End If
            Else
                GetSymbolAvailability m.strLastSymbol
                
                SetLastSymbolRequest gdTickCount, "HandleMessage"
                m.lNumSymbolTries = m.lNumSymbolTries + 1&
            End If
        End If
    End If
    
    ' If we are done requesting the normal symbols and there are symbols in the market
    ' closed queue, then start requesting those symbols every five minutes...
    If (m.bRequestingSymbols = False) And (m.astrMarketClosed.Size > 0) Then
        ' If the last one of these requests took longer than 15 seconds, add it back to
        ' the end of the queue...
        If (m.dLastSymbolRequest > 0#) And (gdTickCount > m.dLastSymbolRequest + 15000) Then
            m.dLastSymbolRequest = 0#
            m.astrMarketClosed.Add m.strLastMarketClosed
            m.Broker.DumpDebug vbTab & vbTab & m.strLastMarketClosed & " added to Market Closed queue because request timed out"
        End If
        
        ' If it is the first request or it has been five minutes since the last request, then
        ' request the next symbol in the queue...
        If (m.dLastMarketClosedRequest = 0#) Or (gdTickCount > (m.dLastMarketClosedRequest + (5 * 60000))) Then
            m.strLastMarketClosed = m.astrMarketClosed(0)
            m.astrMarketClosed.Remove 0
            m.Broker.DumpDebug vbTab & vbTab & m.strLastMarketClosed & " removed from Market Closed queue"
            
            m.dLastMarketClosedRequest = gdTickCount
            
            frmTest2.AddList m.strLastMarketClosed & " requested from Market Closed queue"
            GetSymbolAvailability m.strLastMarketClosed
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.HandleMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountFromInfo
'' Description: Fill an account object from the given message
'' Inputs:      Message, Account
'' Returns:     True if OK, False otherwise
''
'' Fields:      Account Number
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountFromInfo(ByVal strMessage As String, Acct As cPtAccount) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            With Acct
                .AccountNumber = Parse(strMessage, vbTab, 1)
                .AccountType = m.Broker.Broker
                .Broker = ""
                .ClosedProfit = 0#
                .Comms = 0#
                .CurrentBalance = 0#
                .FillRT = False
                .HoldOcoAtBroker = False
                .Name = .AccountNumber
                .Password = ""
                .SecTypeMask = 27
                .StartingBalance = 0#
                .StartingDate = Date
                .UserName = m.strClientID
            End With
            
            bReturn = True
        End If
    End If
    
    AccountFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.AccountFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromInfo
'' Description: Fill an order object from the given message
'' Inputs:      Message, Order
'' Returns:     True if OK, False otherwise
''
'' Fields:      Genesis ID, Broker ID, Account, IB Base Symbol, Symbol, Contract,
''              Strike, P/C, Exchange, Security Type, Currency, Order Type,
''              Buy/Sell, Quantity, Limit Price, Stop Price, Expire Time, TIF,
''              Remaining Quantity, Order Status, Message, Refresh, Leg Info,
''              Oca Group
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderFromInfo(ByVal strMessage As String, Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrOrder As cGdArray           ' Order information split out from the message
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim IbSym As cIbSymbol              ' Broker symbol object
    Dim lQuantityFilled As Long         ' Quantity of the order that has filled
    Dim Bars As New cGdBars             ' Bars object
    Dim dStatusDate As Double           ' Status date
    Dim strFirstField As String         ' First field in the string
    Dim lIndex As Long                  ' Index into a for loop

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set astrOrder = New cGdArray
            astrOrder.SplitFields strMessage, vbTab
            
            If UCase(astrOrder(9)) = "BAG" Then
                If UCase(astrOrder(8)) = "SMART" Then
                    strGenesisSymbol = GenesisSymbol(astrOrder(3), astrOrder(8), "STK", IbSym)
                Else
                    strGenesisSymbol = GenesisSymbol(astrOrder(3), astrOrder(8), "FUT", IbSym)
                End If
            Else
                strGenesisSymbol = GenesisSymbol(astrOrder(4), astrOrder(8), astrOrder(9), IbSym)
            End If
            
            If (Len(strGenesisSymbol) > 0) Or (Len(astrOrder(22)) > 0) Then
                SetBarProperties Bars, strGenesisSymbol
                
                dStatusDate = ConvertToBrokerDate(CurrentTime, m.Broker.Broker, strGenesisSymbol, True)
                
                Set Order = New cPtOrder
                With Order
                    .AccountID = g.Broker.AccountIDForNumber(astrOrder(2))
                    .BrokerID = astrOrder(1)
                    .Buy = ((UCase(astrOrder(12)) = "BUY") Or (UCase(astrOrder(12)) = "BOT"))
                    .GenesisOrderID = astrOrder(0)
                    .LimitPrice = GenesisPrice(astrOrder(14), IbSym)
                    If UCase(astrOrder(20)) <> "ORDER CANCELED - REASON:" Then
                        .Message = astrOrder(20)
                    Else
                        .Message = ""
                    End If
                    .OcaGroup = astrOrder(23)
                    .OrderType = GenesisOrderType(astrOrder(11))
                    .Quantity = CLng(Val(astrOrder(13)))
                    .Status = GenesisOrderStatus(astrOrder(19))
                    .StatusDate = dStatusDate
                    .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.StatusDate, m.Broker.Broker, strGenesisSymbol, False))
                    .StopPrice = GenesisPrice(astrOrder(15), IbSym)
                    .SymbolOrSymbolID = strGenesisSymbol
                    
                    If .OrderType = eTT_OrderType_MIT Then
                        .LimitPrice = .StopPrice
                        .StopPrice = 0#
                    End If
                    
                    Select Case UCase(astrOrder(17))
                        Case "DAY"
                            .Expiration = .SessionDate * -1&
                        Case "GTC"
                            .Expiration = 0
                        Case "GTD"
                            .Expiration = JulFromLong(CLng(Val(Parse(astrOrder(16), " ", 1))))
                    End Select
                                        
                    ' When you submit a Stop or Stop with Limit order, Interactive Brokers calls it PreSubmitted
                    ' because they hold on to it until the stop price is hit and then submit the appropriate
                    ' Market or Limit order from there, so even though they are calling it PreSubmitted, we will
                    ' assume that it is working for our purposes...
                    If (.Status = eTT_OrderStatus_Queued) Then
                        If ((.OrderType = eTT_OrderType_Stop) Or (.OrderType = eTT_OrderType_StopWithLimit)) Then
                            .Status = eTT_OrderStatus_Working
                        End If
                    End If
                End With
                
                bReturn = LegsFromInfo(astrOrder(22), Order)
                If bReturn = False Then
                    If m.HeldOrdersIn.Exists(Order.BrokerID) Then
                        m.HeldOrdersIn(Order.BrokerID) = strMessage
                    Else
                        m.HeldOrdersIn.Add strMessage, Order.BrokerID
                    End If
                End If
                
                ' If the order is a working order, make sure that we have contract information for the symbol on each
                ' of the legs in case the user later amends the order...
                If bReturn = True Then
                    If IsOpenOrder(Order.Status) = True Then
                        For lIndex = 1 To Order.NumberOfLegs
                            If m.ContractIds.Exists(Order.OrderLegs(lIndex).Symbol) = False Then
                                GetContractInfo Order.OrderLegs(lIndex).Symbol
                            End If
                        Next lIndex
                    End If
                End If
            End If
        End If
    End If
    
    OrderFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.OrderFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillFromInfo
'' Description: Fill a fill object from the given message
'' Inputs:      Message, Fill, Genesis Order ID, Carried?
'' Returns:     True if OK, False otherwise
''
'' Fields:      Genesis ID, Broker ID, Account, IB Base Symbol, Symbol, Contract,
''              Strike, P/C, Exchange, Security Type, Currency, Fill ID, Side,
''              Shares, Price, Time, Cum Quantity, Refresh
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillFromInfo(ByVal strMessage As String, Fill As cPtFill, Optional strGenesisOrderID As String = "", Optional ByVal bCarried As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrFill As cGdArray            ' Fill information split out from the message
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim IbSym As cIbSymbol              ' Broker symbol object
    Dim Bars As New cGdBars             ' Bars object
    Dim dFillTime As Double             ' Fill Time
    Dim strFirstField As String         ' First field in the string
    Dim lPos As Long                    ' Position of the symbol in the array
    Dim strKey As String                ' Key into the fill symbols array

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
            m.astrFillSymbols.Clear
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set astrFill = New cGdArray
            astrFill.SplitFields strMessage, vbTab
            
            strGenesisSymbol = GenesisSymbol(astrFill(4), astrFill(8), astrFill(9), IbSym)
            If Len(strGenesisSymbol) > 0 Then
                SetBarProperties Bars, strGenesisSymbol
                
                strKey = astrFill(2) & vbTab & strGenesisSymbol
                If m.astrFillSymbols.BinarySearch(strKey, lPos) = False Then
                    m.astrFillSymbols.Add strKey, lPos
                    m.Broker.DumpDebug vbTab & vbTab & "Added to Fill Symbols: '" & strKey & "'"
                End If
                
                If Len(astrFill(15)) > 0 Then
                    dFillTime = DateToDouble(astrFill(15))
                Else
                    dFillTime = ConvertToBrokerDate(CurrentTime, m.Broker.Broker, strGenesisSymbol, True)
                End If
                
                Set Fill = New cPtFill
                With Fill
                    .AccountID = g.Broker.AccountIDForNumber(astrFill(2))
                    .BrokerID = astrFill(11)
                    .BrokerOrderID = astrFill(1) '  Format(dFillTime, "YYYYMMDD") & "-" & astrFill(1)
                    .Buy = ((UCase(astrFill(12)) = "BUY") Or (UCase(astrFill(12)) = "BOT"))
                    .FillDate = dFillTime
                    .Price = GenesisPrice(astrFill(14), IbSym)
                    .Quantity = CLng(Val(astrFill(13)))
                    .SessionDate = Bars.SessionDateForTradeTime(.FillDateConverted(""))
                    .SymbolOrSymbolID = strGenesisSymbol
                End With
                
                strGenesisOrderID = astrFill(0)
                
                bReturn = True
            End If
        End If
    End If
    
    FillFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.FillFromInfo"
    
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
    RaiseError "cIntBrokers.CarriedFillFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionFromInfo
'' Description: Fill a position object from the given message
'' Inputs:      Message, Position
'' Returns:     True if OK, False otherwise
''
'' Fields:      Account, IB Base Symbol, Symbol, Contract, Strike, P/C, Exchange,
''              Security Type, Currency, Position, Realized, Unrealized,
''              Avg Entry
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionFromInfo(ByVal strMessage As String, Position As cPtPosition) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrPosition As cGdArray        ' Position information split out from the message
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim IbSym As cIbSymbol              ' Broker symbol object
    Dim strFirstField As String         ' First field in the string
    Dim lPos As Long                    ' Position of the symbol in the array
    Dim lIndex As Long                  ' Index into a for loop
    Dim strKey As String                ' Key into the fill symbols array
    Dim strPositionMessage As String    ' Position message

    bReturn = False
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
            Set astrPosition = New cGdArray
            astrPosition.SplitFields strMessage, vbTab
            
            strGenesisSymbol = GenesisSymbol(astrPosition(2), astrPosition(6), astrPosition(7), IbSym)
            If Len(strGenesisSymbol) > 0 Then
                strKey = astrPosition(0) & vbTab & strGenesisSymbol
                If m.astrFillSymbols.BinarySearch(strKey, lPos) = True Then
                    m.astrFillSymbols.Remove lPos
                    m.Broker.DumpDebug vbTab & vbTab & "Removed from Fill Symbols: '" & strKey & "'"
                End If
                
                Set Position = New cPtPosition
                With Position
                    .AccountID = g.Broker.AccountIDForNumber(astrPosition(0))
                    .SymbolOrSymbolID = strGenesisSymbol
                    .CurrentPosition = CLng(Val(astrPosition(9)))
                    .CarriedPosition = kNullData
                    
                    ' DAJ 04/17/2012 - The average price coming back from Interactvie Brokers is
                    ' actually the "Average Cost" (which includes commission), not the "Average
                    ' Entry Price", so it is not the correct number for this purpose...
                    ''.AverageEntry = m.Broker.GenesisPrice(astrPosition(12), IbSym.PriceMult)
                End With
                
                bReturn = True
            End If
        End If
    End If
    
    PositionFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.PositionFromInfo"
    
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
    
    Dim strFileName As String           ' Filename for the LKSyms file
    Dim lLksCrc As Long                 ' CRC for the LKSyms file
    Dim astrSymbols As cGdArray         ' Array of symbols
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position to insert into the array
    
    If bOnlyIfStream = True Then
        If m.astrEnabledSymbols.Size = 0 Then
            strFileName = LksFile
            lLksCrc = LKSCRC
        
            If CalcFileCrc(strFileName) <> lLksCrc Then
                KillFile strFileName
            End If
        
            If m.astrEnabledSymbols.Serialize(strFileName, False) = False Then
                ' If there is no CRC entry in the registry and they have connected before, give
                ' them all of the symbols like we have been doing...
                If (lLksCrc = 0&) And (m.Broker.LastDateConnected > 0) Then
                    Set m.astrEnabledSymbols = m.Symbols.GenesisSymbolList(bOnlyIfStream)
                
                ' Otherwise, if the CRC's don't match, or they don't have a LKS file, give them
                ' a small default list...
                Else
                    m.astrEnabledSymbols.SplitFields "G6E,ES,GF,LE,NQ,YM,ZB,ZC,ZN", ","
                End If
                
                m.astrEnabledSymbols.Sort
                m.astrEnabledSymbols.Serialize strFileName, True
                LKSCRC = CalcFileCrc(strFileName)
            Else
                m.astrEnabledSymbols.Sort
            End If
        End If
        
        strFileName = LksdFile
        lLksCrc = LKSDCRC
        
        If CalcFileCrc(strFileName) <> lLksCrc Then
            KillFile strFileName
        End If
        
        If m.astrEnabledDates.Serialize(strFileName, False) = False Then
            For lIndex = 0 To m.astrEnabledSymbols.Size - 1
                m.astrEnabledDates.Add m.astrEnabledSymbols(lIndex) & vbTab & Str(Int(CurrentTime) - 1)
            Next lIndex
            
            m.astrEnabledDates.Serialize strFileName, True
        End If
        
        ' DAJ 08/28/2014: We need to make sure that the IB Forex symbols get added to the
        ' enabled symbols list so that we ask for a broker override from streaming...
        Set astrSymbols = m.Symbols.GenesisSymbolList(False)
        For lIndex = 0 To astrSymbols.Size - 1
            If InStr(astrSymbols(lIndex), "@IB") > 0 Then
                If m.astrEnabledSymbols.BinarySearch(astrSymbols(lIndex), lPos) = False Then
                    m.astrEnabledSymbols.Add astrSymbols(lIndex), lPos
                End If
            End If
        Next lIndex
        
        Set GenesisSymbolList = m.astrEnabledSymbols
    Else
        Set GenesisSymbolList = m.Symbols.GenesisSymbolList(bOnlyIfStream)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.GenesisSymbolList"
    
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

    TimeZone = ""

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.TimeZone"
    
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

    'Dim strCheckSymbol As String        ' Symbol to use to check for availability
    'Dim strTradeSymbol As String        ' Trade symbol equivalent to check symbol
    
    Set m.astrGenesisSymbols = GenesisSymbolList(False)
    
    ' DAJ 09/29/2015: It looks like if we ask for data enablements over the weekend ( from
    ' Friday night to Sunday some time ), we don't get data for anything and, therefore, strip
    ' all of the future enablements.  So, don't ask over the weekend...
    If AskForEnablements = True Then
        Set m.astrPrevEnabled = m.astrEnabledSymbols.MakeCopy
        m.astrExchanges.Clear
        m.bRequestingSymbols = True
        m.lSymbolIndex = -1
        GetNextSymbol
    End If
    
    ' Ask for exchange availability for the SFE by asking for a symbol off of that
    ' exchange as long as it is in the symbol pool to begin with...
    'strCheckSymbol = GetIniFileProperty("Symbol", "AP-067", "SFE", m.Broker.ConnectIni)
    'strTradeSymbol = ConvertToTradeSymbol(strCheckSymbol, Date)
    'If strTradeSymbol <> strCheckSymbol Then
    '    GetExchangeAvailability strTradeSymbol
    'End If
    
    ' Ask for exchange availability for the LIFFE by asking for a symbol off of that
    ' exchange as long as it is in the symbol pool to begin with...
    'strCheckSymbol = GetIniFileProperty("Symbol", "ZX-067", "LIFFE", m.Broker.ConnectIni)
    'strTradeSymbol = ConvertToTradeSymbol(strCheckSymbol, Date)
    'If strTradeSymbol <> strCheckSymbol Then
    '    GetExchangeAvailability strTradeSymbol
    'End If
    
    m.astrContracts.Clear
    GetContracts True
    GetAccounts

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.JustConnected"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetContracts
'' Description: Ask the server for the available contracts for this user
'' Inputs:      Only Symbols on Forms?
'' Returns:     None
''
'' Fields:      Symbol, Long Symbol, Maturity, Exchange, Security Type,
''              Currency, Put/Call, Strike
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetContracts(Optional ByVal bOnlySymbolsOnForms As Boolean = False)
On Error GoTo ErrSection:

    Dim astrToSend As cGdArray          ' Array of information to send
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrSymbols As cGdArray         ' Array of symbols on forms
    Dim IbSym As cIbSymbol              ' IB Symbol object
    
    If bOnlySymbolsOnForms Then
        m.Broker.DumpDebug "Getting Contracts for Symbols on Forms"
        
        Set astrSymbols = g.Broker.SymbolsOnFormsForBroker(m.Broker.Broker)
        For lIndex = 0 To astrSymbols.Size - 1
            GetContractInfo astrSymbols(lIndex)
        Next lIndex
    Else
        m.Broker.DumpDebug "Getting Contracts"
        
        Set astrToSend = New cGdArray
        astrToSend.Create eGDARRAY_Strings
        
        For lIndex = 1 To m.Symbols.Count
            Set IbSym = m.Symbols(lIndex)
            If Not IbSym Is Nothing Then
                ' Only ask for futures...
                If (IsAlpha(Left(IbSym.GenesisBase, 1)) = True) And (InStr(IbSym.GenesisBase, "O:") = 0) Then
                    astrToSend.Clear
                    
                    astrToSend(0) = IbSym.IbBaseSymbol
                    astrToSend(1) = ""
                    astrToSend(2) = ""
                    astrToSend(3) = IbSym.BrokerExchange
                    astrToSend(4) = "FUT"
                    astrToSend(5) = ""
                    astrToSend(6) = ""
                    astrToSend(7) = ""
                
                    SendIbMessage eGDIbMessageType_GetContracts, astrToSend.JoinFields(vbTab)
                End If
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.GetContracts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetContractInfo
'' Description: Ask the server for the contract information for the given symbol
'' Inputs:      Symbol, Show Message?
'' Returns:     None
''
'' Fields:      Symbol, Local Symbol, Maturity, Exchange, Security Type,
''              Currency, Put/Call, Strike
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetContractInfo(ByVal strGenesisSymbol As String, Optional ByVal bShowMessage As Boolean = False)
On Error GoTo ErrSection:

    Dim astrToSend As cGdArray          ' Array of information to send
    Dim strIbSymbol As String           ' Interactive Brokers symbol
    Dim strIbSecType As String          ' Interactive Brokers security type
    Dim lPos As Long                    ' Position in the array
    Dim IbSym As cIbSymbol              ' IB Symbol object
    Dim strTradeSymbol As String        ' Trade symbol for the Genesis symbol passed in
    
    strTradeSymbol = ConvertToTradeSymbol(strGenesisSymbol, CurrentTime)
    If m.astrContracts.BinarySearch(strTradeSymbol & "|", lPos, eGdSort_MatchUsingSearchStringLength) Then
        m.Broker.DumpDebug "Getting Contract Information for " & strTradeSymbol & " from array"
        HandleContractInfo Parse(m.astrContracts(lPos), "|", 2)
    Else
        If IsExpiredContract(strGenesisSymbol, , True, True) Then
            m.Broker.DumpDebug "Not Getting for Contract Information for '" & strGenesisSymbol & "' from broker because expired"
        Else
            Set astrToSend = New cGdArray
            astrToSend.Create eGDARRAY_Strings
            
            strIbSymbol = BrokerSymbol(strTradeSymbol, IbSym, strIbSecType)
            If Not IbSym Is Nothing Then
                m.Broker.DumpDebug "Getting Contract Information for '" & strGenesisSymbol & "' ( " & strIbSymbol & " ) from broker"
                
                astrToSend(0) = Replace(IbSym.IbBaseSymbol, "O:", "")
                astrToSend(1) = strIbSymbol
                If strIbSecType = "FUT" Then
                    astrToSend(2) = Parse(strTradeSymbol, "-", 2)
                ElseIf strIbSecType = "FOP" Then
                    astrToSend(2) = Parse(Parse(strTradeSymbol, " ", 1), "-", 2)
                ElseIf strIbSecType = "OPT" Then
                    astrToSend(2) = Parse(strTradeSymbol, " ", 2)
                Else
                    astrToSend(2) = ""
                End If
                If strIbSecType = "STK" Then
                    astrToSend(3) = "SMART"
                Else
                    astrToSend(3) = IbSym.BrokerExchange
                End If
                astrToSend(4) = strIbSecType
                astrToSend(5) = IbSym.IbCurrency
                If strIbSecType = "FOP" Then
                    astrToSend(6) = Left(Parse(strIbSymbol, " ", 2), 1)
                    astrToSend(7) = Mid(Parse(strIbSymbol, " ", 2), 2)
                ElseIf strIbSecType = "OPT" Then
                    astrToSend(6) = Left(Parse(strTradeSymbol, " ", 3), 1)
                    astrToSend(7) = Mid(Parse(strTradeSymbol, " ", 3), 2)
                Else
                    astrToSend(6) = ""
                    astrToSend(7) = ""
                End If
                
                SendIbMessage eGDIbMessageType_GetContractDetails, astrToSend.JoinFields(vbTab)
                
                If bShowMessage Then
                    InfBox "Retrieving valid exchanges for " & strGenesisSymbol & ".|Please wait...", , , "Retrieving Exchange List", True
                End If
            Else
                m.Broker.DumpDebug "Not Getting Contract Information for '" & strGenesisSymbol & "' from broker because could not convert the symbol"
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.GetContractInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetExchangeAvailability
'' Description: Determine exchange availability for the given symbol
'' Inputs:      Symbol
'' Returns:     None
''
'' Fields:      Symbol, Local Symbol, Maturity, Exchange, Security Type,
''              Currency, Put/Call, Strike, Trading Class, Multiplier
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetExchangeAvailability(ByVal strGenesisSymbol As String)
On Error GoTo ErrSection:

    Dim astrToSend As cGdArray          ' Array of information to send
    Dim strIbSymbol As String           ' Interactive Brokers symbol
    Dim strIbSecType As String          ' Interactive Brokers security type
    Dim lPos As Long                    ' Position in the array
    Dim IbSym As cIbSymbol              ' IB Symbol object
    
    m.Broker.DumpDebug "Getting Exchange Availability for " & strGenesisSymbol & " from broker"
        
    Set astrToSend = New cGdArray
    astrToSend.Create eGDARRAY_Strings
    
    strIbSymbol = BrokerSymbol(strGenesisSymbol, IbSym, strIbSecType)
    If Not IbSym Is Nothing Then
        astrToSend(0) = IbSym.IbBaseSymbol
        astrToSend(1) = strIbSymbol
        If strIbSecType = "FUT" Then
            astrToSend(2) = Parse(strGenesisSymbol, "-", 2)
        Else
            astrToSend(2) = ""
        End If
        If strIbSecType = "STK" Then
            astrToSend(3) = "SMART"
        Else
            astrToSend(3) = IbSym.BrokerExchange
        End If
        astrToSend(4) = strIbSecType
        astrToSend(5) = IbSym.IbCurrency
        astrToSend(6) = ""
        astrToSend(7) = ""
        astrToSend(8) = IbSym.TradingClass
        astrToSend(9) = IbSym.Multiplier
        
        SendIbMessage eGDIbMessageType_GetExchangeAvailable, astrToSend.JoinFields(vbTab)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.GetExchangeAvailability"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSymbolAvailability
'' Description: Determine symbol availability for the given symbol
'' Inputs:      Symbol
'' Returns:     None
''
'' Fields:      Symbol, Local Symbol, Maturity, Exchange, Security Type,
''              Currency, Put/Call, Strike
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetSymbolAvailability(ByVal strGenesisSymbol As String)
On Error GoTo ErrSection:

    Dim astrToSend As cGdArray          ' Array of information to send
    Dim strIbSymbol As String           ' Interactive Brokers symbol
    Dim strIbSecType As String          ' Interactive Brokers security type
    Dim lPos As Long                    ' Position in the array
    Dim IbSym As cIbSymbol              ' IB Symbol object
    
    frmTest2.AddList "GetSymbolAvailability ( " & strGenesisSymbol & " )"
    
    If Len(strGenesisSymbol) > 0 Then
        m.Broker.DumpDebug "Getting Symbol Availability for '" & strGenesisSymbol & "' from broker"
            
        Set astrToSend = New cGdArray
        astrToSend.Create eGDARRAY_Strings
        
        strIbSymbol = BrokerSymbol(strGenesisSymbol, IbSym, strIbSecType)
        
        ' DAJ 11/24/2014: If we have a symbol in the conversion table that has an exchange
        ' that Trade Navigator code doesn't recognize yet, the IbSym here will not be nothing,
        ' but the symbol will be blank.  When we get the response back from the server, the local
        ' symbol will be blank which won't match the "last symbol requested" and so we will ask
        ' for the symbol over and over again...
        If (Not IbSym Is Nothing) And (Len(strIbSymbol) > 0) Then
            astrToSend(0) = IbSym.IbBaseSymbol
            astrToSend(1) = strIbSymbol
            If strIbSecType = "FUT" Then
                astrToSend(2) = Parse(strGenesisSymbol, "-", 2)
            Else
                astrToSend(2) = ""
            End If
            If strIbSecType = "STK" Then
                astrToSend(3) = "SMART"
            Else
                astrToSend(3) = IbSym.BrokerExchange
            End If
            astrToSend(4) = strIbSecType
            astrToSend(5) = IbSym.IbCurrency
            astrToSend(6) = ""
            astrToSend(7) = ""
            astrToSend(8) = IbSym.TradingClass
            astrToSend(9) = IbSym.Multiplier
            
            SendIbMessage eGDIbMessageType_GetSymbolAvailable, astrToSend.JoinFields(vbTab)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.GetSymbolAvailability"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetNextValidID
'' Description: Ask TWS for the next valid ID to use for orders
'' Inputs:      Number of IDs
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetNextValidID(Optional ByVal lNumIds As Long = 1)
On Error GoTo ErrSection:

    SendIbMessage eGDIbMessageType_GetNextValidID, Str(lNumIds)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.GetNextValidID"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NotifyNewExchanges
'' Description: Notify the user about any new exchanges
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub NotifyNewExchanges()
On Error GoTo ErrSection:

    Dim strNewExchanges As String       ' List of new exchanges that the user is enabled for
    
    If (Len(m.strNewExchanges) > 0) And (g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected) Then
        strNewExchanges = m.strNewExchanges
        m.strNewExchanges = ""
        
        If InfBox("Trade Navigator has recognized that you are authorized to stream " & strNewExchanges & " real time from Interactive Brokers.  In order to start getting this data real time, you will need to restart the stream.||Would you like to reconnect now?", "?", "+Reconnect|-Not Now", "Interactive Brokers") = "R" Then
            g.RealTime.Reconnect 0
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.NotifyNewExchanges"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExchangeInfoFromIni
'' Description: Attempt to retrieve the exchange information for the given
''              symbol from the INI file.  If not found in the INI file, return
''              the Interactive Brokers exchange for the given symbol
'' Inputs:      Genesis Symbol, Exchange List (returned), Default Exchange (returned)
'' Returns:     True if found in INI file, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExchangeInfoFromIni(ByVal strGenesisSymbol As String, strExchangeList As String, strDefaultExchange As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim IbSym As cIbSymbol              ' IB Symbol information
    
    ' Set default return values...
    bReturn = False
    strExchangeList = ""
    strDefaultExchange = ""

    If Len(BrokerSymbol(strGenesisSymbol, IbSym)) > 0 Then
        strExchangeList = GetIniFileProperty(strGenesisSymbol, "", "Exchanges", m.Broker.IniFile)
        If Len(strExchangeList) = 0 Then
            strExchangeList = IbSym.BrokerExchange
        Else
            bReturn = True
        End If
        
        strDefaultExchange = GetIniFileProperty(strGenesisSymbol, "", "DefaultExchange", m.Broker.IniFile)
        If Len(strDefaultExchange) = 0 Then
            strDefaultExchange = IbSym.BrokerExchange
        End If
    End If
    
    ExchangeInfoFromIni = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.ExchangeInfoFromIni"
    
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
    Dim BrokerSym As cIbSymbol          ' Broker symbol information
    
    strReturn = ""
    If Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0 Then
        strReturn = BrokerSym.OrderTypeMask & vbTab & BrokerSym.TifMask
    End If
    
    SymbolInformation = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.SymbolInformation"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RequestHistory
'' Description: Request History from the Traders Workstation
'' Inputs:      Genesis Price, End Date, Duration, Bar Size, What to Show
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RequestHistory(ByVal strGenesisSymbol As String, ByVal strEndDate As String, ByVal strDuration As String, ByVal strBarSize As String, ByVal strWhatToShow As String)
On Error GoTo ErrSection:

    Dim strIbSymbol As String           ' IB Symbol
    Dim IbSymbol As cIbSymbol           ' IB Symbol object
    Dim strIbSecType As String          ' IB Security type
    Dim astrArgs As cGdArray            ' Arguments to the request
    
    strIbSymbol = BrokerSymbol(strGenesisSymbol, IbSymbol, strIbSecType)
    If Len(strIbSymbol) > 0 Then
        Set astrArgs = New cGdArray
        astrArgs.Create eGDARRAY_Strings, 13
        
        astrArgs(0) = Replace(IbSymbol.IbBaseSymbol, "O:", "")
        astrArgs(1) = strIbSymbol
        astrArgs(2) = ""
        astrArgs(3) = IbSymbol.BrokerExchange
        astrArgs(4) = strIbSecType
        astrArgs(5) = IbSymbol.IbCurrency
        astrArgs(6) = ""
        astrArgs(7) = ""
        astrArgs(8) = strEndDate
        astrArgs(9) = strDuration
        astrArgs(10) = strBarSize
        astrArgs(11) = strWhatToShow
        astrArgs(12) = strGenesisSymbol
        
        SendIbMessage eGDIbMessageType_RequestHistory, astrArgs.JoinFields(vbTab)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.RequestHistory"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MessageTypeToString
'' Description: Convert a message type enumeration to a string
'' Inputs:      Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MessageTypeToString(ByVal nType As eGDIbMessageTypes) As String
On Error GoTo ErrSection:

    Select Case nType
        Case eGDIbMessageType_Connect
            MessageTypeToString = "eGDIbMessageType_Connect"
        Case eGDIbMessageType_Disconnect
            MessageTypeToString = "eGDIbMessageType_Disconnect"
        Case eGDIbMessageType_AddOrder
            MessageTypeToString = "eGDIbMessageType_AddOrder"
        Case eGDIbMessageType_AmendOrder
            MessageTypeToString = "eGDIbMessageType_AmendOrder"
        Case eGDIbMessageType_CancelOrder
            MessageTypeToString = "eGDIbMessageType_CancelOrder"
        Case eGDIbMessageType_UnloadApp
            MessageTypeToString = "eGDIbMessageType_UnloadApp"
        Case eGDIbMessageType_GetAccounts
            MessageTypeToString = "eGDIbMessageType_GetAccounts"
        Case eGDIbMessageType_GetOrders
            MessageTypeToString = "eGDIbMessageType_GetOrders"
        Case eGDIbMessageType_GetFills
            MessageTypeToString = "eGDIbMessageType_GetFills"
        Case eGDIbMessageType_GetPositions
            MessageTypeToString = "eGDIbMessageType_GetPositions"
        Case eGDIbMessageType_GetContracts
            MessageTypeToString = "eGDIbMessageType_GetContracts"
        Case eGDIbMessageType_Subscribe
            MessageTypeToString = "eGDIbMessageType_Subscribe"
        Case eGDIbMessageType_Unsubscribe
            MessageTypeToString = "eGDIbMessageType_Unsubscribe"
        
        Case eGDIbMessageType_ConnectionInfo
            MessageTypeToString = "eGDIbMessageType_ConnectionInfo"
        Case eGDIbMessageType_AppLoaded
            MessageTypeToString = "eGDIbMessageType_AppLoaded"
        Case eGDIbMessageType_AppUnloaded
            MessageTypeToString = "eGDIbMessageType_AppUnloaded"
        Case eGDIbMessageType_Heartbeat
            MessageTypeToString = "eGDIbMessageType_Heartbeat"
        Case eGDIbMessageType_Order
            MessageTypeToString = "eGDIbMessageType_Order"
        Case eGDIbMessageType_AccountR
            MessageTypeToString = "eGDIbMessageType_AccountR"
        Case eGDIbMessageType_OrderR
            MessageTypeToString = "eGDIbMessageType_OrderR"
        Case eGDIbMessageType_FillR
            MessageTypeToString = "eGDIbMessageType_FillR"
        Case eGDIbMessageType_PositionR
            MessageTypeToString = "eGDIbMessageType_PositionR"
        Case eGDIbMessageType_ContractR
            MessageTypeToString = "eGDIbMessageType_ContractR"
        Case eGDIbMessageType_QuoteR
            MessageTypeToString = "eGDIbMessageType_QuoteR"
            
        Case eGDIbMessageType_GetNextValidID
            MessageTypeToString = "eGDIbMessageType_GetNextValidID"
        Case eGDIbMessageType_NextValidID
            MessageTypeToString = "eGDIbMessageType_NextValidID"
        Case eGDIbMessageType_GetExchangeAvailable
            MessageTypeToString = "eGDIbMessageType_GetExchangeAvailable"
        Case eGDIbMessageType_ExchangeAvailable
            MessageTypeToString = "eGDIbMessageType_ExchangeAvailable"
        Case eGDIbMessageType_GetContractDetails
            MessageTypeToString = "eGDIbMessageType_GetContractDetails"
        Case eGDIbMessageType_ContractDetails
            MessageTypeToString = "eGDIbMessageType_ContractDetails"
        Case eGDIbMessageType_Fill
            MessageTypeToString = "eGDIbMessageType_Fill"
        Case eGDIbMessageType_Position
            MessageTypeToString = "eGDIbMessageType_Position"
        Case eGDIbMessageType_AddComboOrder
            MessageTypeToString = "eGDIbMessageType_AddComboOrder"
        Case eGDIbMessageType_AmendComboOrder
            MessageTypeToString = "eGDIbMessageType_AmendComboOrder"
        Case eGDIbMessageType_RequestHistory
            MessageTypeToString = "eGDIbMessageType_RequestHistory"
        Case eGDIbMessageType_History
            MessageTypeToString = "eGDIbMessageType_History"
        Case eGDIbMessageType_Message
            MessageTypeToString = "eGDIbMessageType_Message"
        Case eGDIbMessageType_GetSymbolAvailable
            MessageTypeToString = "eGDIbMessageType_GetSymbolAvailable"
        Case eGDIbMessageType_SymbolAvailable
            MessageTypeToString = "eGDIbMessageType_SymbolAvailable"
            
        Case Else
            MessageTypeToString = Str(nType)
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.MessageTypeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MessageTypeFromString
'' Description: Convert a string into a message type enumeration
'' Inputs:      String
'' Returns:     Message Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MessageTypeFromString(ByVal strType As String) As eGDIbMessageTypes
On Error GoTo ErrSection:

    Dim nReturn As eGDIbMessageTypes    ' Return value for the function

    Select Case strType
        Case "eGDIbMessageType_Connect"
            nReturn = eGDIbMessageType_Connect
        Case "eGDIbMessageType_Disconnect"
            nReturn = eGDIbMessageType_Disconnect
        Case "eGDIbMessageType_AddOrder"
            nReturn = eGDIbMessageType_AddOrder
        Case "eGDIbMessageType_AmendOrder"
            nReturn = eGDIbMessageType_AmendOrder
        Case "eGDIbMessageType_CancelOrder"
            nReturn = eGDIbMessageType_CancelOrder
        Case "eGDIbMessageType_UnloadApp"
            nReturn = eGDIbMessageType_UnloadApp
        Case "eGDIbMessageType_GetAccounts"
            nReturn = eGDIbMessageType_GetAccounts
        Case "eGDIbMessageType_GetOrders"
            nReturn = eGDIbMessageType_GetOrders
        Case "eGDIbMessageType_GetFills"
            nReturn = eGDIbMessageType_GetFills
        Case "eGDIbMessageType_GetPositions"
            nReturn = eGDIbMessageType_GetPositions
        Case "eGDIbMessageType_GetContracts"
            nReturn = eGDIbMessageType_GetContracts
        Case "eGDIbMessageType_Subscribe"
            nReturn = eGDIbMessageType_Subscribe
        Case "eGDIbMessageType_Unsubscribe"
            nReturn = eGDIbMessageType_Unsubscribe
        
        Case "eGDIbMessageType_ConnectionInfo"
            nReturn = eGDIbMessageType_ConnectionInfo
        Case "eGDIbMessageType_AppLoaded"
            nReturn = eGDIbMessageType_AppLoaded
        Case "eGDIbMessageType_AppUnloaded"
            nReturn = eGDIbMessageType_AppUnloaded
        Case "eGDIbMessageType_Heartbeat"
            nReturn = eGDIbMessageType_Heartbeat
        Case "eGDIbMessageType_Order"
            nReturn = eGDIbMessageType_Order
        Case "eGDIbMessageType_AccountR"
            nReturn = eGDIbMessageType_AccountR
        Case "eGDIbMessageType_OrderR"
            nReturn = eGDIbMessageType_OrderR
        Case "eGDIbMessageType_FillR"
            nReturn = eGDIbMessageType_FillR
        Case "eGDIbMessageType_PositionR"
            nReturn = eGDIbMessageType_PositionR
        Case "eGDIbMessageType_ContractR"
            nReturn = eGDIbMessageType_ContractR
        Case "eGDIbMessageType_QuoteR"
            nReturn = eGDIbMessageType_QuoteR
            
        Case "eGDIbMessageType_GetNextValidID"
            nReturn = eGDIbMessageType_GetNextValidID
        Case "eGDIbMessageType_NextValidID"
            nReturn = eGDIbMessageType_NextValidID
        Case "eGDIbMessageType_GetExchangeAvailable"
            nReturn = eGDIbMessageType_GetExchangeAvailable
        Case "eGDIbMessageType_ExchangeAvailable"
            nReturn = eGDIbMessageType_ExchangeAvailable
        Case "eGDIbMessageType_GetContractDetails"
            nReturn = eGDIbMessageType_GetContractDetails
        Case "eGDIbMessageType_ContractDetails"
            nReturn = eGDIbMessageType_ContractDetails
        Case "eGDIbMessageType_Fill"
            nReturn = eGDIbMessageType_Fill
        Case "eGDIbMessageType_Position"
            nReturn = eGDIbMessageType_Position
        Case "eGDIbMessageType_AddComboOrder"
            nReturn = eGDIbMessageType_AddComboOrder
        Case "eGDIbMessageType_AmendComboOrder"
            nReturn = eGDIbMessageType_AmendComboOrder
        Case "eGDIbMessageType_RequestHistory"
            nReturn = eGDIbMessageType_RequestHistory
        Case "eGDIbMessageType_History"
            nReturn = eGDIbMessageType_History
        Case "eGDIbMessageType_Message"
            nReturn = eGDIbMessageType_Message
        Case "eGDIbMessageType_GetSymbolAvailable"
            nReturn = eGDIbMessageType_GetSymbolAvailable
        Case "eGDIbMessageType_SymbolAvailable"
            nReturn = eGDIbMessageType_SymbolAvailable
            
        Case Else
            nReturn = -1&
    End Select

    MessageTypeFromString = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.MessageTypeFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HaveContractIdForSymbol
'' Description: Determine if we have the contract ID for the given symbol
'' Inputs:      Genesis Symbol
'' Returns:     True if have contract ID, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HaveContractIdForSymbol(ByVal strGenesisSymbol As String) As Boolean
On Error GoTo ErrSection:

    HaveContractIdForSymbol = m.ContractIds.Exists(strGenesisSymbol)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.HaveContractIdForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerPrice
'' Description: Convert a Genesis price to the broker price
'' Inputs:      Genesis Price, IB Symbol Object
'' Returns:     Broker Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerPrice(ByVal dGenesisPrice As Double, ByVal IbSymbol As cIbSymbol) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = Str(dGenesisPrice)
    If Not IbSymbol Is Nothing Then
        strReturn = m.Broker.BrokerPrice(dGenesisPrice, IbSymbol.PriceMult, IbSymbol.PriceFormat)
    End If
    
    BrokerPrice = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.BrokerPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisPrice
'' Description: Attempt to convert the broker price to a Genesis price
'' Inputs:      Broker Price, IB Symbol Object
'' Returns:     Genesis Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisPrice(ByVal strBrokerPrice As String, ByVal IbSymbol As cIbSymbol) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    
    dReturn = 0#
    If Not IbSymbol Is Nothing Then
        dReturn = m.Broker.GenesisPrice(strBrokerPrice, IbSymbol.PriceMult)
    End If
    
    GenesisPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.GenesisPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderType
'' Description: Convert the broker order type to a Genesis order type
'' Inputs:      Broker order type
'' Returns:     Genesis Order Type (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderType(ByVal strBrokerOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderType        ' Return value for the function

    nReturn = -1&
    Select Case UCase(strBrokerOrderType)
        Case "MKT"
            nReturn = eTT_OrderType_Market
        Case "LMT"
            nReturn = eTT_OrderType_Limit
        Case "STP"
            nReturn = eTT_OrderType_Stop
        Case "STP LMT"
            nReturn = eTT_OrderType_StopWithLimit
        Case "MOC"
            nReturn = eTT_OrderType_MarketOnClose
        Case "LOC"
            nReturn = eTT_OrderType_LimitCloseOnly
        Case "MIT"
            nReturn = eTT_OrderType_MIT
        Case "MOO"
            nReturn = eTT_OrderType_MarketOnOpen
    End Select
    
    GenesisOrderType = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.GenesisOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderStatus
'' Description: Convert the broker order status to a Genesis order status
'' Inputs:      Broker Order Status
'' Returns:     Genesis Order Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderStatus(ByVal strBrokerOrderStatus As String) As eTT_OrderStatus
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderStatus      ' Return value from the function

    Select Case UCase(strBrokerOrderStatus)
        Case "PRESUBMITTED"
            nReturn = eTT_OrderStatus_PreSubmitted
        Case "SUBMITTED"
            nReturn = eTT_OrderStatus_Working
        Case "PENDINGCANCEL"
            nReturn = eTT_OrderStatus_CancelPending
        Case "CANCELLED", "APICANCELLED"
            nReturn = eTT_OrderStatus_Cancelled
        Case "FILLED"
            nReturn = eTT_OrderStatus_Filled
        Case "INACTIVE"
            nReturn = eTT_OrderStatus_Inactive
        Case "ERROR"
            nReturn = eTT_OrderStatus_Error
        Case "REJECTED"
            nReturn = eTT_OrderStatus_Rejected
    End Select
    
    GenesisOrderStatus = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.GenesisOrderStatus"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerBuySell
'' Description: Convert a Genesis Buy/Sell to the broker buy/sell
'' Inputs:      Genesis Buy/Sell
'' Returns:     Broker Buy/Sell
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerBuySell(ByVal bBuySell As Boolean) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    If bBuySell Then
        strReturn = "BUY"
    Else
        strReturn = "SELL"
    End If
    
    BrokerBuySell = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.BrokerBuySell"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerTIF
'' Description: Convert a Genesis Time-in-Force to the broker Time-in-Force
'' Inputs:      Genesis Time-in-Force
'' Returns:     Broker Time-in-Force
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerTIF(ByVal nGenesisTIF As eTT_TimeInForce) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    Select Case nGenesisTIF
        Case eTT_TimeInForce_Day
            strReturn = "DAY"
        Case eTT_TimeInForce_GTC
            strReturn = "GTC"
        Case eTT_TimeInForce_GTD
            strReturn = "CTD"
    End Select
    
    BrokerTIF = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.BrokerTIF"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerOrderString
'' Description: Build an order string to send to the broker
'' Inputs:      Order, New Order?
'' Returns:     Broker Order String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerOrderString(ByVal Order As cPtOrder, ByVal bNewOrder As Boolean) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Order.NumberOfLegs = 1 Then
        strReturn = BrokerOrderStringSingle(Order, bNewOrder)
    ElseIf Order.NumberOfLegs > 1 Then
        strReturn = BrokerOrderStringCombo(Order, bNewOrder)
    End If
    
    BrokerOrderString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.BrokerOrderString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerOrderStringSingle
'' Description: Build an order string for a single leg order to send to the broker
'' Inputs:      Order, New Order?
'' Returns:     Broker Order String
''
'' Fields:      GenesisID, BrokerID, Account, IB Base Symbol, Symbol, Order Type,
''              Buy/Sell, Quantity, Limit, Stop, Strike, Good Til Date, Exchange,
''              Security Type, TIF, Expiration, Currency, Contract ID, Primary Exchange,
''              Oca Group
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerOrderStringSingle(ByVal Order As cPtOrder, ByVal bNewOrder As Boolean) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrOrder As cGdArray           ' Array of order information to send to broker
    Dim IbSym As cIbSymbol              ' Broker symbol object
    Dim strBrokerSymbol As String       ' Broker symbol
    Dim strSecurityType As String       ' Broker security type
    
    strReturn = ""
    strBrokerSymbol = BrokerSymbol(Order.Symbol, IbSym, strSecurityType)
    
    If (Len(strBrokerSymbol) > 0) And (Not IbSym Is Nothing) Then
        Set astrOrder = New cGdArray
        astrOrder(0) = Order.GenesisOrderID
        If bNewOrder Then
            astrOrder(1) = ""
        Else
            If InStr(Order.BrokerID, "-") <> 0 Then
                astrOrder(1) = Parse(Order.BrokerID, "-", 2)
            Else
                astrOrder(1) = Order.BrokerID
            End If
        End If
        astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(3) = Replace(IbSym.IbBaseSymbol, "O:", "")
        astrOrder(4) = strBrokerSymbol
        astrOrder(5) = BrokerOrderType(Order.OrderType)
        astrOrder(6) = BrokerBuySell(Order.Buy)
        astrOrder(7) = Str(Order.Quantity)
        If Order.OrderType = eTT_OrderType_MIT Then
            astrOrder(8) = ""
            astrOrder(9) = BrokerPrice(Order.LimitPrice, IbSym)
        Else
            astrOrder(8) = BrokerPrice(Order.LimitPrice, IbSym)
            astrOrder(9) = BrokerPrice(Order.StopPrice, IbSym)
        End If
        astrOrder(10) = ""
        If (Order.Expiration <= 0) Then
            astrOrder(11) = ""
        Else
            astrOrder(11) = Format(Order.Expiration, "YYYYMMDD")
        End If
        If Len(Order.Exchange) = 0 Then
            astrOrder(12) = IbSym.BrokerExchange
        Else
            astrOrder(12) = Order.Exchange
        End If
        astrOrder(13) = strSecurityType
        astrOrder(14) = BrokerTIF(Order.TimeInForce)
        astrOrder(15) = ""
        astrOrder(16) = IbSym.IbCurrency
        astrOrder(17) = Str(m.ContractIds(Order.Symbol))
        astrOrder(18) = IbSym.IbPrimaryExchange
        astrOrder(19) = Order.OcaGroup
        
        strReturn = astrOrder.JoinFields(vbTab)
    End If
    
    BrokerOrderStringSingle = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.BrokerOrderStringSingle"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerOrderStringCombo
'' Description: Build an order string for a multi-leg order to send to the broker
'' Inputs:      Order, New Order?
'' Returns:     Broker Order String
''
'' Fields:      GenesisID, BrokerID, Account, Symbol, Exchange, Security Type,
''              Currency, Order Type, Quantity, Limit, Stop, TIF, Good Til Date,
''              Number of Legs, Leg1, Leg2, Leg3, Leg4
''
'' Leg:         Contract ID;Multiplier;Side;Exchange;O/C
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerOrderStringCombo(ByVal Order As cPtOrder, ByVal bNewOrder As Boolean) As String
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Array of order information to send to broker
    Dim IbSym As cIbSymbol              ' Broker symbol object
    Dim strBrokerSymbol As String       ' Broker symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrLeg As cGdArray             ' Array of leg information to send to broker
    Dim OrderLeg As cOrderLeg           ' Order leg
    
    Set astrOrder = New cGdArray
    astrOrder.Create eGDARRAY_Strings, 14 + Order.NumberOfLegs
    
    astrOrder(0) = Order.GenesisOrderID
    If bNewOrder Then
        astrOrder(1) = ""
    Else
        astrOrder(1) = Order.BrokerID
    End If
    astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
    astrOrder(3) = "" ' Symbol -- Will be filled in with the first leg below
    astrOrder(4) = "" ' Exchange -- Will be filled in with the first leg below
    astrOrder(5) = "BAG"
    astrOrder(6) = "" ' Currency -- Will be filled in with the first leg below
    astrOrder(7) = BrokerOrderType(Order.OrderType)
    astrOrder(8) = Str(Order.Quantity)
    astrOrder(9) = BrokerPrice(Order.LimitPrice, IbSym)
    astrOrder(10) = BrokerPrice(Order.StopPrice, IbSym)
    astrOrder(11) = BrokerTIF(Order.TimeInForce)
    If (Order.Expiration <= 0) Then
        astrOrder(12) = ""
    Else
        astrOrder(12) = Format(Order.Expiration, "YYYYMMDD")
    End If
    astrOrder(13) = Str(Order.NumberOfLegs)
    
    For lIndex = 1 To Order.NumberOfLegs
        Set OrderLeg = Order.OrderLegs(lIndex)
        
        Set astrLeg = New cGdArray
        astrLeg.Create eGDARRAY_Strings, 5
        
        strBrokerSymbol = BrokerSymbol(OrderLeg.Symbol, IbSym)
        If lIndex = 1 Then
            astrOrder(3) = IbSym.IbCurrency
            astrOrder(4) = IbSym.BrokerExchange
            astrOrder(6) = IbSym.IbCurrency
        End If
        
        astrLeg(0) = m.ContractIds(OrderLeg.Symbol)
        astrLeg(1) = Str(OrderLeg.Multiplier)
        astrLeg(2) = BrokerBuySell(OrderLeg.IsBuy)
        astrLeg(3) = IbSym.BrokerExchange
        If OrderLegIsEntry(Order, lIndex, False) Then
            astrLeg(4) = "O"
        Else
            astrLeg(4) = "C"
        End If
        
        astrOrder(lIndex + 13) = astrLeg.JoinFields(";")
    Next lIndex
    
    BrokerOrderStringCombo = astrOrder.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.BrokerOrderStringCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleContractInfo
'' Description: Do the appropriate thing with the given contract information
'' Inputs:      Contract Information
'' Returns:     None
''
'' Fields:      IB Base Symbol, Symbol, Expiration, Exchange, Sec Type, Currency,
''              Put/Call, Strike Price, Order Types, Exchanges, Contract ID,
''              Primary Exchange
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleContractInfo(ByVal strContractInfo As String)
On Error GoTo ErrSection:

    Dim astrInfo As cGdArray            ' Contract information split out
    Dim strGenesisSymbol As String      ' Genesis symbol for contact information
    Dim lPos As Long                    ' Position in the array
    Dim IbSym As cIbSymbol              ' IB Symbol object
    Dim strExchange As String           ' IB Exchange
    Dim lExpiration As Long             ' Expiration date for the symbol
    
    If Len(strContractInfo) > 0 Then
        Set astrInfo = New cGdArray
        astrInfo.SplitFields strContractInfo, vbTab
        
        If UCase(astrInfo(11)) = "LSE" Then
            strExchange = "LSE"
        Else
            strExchange = astrInfo(3)
        End If
        
        strGenesisSymbol = GenesisSymbol(astrInfo(1), strExchange, astrInfo(4), IbSym)
        If Len(strGenesisSymbol) > 0 Then
            If m.astrContracts.BinarySearch(strGenesisSymbol & "|", lPos, eGdSort_MatchUsingSearchStringLength) Then
                m.astrContracts(lPos) = strGenesisSymbol & "|" & strContractInfo
            Else
                m.astrContracts.Add strGenesisSymbol & "|" & strContractInfo, lPos
            End If
            
            If Len(astrInfo(2)) > 0 Then
                lExpiration = JulFromLong(CLng(Val(astrInfo(2))))
                g.Broker.SetExpirationDateForSymbol strGenesisSymbol, lExpiration
                g.Broker.SaveSymbolInfo strGenesisSymbol, lExpiration, 1#
            End If
            
            IbSym.IbExchanges = astrInfo(9)
            IbSym.IbPrimaryExchange = astrInfo(11)
            m.Symbols.UpdateItem IbSym
        
            SetIniFileProperty strGenesisSymbol, astrInfo(9), "Exchanges", m.Broker.IniFile
            SetIniFileProperty strGenesisSymbol, astrInfo(3), "DefaultExchange", m.Broker.IniFile
            g.Broker.UpdateBrokerContractInfo m.Broker.Broker, strGenesisSymbol, astrInfo(9), astrInfo(3)
            
            If Len(astrInfo(10)) > 0 Then
                If m.ContractIds.Exists(strGenesisSymbol) Then
                    m.ContractIds(strGenesisSymbol) = astrInfo(10)
                Else
                    m.ContractIds.Add astrInfo(10), strGenesisSymbol
                End If
            End If
            
            CheckHeldOrders
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.HandleContractInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleExchangeInfo
'' Description: Do the appropriate thing with the given exchange information
'' Inputs:      Exchange Information
'' Returns:     None
''
'' Fields:      Exchange, Error
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleExchangeInfo(ByVal strExchangeInfo As String)
On Error GoTo ErrSection:

    Dim strExchange As String           ' Exchange
    Dim strError As String              ' Error Message
    Dim lPos As Long                    ' Position in an array

    strExchange = Parse(strExchangeInfo, vbTab, 1)
    strError = Parse(strExchangeInfo, vbTab, 2)

    If strError = "0" Then
        If m.astrExchanges.BinarySearch(strExchange, lPos) = False Then
            m.astrExchanges.Add strExchange, lPos
            m.astrExchanges.Serialize m.strExchFile, True
            
            If (UCase(strExchange) <> "SNFE") Or (HasModule("E_SFE") = False) Then
                If ((g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected) Or (g.RealTime.ConnectionStatus = eGDConnectionStatus_Connecting)) Then
                    If Len(m.strNewExchanges) = 0 Then
                        m.strNewExchanges = strExchange
                    ElseIf InStr(m.strNewExchanges, strExchange) = 0 Then
                        m.strNewExchanges = m.strNewExchanges & "," & strExchange
                    End If
                End If
            End If
            
            If UCase(strExchange) = "SNFE" Then
                g.Broker.SfeAllowed(m.Broker.Broker) = Int(CurrentTime)
            End If
        End If
    Else
        If m.astrExchanges.BinarySearch(strExchange, lPos) = True Then
            m.astrExchanges.Remove lPos
            m.astrExchanges.Serialize m.strExchFile, True
        End If
    
        If UCase(strExchange) = "SNFE" Then
            g.Broker.SfeAllowed(m.Broker.Broker) = kNullData
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.HandleExchangeInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleSymbolAvailability
'' Description: Do the appropriate thing with the given symbol availability
'' Inputs:      Symbol Information
'' Returns:     None
''
'' Fields:      Symbol, Local Symbol, Expiry, Exchange, Security Type, Currency,
''              Put/Call, Strike Price, Error Code, Error String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleSymbolAvailability(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Fields split out from the message
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim IbSymbol As cIbSymbol           ' IB Symbol object
    Dim lPos As Long                    ' Position in the array
    Dim strBaseSymbol As String         ' Base symbol
    Dim dPreviousRequest As Double      ' Previous value for last symbol request

    frmTest2.AddList "HandleSymbolAvailability ( " & strMessage & " )"

    If Len(strMessage) > 0 Then
        dPreviousRequest = m.dLastSymbolRequest
        SetLastSymbolRequest 0#, "HandleSymbolAvailability1"
        m.lNumSymbolTries = 0&
        
        Set astrFields = New cGdArray
        astrFields.SplitFields strMessage, vbTab
        
        ' DAJ 07/28/2015: 2121 is more of a warning than an error in this case so I think we can allow real-time for it...
        ' Warning: ... product is trading on the basis other than currency price
        If astrFields(8) = "2121" Then
            astrFields(8) = "0"
        End If
        
        strGenesisSymbol = GenesisSymbol(astrFields(1), astrFields(3), astrFields(4), IbSymbol, astrFields(2))
        If Len(strGenesisSymbol) > 0 Then
            strBaseSymbol = Parse(strGenesisSymbol, "-", 1)
                
            ' DAJ 09/25/2014: The -1 error code for this message was a special case that I
            ' put in that basically means that the market was closed when I tried to subscribe
            ' to data.  If we get a -1 back, add it to a queue to keep requesting until we get it...
            If astrFields(8) = "-1" Then
                m.astrMarketClosed.Add strGenesisSymbol
                m.Broker.DumpDebug vbTab & vbTab & strGenesisSymbol & " added to Market Closed queue"
            
                If m.astrEnabledDates.BinarySearch(strBaseSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                    If Val(Parse(m.astrEnabledDates(lPos), vbTab, 2)) < (Int(CurrentTime) - 7) Then
                        m.astrEnabledDates.Remove lPos
                        
                        If m.astrEnabledSymbols.BinarySearch(strBaseSymbol, lPos) = True Then
                            m.astrEnabledSymbols.Remove lPos
                            m.Broker.DumpDebug vbTab & vbTab & strGenesisSymbol & " removed from enabled symbols array because last success was a while ago"
                        End If
                    End If
                End If
            Else
                If astrFields(8) = "0" Then
                    If m.astrEnabledDates.BinarySearch(strBaseSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                        m.astrEnabledDates(lPos) = strBaseSymbol & vbTab & Str(Int(CurrentTime))
                    Else
                        m.astrEnabledDates.Add strBaseSymbol & vbTab & Str(Int(CurrentTime)), lPos
                    End If
                ElseIf m.astrEnabledDates.BinarySearch(strBaseSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                    m.astrEnabledDates.Remove lPos
                End If
                
                If m.astrEnabledSymbols.BinarySearch(strBaseSymbol, lPos) = True Then
                    If astrFields(8) = "0" Then
                        m.astrEnabledSymbols(lPos) = strBaseSymbol
                    Else
                        m.astrEnabledSymbols.Remove lPos
                        m.Broker.DumpDebug vbTab & vbTab & strGenesisSymbol & " removed from enabled symbols array"
                    End If
                ElseIf astrFields(8) = "0" Then
                    m.astrEnabledSymbols.Add strBaseSymbol, lPos
                    m.Broker.DumpDebug vbTab & vbTab & strGenesisSymbol & " added to enabled symbols array"
                End If
                
                If m.astrExchanges.BinarySearch(astrFields(3), lPos) = False Then
                    If astrFields(8) = "0" Then
                        m.astrExchanges.Add astrFields(3), lPos
                        m.astrExchanges.Serialize m.strExchFile, True
                        m.Broker.DumpDebug vbTab & vbTab & astrFields(3) & " added to enabled exchanges array"
                    End If
                End If
            End If
        End If
    End If
    
    If strGenesisSymbol = m.strLastSymbol Then
        m.Broker.DumpDebug vbTab & "HandleSymbolAvailability -- Symbols match ( Last = '" & m.strLastSymbol & "'; Received = '" & Str(strGenesisSymbol) & "' )"
        
        If GetNextSymbol = False Then
            FinishedLastSymbol
        End If
    Else
        m.Broker.DumpDebug vbTab & "HandleSymbolAvailability -- Symbols don't match ( Last = '" & m.strLastSymbol & "'; Received = '" & Str(strGenesisSymbol) & "' )"
        SetLastSymbolRequest dPreviousRequest, "HandleSymbolAvailability2"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.HandleSymbolAvailability"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckHeldOrders
'' Description: Check held orders to see if we now have all contract ID's
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckHeldOrders()
On Error GoTo ErrSection:

    Dim lOrder As Long                  ' Index into a for loop
    Dim Order As cPtOrder               ' Held order
    
    For lOrder = m.HeldOrdersOut.Count To 1 Step -1
        Set Order = m.HeldOrdersOut(lOrder)
        If OrderReadyToSubmit(Order) Then
            AddOrder Order
            m.HeldOrdersOut.Remove lOrder
        End If
    Next lOrder
    
    For lOrder = m.HeldOrdersIn.Count To 1 Step -1
        Set Order = New cPtOrder
        If OrderFromInfo(m.HeldOrdersIn(lOrder), Order) Then
            HandleMessage eGDIbMessageType_Order, m.HeldOrdersIn(lOrder)
            m.HeldOrdersIn.Remove lOrder
        End If
    Next lOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.CheckHeldOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderReadyToSubmit
'' Description: Verify that the given order has contract ID's for all legs
'' Inputs:      Order
'' Returns:     True if has all contract ids, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderReadyToSubmit(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lLeg As Long                    ' Index into a for loop
    
    bReturn = False
    If Not Order Is Nothing Then
        bReturn = True
        
        If Order.NumberOfLegs > 1 Then
            For lLeg = 1 To Order.NumberOfLegs
                If m.ContractIds.Exists(Order.OrderLegs(lLeg).Symbol) = False Then
                    bReturn = False
                    Exit For
                End If
            Next lLeg
        Else
            bReturn = m.ContractIds.Exists(Order.Symbol)
        End If
    End If
    
    OrderReadyToSubmit = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.OrderReadyToSubmit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RequestContractInfoForOrderLegs
'' Description: Request contract information for any legs for which we haven't already
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RequestContractInfoForOrderLegs(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lLeg As Long                    ' Index into a for loop
    Dim strSymbol As String             ' Symbol for the order leg
    
    If Not Order Is Nothing Then
        If Order.NumberOfLegs > 1 Then
            For lLeg = 1 To Order.NumberOfLegs
                strSymbol = Order.OrderLegs(lLeg).Symbol
                If m.ContractIds.Exists(strSymbol) = False Then
                    GetContractInfo strSymbol
                End If
            Next lLeg
        Else
            GetContractInfo Order.Symbol
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.RequestContractInfoForOrderLegs"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbolForContractID
'' Description: Determine the Genesis Symbol for the given contract ID
'' Inputs:      Contract ID
'' Returns:     Genesis Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisSymbolForContractID(ByVal strContractID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    strReturn = ""
    For lIndex = 1 To m.ContractIds.Count
        If m.ContractIds(lIndex) = strContractID Then
            strReturn = m.ContractIds.Key(lIndex)
            Exit For
        End If
    Next lIndex
    
    GenesisSymbolForContractID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.GenesisSymbolForContractID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LegsFromInfo
'' Description: Build order legs out of the string of information
'' Inputs:      Leg Information, Order (out)
'' Returns:     True if successful, False otherwise
''
'' Fields:      ContractID|Ratio,ContractID|Ratio,...
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LegsFromInfo(ByVal strLegs As String, Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrLegs As cGdArray            ' Legs information broken out into an array
    Dim astrLeg As cGdArray             ' Leg information broken out into an array
    Dim strGenesisSymbol As String      ' Genesis symbol for the contract ID
    Dim OrderLeg As cOrderLeg           ' Order leg object
    Dim lRatio As Long                  ' Ratio for the leg
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = True
    If Len(strLegs) > 0 Then
        Set astrLegs = New cGdArray
        astrLegs.SplitFields strLegs, ","
        
        For lIndex = 0 To astrLegs.Size - 1
            Set astrLeg = New cGdArray
            astrLeg.SplitFields astrLegs(lIndex), "|"
            
            strGenesisSymbol = GenesisSymbolForContractID(astrLeg(0))
            If Len(strGenesisSymbol) > 0 Then
                Set OrderLeg = New cOrderLeg
                lRatio = CLng(Val(astrLeg(1)))
                
                OrderLeg.SymbolOrSymbolID = strGenesisSymbol
                OrderLeg.IsBuy = (lRatio > 0)
                OrderLeg.LegNumber = lIndex + 1
                OrderLeg.Multiplier = Abs(lRatio)
                OrderLeg.OrderID = Order.OrderID
                
                Order.OrderLegs(lIndex + 1) = OrderLeg
            Else
                bReturn = False
                Exit For
            End If
        Next lIndex
    End If
    
    LegsFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.LegFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OcaGroupForOrder
'' Description: Build the Oca Group name for an order
'' Inputs:      Order
'' Returns:     Oca group name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OcaGroupForOrder(ByVal Order As cPtOrder) As String
On Error GoTo ErrSection:

    Dim strMachineID As String          ' Machine ID
    Dim dCurrentTime As Double          ' Current time
    
    strMachineID = StripStr(UCase(RI_GetMachineID), "- ")
    dCurrentTime = CurrentTime("GMT")
        
    OcaGroupForOrder = strMachineID & "_" & Order.GenesisOrderID & "_" & Format(dCurrentTime, "YYMMDD_HHNNSS")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.OcaGroupForOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMessage
'' Description: Show a message from Interactive Brokers to the user
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowMessage(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim brokerMessage As cBrokerMessage ' Broker message

    If Len(strMessage) > 0 Then
        Set brokerMessage = New cBrokerMessage
        brokerMessage.FromString strMessage
        
        If brokerMessage("code") <> "2109" Then
            ShowAlertPopup brokerMessage("text"), m.Broker.BrokerName, vbLeftJustify
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.ShowMessage"
    
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
    Dim BrokerSym As cIbSymbol          ' Broker symbol object
    Dim strBrokerSymbol As String       ' Broker symbol
    Dim strIbSecType As String          ' IB Security type
    Dim astrReturn As cGdArray          ' Array of return information
    
    strReturn = ""
    strBrokerSymbol = BrokerSymbol(strGenesisSymbol, BrokerSym, strIbSecType)
    If Len(strBrokerSymbol) > 0 Then
        Set astrReturn = New cGdArray
        astrReturn.Create eGDARRAY_Strings, 13
        
        astrReturn(0) = strAccount
        astrReturn(2) = strBrokerSymbol
        astrReturn(6) = BrokerSym.BrokerExchange
        astrReturn(7) = strIbSecType
        
        strReturn = astrReturn.JoinFields(vbTab)
    End If
    
    PositionMessageForGenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.PositionMessageForGenesisSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetNextSymbol
'' Description: Get symbol availability for the next symbol in the list
'' Inputs:      None
'' Returns:     True if requested next symbol, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetNextSymbol() As Boolean
On Error GoTo ErrSection

    Dim bReturn As Boolean              ' Return value for the function
    Dim strSymbol As String             ' Symbol to request

    bReturn = False
    m.lSymbolIndex = m.lSymbolIndex + 1
    If m.lSymbolIndex < m.astrGenesisSymbols.Size Then
        strSymbol = FrontMonth(m.astrGenesisSymbols(m.lSymbolIndex))
        Do While (Len(strSymbol) = 0) And (m.lSymbolIndex < m.astrGenesisSymbols.Size)
            m.lSymbolIndex = m.lSymbolIndex + 1
            If m.lSymbolIndex < m.astrGenesisSymbols.Size Then
                strSymbol = FrontMonth(m.astrGenesisSymbols(m.lSymbolIndex))
            End If
        Loop
        
        If Len(strSymbol) > 0 Then
            If g.RealTime.ConnectionStatus <> eGDConnectionStatus_Connecting Then
                StatusMsg "Retreiving data availability for symbol " & Str(m.lSymbolIndex + 1) & " of " & Str(m.astrGenesisSymbols.Size) & " from " & m.Broker.BrokerName
            End If
            GetSymbolAvailability strSymbol
            
            SetLastSymbolRequest gdTickCount, "GetNextSymbol"
            m.Broker.DumpDebug vbTab & "Set Last Symbol changed from '" & m.strLastSymbol & "' to '" & strSymbol & "'"
            m.strLastSymbol = strSymbol
            m.lNumSymbolTries = 1&
            
            bReturn = True
        End If
    End If
    
    GetNextSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.GetNextSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SfeAllowed
'' Description: Determine if the user is enabled to see SFE data
'' Inputs:      None
'' Returns:     True if user is allowed to see SFE, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SfeAllowed() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim IbSym As cIbSymbol              ' IB Symbol object
    
    bReturn = False
    For lIndex = 1 To m.Symbols.Count
        Set IbSym = m.Symbols(lIndex)
        If IbSym.BrokerExchange = "SNFE" Then
            If EnabledSymbols.BinarySearch(IbSym.GenesisBase) Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    SfeAllowed = bReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cIntBrokers.SfeAllowed"
    
End Function

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbolFromPieces
'' Description: Attempt to convert the broker symbol pieces to a Genesis symbol
'' Inputs:      Broker Base Symbol, Broker Expiry, Broker Exchange,
''              Broker Security Type, Broker Currency, Broker Right ( P/C ),
''              Broker Strike, IB Symbol Object
'' Returns:     Genesis Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbolFromPieces(ByVal strBrokerBase As String, ByVal strBrokerExpiry As String, ByVal strBrokerExchange As String, ByVal strBrokerSecType As String, ByVal strBrokerCurrency As String, ByVal strBrokerRight As String, ByVal strBrokerStrike As String, Optional IbSymbol As cIbSymbol) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return from the function
    Dim OptSym As cOptionSymbol         ' Option symbol object
                
    strReturn = ""
    If Len(strBrokerBase) > 0 Then
        Select Case strBrokerSecType
            Case "FUT"
                Set IbSymbol = m.Symbols.GetByBrokerSymbol(strBrokerBase, strBrokerExchange)
                If Not IbSymbol Is Nothing Then
                    strReturn = IbSymbol.GenesisBase & "-" & strBrokerExpiry
                End If
                
            Case "STK"
                Set IbSymbol = m.Symbols.GetByBrokerSymbol("!", "SMART")
                If Not IbSymbol Is Nothing Then
                    strReturn = strBrokerSymbol
                End If
                
            Case "FOP"
                ' CME: ESG1 C1330
                ' ECBOT: C OYM  MAR 11  12250
                Set OptSym = New cOptionSymbol
                OptSym.FromIb strBrokerSymbol, strBrokerExchange, True
                Set IbSymbol = m.Symbols.GetByBrokerSymbol("O:" & OptSym.BrokerBase, strBrokerExchange)
                If Not IbSymbol Is Nothing Then
                    OptSym.Strike = OptSym.BrokerStrike * IbSymbol.StrikeMult
                    OptSym.BaseSymbol = Mid(IbSymbol.GenesisBase, 3)
                    strReturn = OptSym.ToGenesis
                End If
            
            Case "OPT"
                ' F     110319P00017000
                Set OptSym = New cOptionSymbol
                OptSym.FromIb strBrokerSymbol, strBrokerExchange, False
                Set IbSymbol = m.Symbols.GetByBrokerSymbol("@", "SMART")
                If Not IbSymbol Is Nothing Then
                    OptSym.Strike = OptSym.BrokerStrike * IbSymbol.StrikeMult
                    strReturn = OptSym.ToGenesis
                End If
                
            Case "CASH"
                Set IbSymbol = m.Symbols.GetByBrokerSymbol(strBrokerSymbol, strBrokerExchange)
                If Not IbSymbol Is Nothing Then
                    If GetSymbolID(IbSymbol.GenesisBase) > 0 Then
                        strReturn = IbSymbol.GenesisBase
                    End If
                End If
                
        End Select
    End If

    GenesisSymbol = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.GenesisSymbol"
    
End Function
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetLastSymbolRequest
'' Description: Set the last symbol request value
'' Inputs:      Last Symbol Request Value, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetLastSymbolRequest(ByVal dLastSymbolRequest As Double, ByVal strSource As String)
On Error GoTo ErrSection:

    m.Broker.DumpDebug vbTab & "LastSymbolRequest changed from " & Str(m.dLastSymbolRequest) & " to " & Str(dLastSymbolRequest) & " ( " & strSource & " )"
    m.dLastSymbolRequest = dLastSymbolRequest
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.SetLastSymbolRequest"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FrontMonth
'' Description: Determine the front month for the base symbol given
'' Inputs:      Base symbol
'' Returns:     Front month symbol ( blank if invalid )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FrontMonth(ByVal strBaseSymbol As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If (InStr(strBaseSymbol, "O:") = 0) And (IsAlpha(Left(strBaseSymbol, 1)) = True) Then
        strReturn = mMain.ConvertToTradeSymbol(strBaseSymbol & "-067", Date)
        If InStr(strReturn, "-067") > 0 Then
            strReturn = ""
            
        ' DAJ 11/24/2014: Blank out the return value if we cannot convert the Genesis
        ' symbol to an Interactive Brokers symbol...
        ElseIf Len(BrokerSymbol(strReturn)) = 0 Then
            strReturn = ""
        End If
    End If
    
    FrontMonth = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.FrontMonth"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FinishedLastSymbol
'' Description: Update things after the all of the symbol availability checks
''              are finished
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FinishedLastSymbol()
On Error GoTo ErrSection:

    Dim strFileName As String           ' Name of the last known symbols file
    Dim strPrevious As String           ' Previous symbols enabled for real-time streaming
    Dim strCurrent As String            ' Current symbols enabled for real-time streaming
    Dim strBrokerName As String         ' Display name for the broker
    Dim strMessage As String            ' Message displayed to the user
    Dim strReturn As String             ' Return value from the InfBox

    strFileName = LksFile
    m.astrEnabledSymbols.Serialize strFileName, True
    LKSCRC = CalcFileCrc(strFileName)
    
    strFileName = LksdFile
    m.astrEnabledDates.Serialize strFileName, True
    LKSDCRC = CalcFileCrc(strFileName)
    
    If m.astrExchanges.BinarySearch("SNFE") = True Then
        g.Broker.SfeAllowed(m.Broker.Broker) = Int(CurrentTime)
    Else
        g.Broker.SfeAllowed(m.Broker.Broker) = kNullData
    End If
    
    strBrokerName = m.Broker.BrokerName
    m.bRequestingSymbols = False
    StatusMsg "Finished retreiving data availability from " & strBrokerName
    
    strPrevious = m.astrPrevEnabled.JoinFields("|")
    strCurrent = m.astrEnabledSymbols.JoinFields("|")
    
    If strCurrent = strPrevious Then
        mMain.ShowAlertPopup "Real-time symbols have been verified through|" & strBrokerName, strBrokerName, vbCenter
    Else
        m.Broker.DumpDebug vbTab & "Previous symbols = '" & strPrevious & "'"
        m.Broker.DumpDebug vbTab & "Current symbols = '" & strCurrent & "'"
        
        strMessage = "The list of symbols for which you are|enabled for real-time streaming through " & strBrokerName & " has changed.||You will need to restart your data stream|for the changes to take effect.||Would you like to do that now?|"
        strReturn = InfBox(strMessage, "?", "Yes|+-No", strBrokerName, , 15)
        m.Broker.DumpDebug vbTab & "User answered '" & strReturn & "' to '" & Replace(strMessage, "|", " ") & "'"
        
        If strReturn = "Y" Then
            g.RealTime.Reconnect 0
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.FinishedLastSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AskForEnablements
'' Description: Determine if we need to ask for streaming enablements
'' Inputs:      None
'' Returns:     True if we need to ask, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AskForEnablements() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim dCurrentTime As Double          ' Current time
    
    dCurrentTime = CurrentTime("NY")
    
    Select Case Weekday(dCurrentTime)
        Case vbFriday
            bReturn = (Hour(dCurrentTime) < 18)
        
        Case vbSaturday
            bReturn = False
            
        Case vbSunday
            bReturn = (Hour(dCurrentTime) > 11)
        
        Case Else
            bReturn = True
    
    End Select
    
    AskForEnablements = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIntBrokers.AskForEnablements"
    
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
    Set m.Symbols = New cBrokerSymbols
    Set m.ContractIds = New cGdTree
    Set m.HeldOrdersOut = New cGdTree
    Set m.HeldOrdersIn = New cGdTree
    
    Set m.astrFillSymbols = New cGdArray
    m.astrFillSymbols.Create eGDARRAY_Strings
    
    Set m.astrEnabledSymbols = New cGdArray
    m.astrEnabledSymbols.Create eGDARRAY_Strings
    
    Set m.astrPrevEnabled = New cGdArray
    m.astrPrevEnabled.Create eGDARRAY_Strings
    
    Set m.astrGenesisSymbols = New cGdArray
    m.astrGenesisSymbols.Create eGDARRAY_Strings
    
    m.bRequestingSymbols = False
    m.dLastSymbolRequest = 0#
    m.strLastSymbol = ""
    m.lNumSymbolTries = 0&
    
    Set m.astrMarketClosed = New cGdArray
    m.astrMarketClosed.Create eGDARRAY_Strings
    m.dLastMarketClosedRequest = 0#
    m.strLastMarketClosed = ""
    
    Set m.astrEnabledDates = New cGdArray
    m.astrEnabledDates.Create eGDARRAY_Strings
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.Class_Initialize"
    
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
    Set m.Symbols = Nothing
    Set m.astrExchanges = Nothing
    Set m.astrContracts = Nothing
    Set m.ContractIds = Nothing
    Set m.HeldOrdersOut = Nothing
    Set m.HeldOrdersIn = Nothing
    Set m.astrFillSymbols = Nothing
    Set m.astrEnabledSymbols = Nothing
    Set m.astrPrevEnabled = Nothing
    Set m.astrGenesisSymbols = Nothing
    Set m.astrMarketClosed = Nothing
    Set m.astrEnabledDates = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIntBrokers.Class_Terminate"
    
End Sub
