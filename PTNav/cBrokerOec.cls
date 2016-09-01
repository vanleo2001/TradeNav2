VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerOec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerOec.cls
'' Description: Wrapper class for interacting with the Open E-Cry servers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/25/2011   DAJ         Utilized standard message enums
'' 03/21/2012   DAJ         Fix for amending a partially filled order
'' 10/16/2013   DAJ         Resurrected and moved over to newer stuff
'' 10/17/2013   DAJ         Copy OEC dlls over to appropriate place
'' 10/24/2013   DAJ         Contract size override for forex; Don't show "Cancelled"
'' 10/25/2013   DAJ         Handle position fills to help with carried position mismatches
'' 10/28/2013   DAJ         Don't show "Auto-Updated" message
'' 01/02/2014   DAJ         Use current stream time for first StatusDate on an order
'' 01/09/2014   DAJ         Changed BrokerName for FuturePath Trading
'' 04/10/2014   DAJ         Ignore certain messages coming back on an order
'' 11/07/2014   DAJ         Fix for duplicate order in BrokerInfo upon unsolicited order modify
'' 11/14/2014   DAJ         Added support for MIT, On-Close, and On-Open orders
'' 02/19/2015   DAJ         Changed 'Open E-Cry' to 'GAIN Capital'; Only give enabled symbols real-time
'' 02/19/2015   DAJ         Add @OEC Forex symbols to the enabled symbols as well
'' 07/21/2015   DAJ         Reverse Side value on fill if fill type is Opposite or SpreadPair
'' 08/17/2015   DAJ         Changed 'MGE' exchange to 'MGF'
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kRegKey = "Software\Genesis Financial Data Services\Navigator Suite\General"

Private Type mPrivate
    Broker As cBroker                   ' Broker base class object
    BrokerKv As cBrokerKeyValue         ' Routines for key-value broker messages
    Symbols As cBrokerSymbols           ' Broker symbol conversion object
    
    astrEnabledSymbols As cGdArray      ' Array of Genesis symbols the user is enabled for
    Fills As cGdTree                    ' Collection of fills
    PositionFills As cGdTree            ' Collection of position fills
    ContractSizes As cGdTree            ' Collection of contract sizes
    
    strUserName As String               ' User Name to login with
    strPassword As String               ' Password for the login
    strIP As String                     ' IP address of the server to connect to
    strPort As String                   ' Port to connect to on the server
    strIgnoreComments As String         ' Comments to ignore
    
    strOcsFile As String                ' OEC Contract Sizes file
End Type
Private m As mPrivate

Public Property Get Broker() As cBroker
    Set Broker = m.Broker
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
    m.Broker.AskedPassword = True
End Property

Public Property Get IP() As String
    IP = m.strIP
End Property
Public Property Let IP(ByVal strIP As String)
    m.strIP = strIP
End Property

Public Property Get Port() As String
    Port = m.strPort
End Property
Public Property Let Port(ByVal strPort As String)
    m.strPort = strPort
End Property

Public Property Get ContractSize(ByVal strGenesisSymbol As String, ByVal strAccountNumber As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the property
    Dim strKey As String                ' Key into the collection
    
    lReturn = kNullData
    strKey = strGenesisSymbol & "|" & strAccountNumber
    
    If m.ContractSizes.Exists(strKey) Then
        lReturn = m.ContractSizes(strKey)
    End If
    
    ContractSize = lReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerOec.ContractSize.Get"
    
End Property
Private Property Let ContractSize(ByVal strGenesisSymbol As String, ByVal strAccountNumber As String, ByVal lContractSize As Long)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    
    strKey = strGenesisSymbol & "|" & strAccountNumber
    If m.ContractSizes.Exists(strKey) Then
        m.ContractSizes(strKey) = lContractSize
    Else
        m.ContractSizes.Add lContractSize, strKey
    End If
    
    If Not g.Profit Is Nothing Then
        g.Profit.ContractSize(strGenesisSymbol, strAccountNumber) = lContractSize
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerOec.ContractSize.Let"
    
End Property

Private Property Get LKSCRC() As Long
    LKSCRC = GetRegistryValue(rkLocalMachine, kRegKey, "LKSC" & Str(m.Broker.Broker), 0&)
End Property
Private Property Let LKSCRC(ByVal lLksCrc As Long)
    SetRegistryValue rkLocalMachine, kRegKey, "LKSC" & Str(m.Broker.Broker), lLksCrc, True
End Property

Private Property Get LksFile() As String
    LksFile = AddSlash(App.Path) & "Provided\LKS." & Str(m.Broker.Broker)
End Property

'Private Property Get ConnectString() As String
'    ConnectString = m.strUserName & vbTab & m.strPassword & vbTab & m.strIP & vbTab & m.strPort
'End Property
'Private Property Get ConnectStringLog() As String
'    ConnectStringLog = m.strUserName & vbTab & SetStrToChar(m.strPassword, "*") & vbTab & m.strIP & vbTab & m.strPort
'End Property

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
    Dim strBrokerPath As String         ' Broker path

    Select Case nBroker
        Case eTT_AccountType_Oec
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "Oec"
            m.Broker.IniFile = AddSlash(App.Path) & "Oec.INI"
            m.Broker.BrokerName = "GAIN Capital" ' "Open E-Cry"
            m.Broker.ProcessName = "GenOec"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\Oec\"
            m.Broker.ProcessPath = strBrokerPath & "GenOec.EXE"
            m.Broker.ControlID = "Oec"
            strSymbolFile = AddSlash(App.Path) & "Provided\OecToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\OecCnct.INI"
            
        Case eTT_AccountType_FptOec
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "FptOec"
            m.Broker.IniFile = AddSlash(App.Path) & "FptOec.INI"
            m.Broker.BrokerName = "FuturePath Trading (GAIN)"
            m.Broker.ProcessName = "GenFptOec"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\FptOec\"
            m.Broker.ProcessPath = strBrokerPath & "GenFptOec.EXE"
            m.Broker.ControlID = "FptOec"
            strSymbolFile = AddSlash(App.Path) & "Provided\FptOecToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\FptOecCnct.INI"
            
    End Select
    
    m.Broker.CopyProcess AddSlash(App.Path) & "..\Brokers\GenOec.EXE"
    FileCopy AddSlash(App.Path) & "..\Brokers\OecDlls\*.DLL", strBrokerPath, True
    
    Set m.Symbols = New cBrokerSymbols
    m.Symbols.FromFile strSymbolFile

    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Connect) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Disconnect) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccounts) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccountsMode) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ConnectInfo) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ChangePassword) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Refresh) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewActivity) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_BrokerView) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewOnline) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_VerifyPositions) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_AccountDetails) = False
    
    m.strOcsFile = AddSlash(App.Path) & "Provided\OCS." & Str(nBroker)
    LoadContractSizes
    
    m.strIgnoreComments = GetIniFileProperty("Ignore", "", "Comments", m.Broker.ConnectIni)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerOec.Init"

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

    ShowConfig = (Len(m.strUserName) = 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.ShowConfig"
    
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

    ShowBrokerConnectionInfo = GetLoginInfo(strUserName, True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.ShowBrokerConnectionInfo"
    
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
            m.Broker.AskedPassword = bReturn
        End If
    Else
        If ((strNewUserName <> m.strUserName) And (Len(strNewUserName) > 0)) Or (bSwitching = True) Then
            bReturn = GetLoginInfo(strNewUserName, , bSwitching)
            If bReturn = True Then
                m.Broker.AskedPassword = True
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
    RaiseError "cBrokerOec.ShowLogin"
    
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
    RaiseError "cBrokerOec.SendConnect"
    
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
    RaiseError "cBrokerOec.SendDisconnect"
    
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
    RaiseError "cBrokerOec.UnloadApp"
    
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
    Dim OptionSym As cOptionSymbol      ' Option symbol object
    
    strReturn = ""
    If Len(strGenesisSymbol) > 0 Then
        Select Case UCase(SecurityType(strGenesisSymbol, True))
            Case "F"
                Set BrokerSym = m.Symbols.GetByGenesisSymbol(Parse(strGenesisSymbol, "-", 1))
                If Not BrokerSym Is Nothing Then
                    strReturn = BrokerSym.BrokerBase & BrokerSym.ContractToMY(Parse(strGenesisSymbol, "-", 2))
                End If
                
            Case "FO"
                Set OptionSym = New cOptionSymbol
                OptionSym.FromGenesis strGenesisSymbol
                Set BrokerSym = m.Symbols.GetByGenesisSymbol("O:" & OptionSym.BaseSymbol)
                If Not BrokerSym Is Nothing Then
                    strReturn = OptionSym.ToOec(Replace(BrokerSym.BrokerBase, "O:", ""), BrokerSym.StrikeMult, BrokerSym.StrikePriceFormat)
                End If
            
            Case "S"
                Set BrokerSym = m.Symbols.GetByGenesisSymbol("!")
                If Not BrokerSym Is Nothing Then
                    If InStr(strGenesisSymbol, "_P") <> 0 Then
                        strReturn = Replace(strGenesisSymbol, "_P", "p")
                    ElseIf InStr(strGenesisSymbol, "_R") <> 0 Then
                        strReturn = Replace(strGenesisSymbol, "_R", "r")
                    ElseIf InStr(strGenesisSymbol, "_U") <> 0 Then
                        strReturn = Replace(strGenesisSymbol, "_U", "/U")
                    ElseIf InStr(strGenesisSymbol, "_W") <> 0 Then
                        If Right(strGenesisSymbol, 2) = "_W" Then
                            strReturn = Replace(strGenesisSymbol, "_W", "/WS")
                        Else
                            strReturn = Replace(strGenesisSymbol, "_W", "w")
                        End If
                    ElseIf InStr(strGenesisSymbol, "_") <> 0 Then
                        strReturn = Replace(strGenesisSymbol, "_", "/")
                    Else
                        strReturn = strGenesisSymbol
                    End If
                End If
                
            Case "SO"
            
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
    RaiseError "cBrokerOec.BrokerSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Attempt to convert the broker symbol to the Genesis symbology
'' Inputs:      Broker Symbol, Broker Exchange, Broker Symbol object, For Position?,
''              Security Type
'' Returns:     Genesis Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbol(ByVal strBrokerSymbol As String, ByVal strBrokerExchange As String, Optional BrokerSym As cBrokerSymbol, Optional ByVal bPosition As Boolean = False, Optional ByVal strSecurityType As String = "Future") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strBase As String               ' Base symbol
    Dim OptionSym As New cOptionSymbol  ' Option symbol object
    Dim astrSymbol As cGdArray          ' Symbol broken out into fields
    
    strReturn = ""
    Select Case UCase(strBrokerExchange)
        Case "CBL", "CBOT", "CME", "COE", "EUR", "HKF", "IPE", "KCBT", "LFF", "MGF", "NLI", "NYM", "SFE", "SGX"
            Select Case UCase(strSecurityType)
                Case "OPTION"
                    Set astrSymbol = New cGdArray
                    astrSymbol.SplitFields strBrokerSymbol, " "
                    
                    If astrSymbol.Size = 2 Then
                        ' <BaseSym><M><Y> {C|P}<Strike> ( e.g. OESV3 C1710 )
                        strBase = "O:" & Left(astrSymbol(0), Len(astrSymbol(0)) - 2)
                        
                        If bPosition Then
                            Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol(strBase, strBrokerExchange)
                        Else
                            Set BrokerSym = m.Symbols.GetByBrokerSymbol(strBase, strBrokerExchange)
                        End If
                        
                        If Not BrokerSym Is Nothing Then
                            Set OptionSym = New cOptionSymbol
                            OptionSym.FromOec strBrokerSymbol
                            OptionSym.ConvertBrokerInfo Replace(BrokerSym.GenesisBase, "O:", ""), BrokerSym.StrikeMult
                            strReturn = OptionSym.ToGenesis
                        End If
                    End If
                
                Case "FUTURE"
                    If InStr(strBrokerSymbol, " ") = 0 Then
                        strBase = Left(strBrokerSymbol, Len(strBrokerSymbol) - 2)
                        
                        If bPosition Then
                            Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol(strBase, strBrokerExchange)
                        Else
                            Set BrokerSym = m.Symbols.GetByBrokerSymbol(strBase, strBrokerExchange)
                        End If
                        If Not BrokerSym Is Nothing Then
                            strReturn = BrokerSym.GenesisBase & "-" & BrokerSym.ContractFromMY(Right(strBrokerSymbol, 2))
                        End If
                    End If
            
            End Select
            
        Case "FX"
            If UCase(strSecurityType) = "FOREX" Then
                If bPosition Then
                    Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol(strBrokerSymbol, strBrokerExchange)
                Else
                    Set BrokerSym = m.Symbols.GetByBrokerSymbol(strBrokerSymbol, strBrokerExchange)
                End If
                
                If Not BrokerSym Is Nothing Then
                    strReturn = BrokerSym.GenesisBase
                End If
            End If
            
    End Select
    
    GenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.GenesisSymbol"
    
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
            strReturn = "Market"
        Case eTT_OrderType_Stop
            strReturn = "Stop"
        Case eTT_OrderType_Limit
            strReturn = "Limit"
        Case eTT_OrderType_StopWithLimit
            strReturn = "StopWithLimit"
        Case eTT_OrderType_MarketOnClose
            strReturn = "MarketOnClose"
        Case eTT_OrderType_MIT
            strReturn = "MarketIfTouched"
        Case eTT_OrderType_MarketOnOpen
            strReturn = "MarketOnOpen"
    End Select
    
    BrokerOrderType = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.BrokerOrderType"
    
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
    RaiseError "cBrokerOec.AddOrder"
    
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
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AmendOrder, strOrderString
        bReturn = True
    End If
    
    AmendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.AmendOrder"
    
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
        m.Broker.SendBrokerMessage eGDBrokerMessageType_CancelOrder, strOrderString
        bReturn = True
    End If
    
    CancelOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.CancelOrder"
    
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
        m.Broker.SendBrokerMessage eGDBrokerMessageType_CancelOrder, strOrderString
        bReturn = True
    End If
    
    ParkOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.ParkOrder"
    
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
    RaiseError "cBrokerOec.Refresh"
    
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
    RaiseError "cBrokerOec.GetAccounts"
    
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

    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetOrders, strAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerOec.GetOrders"
    
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

    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetFills, strAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerOec.GetFills"
    
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

    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetPositions, strAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerOec.GetPositions"
    
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
    RaiseError "cBrokerOec.GetSymbols"
    
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
    RaiseError "cBrokerOec.NextGenesisID"
    
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
    RaiseError "cBrokerOec.StartingGenesisIdForOptNav"
    
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
    RaiseError "cBrokerOec.IsTradeableSymbol"
    
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
    
    IsEnabledSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.IsEnabledSymbol"
    
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
    RaiseError "cBrokerOec.OrderTypeAllowed"
    
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
    RaiseError "cBrokerOec.TimeInForceAllowed"
    
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

    Dim dReturn As Double               ' Return value of the function
    Dim strDate As String               ' Date portion of the string
    Dim strTime As String               ' Time portion of the string
    Dim dTime As Double                 ' Time as fractional part of day
    
    dReturn = 0#
    If Len(strDateTime) > 0 Then
        ' YYYYMMDD HH:MM:SS
        strDate = Parse(strDateTime, " ", 1)
        strTime = Parse(strDateTime, " ", 2)
        
        If Len(strTime) > 0 Then
            dTime = (Val(Parse(strTime, ":", 1)) / 24#) + (Val(Parse(strTime, ":", 2)) / 1440#) + (Val(Parse(strTime, ":", 3)) / 86400#)
        Else
            dTime = 0#
        End If
        
        dReturn = CDbl(JulFromLong(CLng(Val(strDate)))) + dTime
    End If
    
    DateToDouble = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.DateToDouble"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming App Mail message from the GenOec program
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDBrokerMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    Select Case nType
        Case eGDBrokerMessageType_OecOrderIdChanged
            HandleOrderIdChanged strMessage
            
        Case eGDBrokerMessageType_PositionFillRefresh
            HandlePositionFill strMessage
            
        Case eGDBrokerMessageType_Symbols
            HandleSymbols strMessage
        
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerOec.HandleMessage"
    
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
    RaiseError "cBrokerOec.StatusFromInfo"
    
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
    RaiseError "cBrokerOec.AccountFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromInfo
'' Description: Fill an order object from the given message
'' Inputs:      Message, Order
'' Returns:     True if OK, False otherwise
''
'' Fields:      GenesisID, BrokerID, Account, Symbol, Exchange, Order Type,
''              Buy/Sell, Quantity, Stop Price, Limit Price, TIF, Expiration,
''              Order Status, Status Date, Quantity Filled, Message
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderFromInfo(ByVal strMessage As String, Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker message object
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim strError As String              ' Error Message
    Dim lAccountID As Long              ' Account ID
    Dim strFirstField As String         ' First field in the string
    Dim astrID As cGdArray              ' Broker ID broken out into an array
    Dim astrStatus As cGdArray          ' Status broken out into an array

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            
            brokerMsg.FromString strMessage
            
            brokerMsg("Message") = Replace(brokerMsg("Message"), "<equals>", "=")
            
            strGenesisSymbol = GenesisSymbol(brokerMsg("Symbol"), brokerMsg("Exchange"), BrokerSym, , brokerMsg("SecurityType"))
            strError = brokerMsg("Message")
            
            If Len(strGenesisSymbol) > 0 Then
                If m.Broker.IsBroker Then
                    lAccountID = -1&
                Else
                    lAccountID = kNullData
                End If
                
                Set astrID = New cGdArray
                astrID.SplitFields brokerMsg("BrokerID"), ";"
                
                brokerMsg.Add "ExchangeID", astrID(0)
                brokerMsg("BrokerID") = astrID(1)
                
                Set astrStatus = New cGdArray
                astrStatus.SplitFields brokerMsg("Status"), ";"
                
                If UCase(astrStatus(0)) = "WORKING" Then
                    If Len(astrStatus(2)) = 0 Then
                        brokerMsg("Status") = astrStatus(0)
                    Else
                        brokerMsg("Status") = astrStatus(2)
                    End If
                Else
                    brokerMsg("Status") = astrStatus(0)
                End If
                
                Set Order = m.BrokerKv.OrderFromMessage(brokerMsg, m.Broker, strGenesisSymbol, BrokerSym, lAccountID)
                
                ' DAJ 01/02/2014: When we get the first "Sent" message back from OEC, they are setting
                ' the Status Date to the wall time.  If your wall time is off ( like Pete's was ), the
                ' order can immediately look old and we will do a refresh because of the unconfirmed
                ' order ( the first message we get has a negative broker ID )...
                If Left(Order.BrokerID, 1) = "-" Then
                    Order.StatusDate = ConvertToBrokerDate(CurrentTime, m.Broker.Broker, Order.Symbol)
                End If
                
                ' Clear out the "Cancelled" message for the order so we don't display it to the user...
                If (Order.Status = eTT_OrderStatus_Cancelled) And (UCase(Order.Message) = "CANCELLED") Then
                    Order.Message = ""
                
                ' Also clear out the "Auto-Updated" message when a Stop order gets changed to a Market order...
                ' DAJ 10/31/2014: Ignore any "Auto-*" messages...
                'ElseIf (Order.Status = eTT_OrderStatus_Working) And (UCase(Order.Message) = "AUTO-UPDATED") Then
                ElseIf Left(UCase(Order.Message), 5) = "AUTO-" Then
                    Order.Message = ""
                
                ' DAJ 04/10/2014: Looks like some other software is setting the comment in the order to
                ' "chart" which is causing us to show a rejected message to the user...
                ElseIf InStr("," & UCase(m.strIgnoreComments) & ",", "," & UCase(Order.Message) & ",") <> 0 Then
                    Order.Message = ""
                End If
                
                ' DAJ 07/20/2015: If the fill type is 'Opposite' or 'SpreadPair', that means that it is the
                ' negative leg of the spread and, therefore, the side is the opposite of what is coming in ( my
                ' guess is that the side is being inherited from the original order )...
                If (UCase(brokerMsg("Type")) = "OPPOSITE") Or (UCase(brokerMsg("Type")) = "SPREADPAIR") Then
                    Order.Buy = Not Order.Buy
                End If
                
                bReturn = True
            End If
        End If
    End If
    
    OrderFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.OrderFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillFromInfo
'' Description: Fill a fill object from the given message
'' Inputs:      Message, Fill, Genesis Order ID
'' Returns:     True if OK, False otherwise
''
'' Fields:      FillID, BrokerID, Account, Symbol, Exchange, Side, Quantity,
''              Price, Time, Type, Commission, Comments
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillFromInfo(ByVal strMessage As String, Fill As cPtFill, Optional strGenesisOrderID As String = "") As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker message object
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim lAccountID As Long              ' Account ID
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
            m.Fills.Clear
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            
            brokerMsg.FromString strMessage
            
            strGenesisSymbol = GenesisSymbol(brokerMsg("Symbol"), brokerMsg("Exchange"), BrokerSym, , brokerMsg("SecurityType"))
            If Len(strGenesisSymbol) > 0 Then
                If m.Broker.IsBroker Then
                    lAccountID = -1&
                Else
                    lAccountID = kNullData
                End If
                Set Fill = m.BrokerKv.FillFromMessage(brokerMsg, m.Broker, strGenesisSymbol, BrokerSym, lAccountID)
                strGenesisOrderID = brokerMsg("GenesisID")
                
                If m.Fills.Exists(Fill.BrokerID) = False Then
                    m.Fills.Add Fill, Fill.BrokerID
                End If
                
                bReturn = True
            End If
        End If
    End If
    
    FillFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.FillFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionFromInfo
'' Description: Fill a position object from the given message
'' Inputs:      Message, Position
'' Returns:     True if OK, False otherwise
''
'' Fields:      Account, Symbol, Exchange, Date, Long, Short, Previous, Fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionFromInfo(ByVal strMessage As String, Position As cPtPosition) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker message object
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim lAccountID As Long              ' Account ID
    Dim strFirstField As String         ' First field in the string
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
            GetSymbols
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
            
            strGenesisSymbol = GenesisSymbol(brokerMsg("Symbol"), brokerMsg("Exchange"), BrokerSym, , brokerMsg("SecurityType"))
            If Len(strGenesisSymbol) > 0 Then
                If m.Broker.IsBroker Then
                    lAccountID = -1&
                Else
                    lAccountID = kNullData
                End If
                
                Set Position = m.BrokerKv.PositionFromMessage(brokerMsg, m.Broker, strGenesisSymbol, lAccountID)
                
                For lIndex = 1 To m.PositionFills.Count
                    Set Fill = m.PositionFills(lIndex)
                    If Fill.SymbolOrSymbolID = Position.SymbolOrSymbolID Then
                        Position.CarriedFills.Add Fill
                    End If
                Next lIndex
                
                For lIndex = 1 To m.Fills.Count
                    Set Fill = m.Fills(lIndex)
                    If Fill.SymbolOrSymbolID = Position.SymbolOrSymbolID Then
                        Position.CurrentFills.Add Fill
                    End If
                Next lIndex
                
                bReturn = True
            End If
        End If
    End If
    
    PositionFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.PositionFromInfo"
    
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
    
            If CalcFileCrc(strFileName) <> LKSCRC Then
                KillFile strFileName
            End If
    
            If m.astrEnabledSymbols.Serialize(strFileName, False) = False Then
                ' If there is no CRC entry in the registry and they have connected before, give
                ' them all of the symbols like we have been doing...
                If (LKSCRC = 0&) And (m.Broker.LastDateConnected > 0) Then
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
    
        Set astrSymbols = m.Symbols.GenesisSymbolList(False)
        For lIndex = 0 To astrSymbols.Size - 1
            If InStr(astrSymbols(lIndex), "@OEC") > 0 Then
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
    RaiseError "cBrokerOec.GenesisSymbolList"
    
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
    RaiseError "cBrokerOec.TimeZone"
    
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
    RaiseError "cBrokerOec.JustConnected"
    
End Sub

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
    RaiseError "cBrokerOec.BrokerPrice"
    
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
    RaiseError "cBrokerOec.GenesisPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderType
'' Description: Convert the broker order type to a Genesis order type
'' Inputs:      Broker Order Type
'' Returns:     Genesis Order Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderType(ByVal strBrokerOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderType        ' Return value for the function

    nReturn = -1&
    Select Case UCase(strBrokerOrderType)
        Case "MARKET"
            nReturn = eTT_OrderType_Market
        Case "STOP"
            nReturn = eTT_OrderType_Stop
        Case "LIMIT"
            nReturn = eTT_OrderType_Limit
        Case "STOPWITHLIMIT"
            nReturn = eTT_OrderType_StopWithLimit
        Case "MARKETONCLOSE"
            nReturn = eTT_OrderType_MarketOnClose
        Case "MARKETIFTOUCHED"
            nReturn = eTT_OrderType_MIT
        Case "MARKETONOPEN"
            nReturn = eTT_OrderType_MarketOnOpen
    End Select
    
    GenesisOrderType = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.GenesisOrderType"
    
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

    Dim nReturn As eTT_OrderStatus      ' Return value for the function
    Dim astrStatus As cGdArray          ' Array of status information
    
    Set astrStatus = New cGdArray
    astrStatus.SplitFields strBrokerOrderStatus, ";"
    
    nReturn = -1&
    If astrStatus.Size > 0 Then
        Select Case UCase(astrStatus(0))
            Case "SENT"
                nReturn = eTT_OrderStatus_Sent
            Case "ACCEPTED"
                nReturn = eTT_OrderStatus_Approved
            Case "HELD"
                nReturn = eTT_OrderStatus_Held
            Case "WORKING"
                Select Case UCase(astrStatus(3))
                    Case ""
                        nReturn = eTT_OrderStatus_Working
                    Case "MODIFYING"
                        nReturn = eTT_OrderStatus_AmendPending
                    Case "MODIFIED"
                        nReturn = eTT_OrderStatus_Working
                    Case "CANCELLING"
                        nReturn = eTT_OrderStatus_CancelPending
                End Select
            Case "COMPLETED"
                nReturn = eTT_OrderStatus_Filled
            Case "CANCELLED"
                nReturn = eTT_OrderStatus_Cancelled
            Case "REJECTED"
                nReturn = eTT_OrderStatus_Rejected
        End Select
    End If
    
    GenesisOrderStatus = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.GenesisOrderStatus"
    
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
    RaiseError "cBrokerOec.BrokerBuySell"
    
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
            strReturn = "GTD"
    End Select
    
    BrokerTIF = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.BrokerTIF"
    
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
    
    brokerMsg.Add "UserName", m.strUserName
    brokerMsg.Add "Password", m.strPassword
    brokerMsg.Add "IP", m.strIP
    brokerMsg.Add "Port", m.strPort
    brokerMsg.Add "Uuid", "11b73161-29e0-48fa-9011-8b3640238edc"
    
    ConnectString = brokerMsg.ToString(bForLog)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.ConnectString"
    
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
    'If bNewOrder = False Then
    '    Order.BrokerID = Order.ExchangeID & ";" & Order.BrokerID
    'End If
    BrokerOrderString = m.BrokerKv.OrderToMessage(Order, bNewOrder, strBrokerSymbol, BrokerSym, m.Broker, strPreviousGenesisID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.BrokerOrderString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleOrderIdChanged
'' Description: Handle an Order ID Changed event
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      Account, Genesis ID, Old Order ID, New Order ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleOrderIdChanged(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker message object
    Dim Order As cPtOrder               ' Order object
    Dim strTicketNumber As String       ' Current ticket number for the order ID
    Dim PrevOrder As cPtOrder           ' Previous order
    
    If Len(strMessage) > 0 Then
        Set brokerMsg = New cBrokerMessage
        
        brokerMsg.FromString strMessage
        
        Set Order = m.Broker.LoadOrder(brokerMsg("Account"), brokerMsg("GenesisID"), brokerMsg("OldOrderID"))
        If Order Is Nothing Then
            strTicketNumber = TicketNumberFromID(brokerMsg("OldOrderID"))
            If Len(strTicketNumber) > 0 Then
                Set Order = m.Broker.LoadOrder(brokerMsg("Account"), brokerMsg("GenesisID"), strTicketNumber)
            End If
        End If
        If Not Order Is Nothing Then
            m.Broker.DumpDebug "Order with BrokerID = '" & Order.BrokerID & "' and GenesisID = '" & Order.GenesisOrderID & "' exists"
            
            Set PrevOrder = New cPtOrder
            PrevOrder.LoadForBrokerID Order.BrokerID
            Order.PreviousBrokerID = Order.BrokerID
            Order.PreviousOrder = PrevOrder
            Order.BrokerID = brokerMsg("NewOrderID")
            Order.Save
            
            m.Broker.DumpDebug Order.OrderText(True, True, True) & ": " & OrderStatus(Order.Status) & " --> Broker Order Changed from '" & Order.PreviousBrokerID & "'"
            
            OrderCallback Order
            
            ' 11/07/2014 DAJ: A customer was running into some issues with unsolicited orders since he
            ' is using the Open E-Cry software to trade.  Turns out that when we get an unsolicited order
            ' modification, we are leaving two copies of the same order in the BrokerInfo which is causing
            ' other issues.  Passing the "UsePreviousBrokerId" variable as True now so that we can find
            ' the order in the BrokerInfo structure instead of creating a new copy...
            m.Broker.BrokerInfo.AddOrder Order, False, True, True
            SendOrderToOptionNav Order, False
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerOec.HandleOrderIdChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TicketNumberFromID
'' Description: Attempt to find the current ticket number for the given ID
'' Inputs:      ID
'' Returns:     Ticket Number (blank if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TicketNumberFromID(ByVal strID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    
    strReturn = ""
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [ExchangeOrderID]='" & strID & "';", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        strReturn = rs!BrokerOrderID
    End If
    
    TicketNumberFromID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.TicketNumberFromID"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLoginInfo
'' Description: Get login information from the user
'' Inputs:      User Name, Show IP?, Switching
'' Returns:     True if user OKed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetLoginInfo(Optional ByVal strUserName As String = "", Optional ByVal bShowIP As Boolean = False, Optional ByVal bSwitching As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = frmLoginOec.ShowMe(m.Broker.Broker, strUserName, bShowIP, bSwitching)
    If bReturn = True Then
        m.strUserName = frmLoginOec.UserName
        m.strPassword = frmLoginOec.Password
        m.strIP = frmLoginOec.IP
        m.strPort = frmLoginOec.Port
        
        m.Broker.AskedPassword = True
    End If
    
    GetLoginInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.GetLoginInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandlePositionFill
'' Description: Handle a position fill message from OEC
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandlePositionFill(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    Dim Fill As cPtFill                 ' Fill object

    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
            m.PositionFills.Clear
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set Fill = New cPtFill
            
            If FillFromInfo(strMessage, Fill) Then
                If Len(Fill.BrokerID) = 0 Then
                    Fill.BrokerID = Fill.Symbol & "_" & Format(Fill.FillDate, "yyyyMMdd_HHmmss")
                End If
                
                If m.Fills.Exists(Fill.BrokerID) = False Then
                    If m.PositionFills.Exists(Fill.BrokerID) = False Then
                        m.PositionFills.Add Fill, Fill.BrokerID
                    End If
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cBrokerOec.HandlePositionFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleSymbols
'' Description: Handle a list of symbols enabled for market data from OEC
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleSymbols(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    Dim brokerMessage As cBrokerMessage ' Broker message
    Dim strGenesisBase As String        ' Genesis base symbol
    Dim lPos As Long                    ' Position in the array
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim astrMessage As cGdArray         ' Message broken out into fields
    Dim strFileName As String           ' Filename for the LKSyms file
    Dim astrSymbol As cGdArray          ' Symbol information
    Dim strOecBase As String            ' OEC base symbol
    Dim bIsFuture As Boolean            ' Is the symbol a future?
    Dim strFront As String              ' Front month
    Dim astrToAsk As cGdArray           ' Array of symbols to ask for
    Dim astrSizes As cGdArray           ' Array of contract sizes
    
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        If UCase(strFirstField) = "BEGIN" Then
            m.astrEnabledSymbols.Clear
        ElseIf UCase(strFirstField) = "END" Then
            m.Broker.DumpDebug vbTab & "Enabled Symbols = '" & m.astrEnabledSymbols.JoinFields(",") & "'"
            
            If m.astrEnabledSymbols.BinarySearch("AP") Then
                g.Broker.SfeAllowed(m.Broker.Broker) = Int(CurrentTime)
            Else
                g.Broker.SfeAllowed(m.Broker.Broker) = kNullData
            End If
            
            strFileName = LksFile
            
            m.astrEnabledSymbols.Serialize strFileName, True
            LKSCRC = CalcFileCrc(strFileName)
            
            SaveContractSizes
        Else
            Set astrMessage = New cGdArray
            astrMessage.SplitFields strMessage, vbTab
            
            Set astrToAsk = New cGdArray
            astrToAsk.SplitFields "ND3,QH,QU,SP3,XK,XY,ZD,ZG,ZI,ZQ", ","
                        
            For lIndex = 0 To astrMessage.Size - 1
                Set brokerMessage = New cBrokerMessage
                brokerMessage.FromString astrMessage(lIndex)
                
                Set astrSymbol = New cGdArray
                astrSymbol.SplitFields brokerMessage("Symbol"), ";"
                
                Set astrSizes = New cGdArray
                astrSizes.SplitFields astrSymbol(3), ","
                
                bIsFuture = False
                Select Case UCase(astrSymbol(2))
                    Case "FOREX"
                        strOecBase = astrSymbol(0)
                    
                    Case "FUTURE"
                        bIsFuture = True
                        strOecBase = astrSymbol(0)
                    
                    Case "OPTION"
                        strOecBase = "O:" & astrSymbol(0)
                End Select
                
                strGenesisBase = Replace(GenesisBaseSymbol(strOecBase), "O:", "")
                If Len(strGenesisBase) > 0 Then
                    If m.astrEnabledSymbols.BinarySearch(strGenesisBase, lPos) = False Then
                        m.astrEnabledSymbols.Add strGenesisBase, lPos
                    End If
                    
                    If UCase(astrSymbol(2)) = "FOREX" Then
                        If astrSizes.Size > 0 Then
                            For lIndex2 = 0 To astrSizes.Size - 1
                                ContractSize(strGenesisBase, Parse(astrSizes(lIndex2), "|", 1)) = CLng(Val(Parse(astrSizes(lIndex2), "|", 2)))
                            Next lIndex2
                        End If
                    End If
                End If
            Next lIndex
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerOec.HandleSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisBaseSymbol
'' Description: Determine the Genesis base symbol for the given OEC base symbol
'' Inputs:      Oec Base Symbol, For Position?, Leg Gap (out), Broker Symbol (out)
'' Returns:     Genesis Base Symbol ( Blank if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisBaseSymbol(ByVal strOecBaseSymbol As String, Optional ByVal bForPosition As Boolean = False, Optional strLegGap As String = "", Optional BrokerSym As cBrokerSymbol) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    Set BrokerSym = Nothing
    
    If Len(strOecBaseSymbol) > 0 Then
        If bForPosition Then
            Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol(strOecBaseSymbol)
        Else
            Set BrokerSym = m.Symbols.GetByBrokerSymbol(strOecBaseSymbol)
        End If
        
        If Not BrokerSym Is Nothing Then
            strReturn = BrokerSym.GenesisBase
        End If
    End If
    
    GenesisBaseSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerOec.GenesisBaseSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadContractSizes
'' Description: Load the contract sizes from the file if it exists
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadContractSizes()
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' Array from the file
    Dim lIndex As Long                  ' Index into a for loop

    If FileExist(m.strOcsFile) Then
        Set astrFile = New cGdArray
        astrFile.Serialize m.strOcsFile, False
        
        For lIndex = 0 To astrFile.Size - 1
            m.ContractSizes.Add CLng(Val(Parse(astrFile(lIndex), vbTab, 2))), Parse(astrFile(lIndex), vbTab, 1)
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerOec.LoadContractSizes"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveContractSizes
'' Description: Save the contract sizes to a file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveContractSizes()
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' Array from the file
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrFile = New cGdArray
    astrFile.Create eGDARRAY_Strings
    
    For lIndex = 1 To m.ContractSizes.Count
        astrFile.Add m.ContractSizes.Key(lIndex) & vbTab & Str(m.ContractSizes(lIndex))
    Next lIndex
    
    astrFile.Serialize m.strOcsFile, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerOec.SaveContractSizes"
    
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
    Set m.Fills = New cGdTree
    Set m.PositionFills = New cGdTree
    Set m.ContractSizes = New cGdTree
    
    Set m.astrEnabledSymbols = New cGdArray
    m.astrEnabledSymbols.Create eGDARRAY_Strings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerOec.Class_Initialize"
    
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
    Set m.Symbols = Nothing
    Set m.astrEnabledSymbols = Nothing
    Set m.Fills = Nothing
    Set m.PositionFills = Nothing
    Set m.ContractSizes = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerOec.Class_Terminate"
    
End Sub
