VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cOec.cls
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
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Broker As cBroker                   ' Broker base class object
    Symbols As cBrokerSymbols           ' Broker symbol conversion object
    
    strUserName As String               ' User Name to login with
    strPassword As String               ' Password for the login
    strIP As String                     ' IP address of the server to connect to
    strPort As String                   ' Port to connect to on the server
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

Private Property Get ConnectString() As String
    ConnectString = m.strUserName & vbTab & m.strPassword & vbTab & m.strIP & vbTab & m.strPort
End Property
Private Property Get ConnectStringLog() As String
    ConnectStringLog = m.strUserName & vbTab & SetStrToChar(m.strPassword, "*") & vbTab & m.strIP & vbTab & m.strPort
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
        Case eTT_AccountType_Oec
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "Oec"
            m.Broker.IniFile = AddSlash(App.Path) & "Oec.INI"
            m.Broker.BrokerName = "Open E-Cry"
            m.Broker.ProcessName = "GenOec"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\Oec\GenOec.EXE"
            m.Broker.ControlID = "Oec"
            strSymbolFile = AddSlash(App.Path) & "Provided\OecToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\OecCnct.INI"
            
        Case eTT_AccountType_OptionsXpress
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "OptXpress"
            m.Broker.IniFile = AddSlash(App.Path) & "OptXpress.INI"
            m.Broker.BrokerName = "OptionsXpress"
            m.Broker.ProcessName = "GenOptXpress"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\OptXpress\GenOptXpress.EXE"
            m.Broker.ControlID = "OptXpress"
            strSymbolFile = AddSlash(App.Path) & "Provided\OxpToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\OxpCnct.INI"
    
    End Select
    
    m.Broker.CopyProcess AddSlash(App.Path) & "..\Brokers\GenOec.EXE"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOec.Init"

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
    RaiseError "cOec.ShowConfig"
    
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

    ShowBrokerConnectionInfo = frmOecLogin.ShowMe(m.Broker.Broker, strUserName, True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOec.ShowBrokerConnectionInfo"
    
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
            bReturn = frmOecLogin.ShowMe(m.Broker.Broker, strNewUserName)
            m.Broker.AskedPassword = bReturn
        End If
    Else
        If ((strNewUserName <> m.strUserName) And (Len(strNewUserName) > 0)) Or (bSwitching = True) Then
            bReturn = frmOecLogin.ShowMe(m.Broker.Broker, strNewUserName, , bSwitching)
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
    RaiseError "cOec.ShowLogin"
    
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
    m.Broker.SendBrokerMessage eGDBrokerMessageType_Connect, ConnectString, , ConnectStringLog

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOec.SendConnect"
    
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
    RaiseError "cOec.SendDisconnect"
    
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
    RaiseError "cOec.UnloadApp"
    
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
            Case "F"
                Set BrokerSym = m.Symbols.GetByGenesisSymbol(Parse(strGenesisSymbol, "-", 1))
                If Not BrokerSym Is Nothing Then
                    strReturn = BrokerSym.BrokerBase & BrokerSym.ContractToMY(Parse(strGenesisSymbol, "-", 2))
                End If
                
            Case "FO"
            
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
        End Select
    End If
    
    BrokerSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOec.BrokerSymbol"
    
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
    Dim OptSym As New cOptionSymbol     ' Option symbol object
    
    strReturn = ""
    Select Case UCase(strBrokerExchange)
        Case "CBL", "CBOT", "CME", "COMEX", "EUREX", "HKF", "ICE", "KCBT", "LIFFE", "MGE", "NLI", "NYMEX", "SFE", "SGX"
            If UCase(Left(strBrokerSymbol, 1)) = "O" Then
            Else
                If InStr(strBrokerSymbol, " ") = 0 Then
                    If bPosition Then
                        Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol(Left(strBrokerSymbol, Len(strBrokerSymbol) - 2), strBrokerExchange)
                    Else
                        Set BrokerSym = m.Symbols.GetByBrokerSymbol(Left(strBrokerSymbol, Len(strBrokerSymbol) - 2), strBrokerExchange)
                    End If
                    If Not BrokerSym Is Nothing Then
                        strReturn = BrokerSym.GenesisBase & "-" & BrokerSym.ContractFromMY(Right(strBrokerSymbol, 2))
                    End If
                End If
            End If
            
    End Select
    
    GenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOec.GenesisSymbol"
    
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
    End Select
    
    BrokerOrderType = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOec.BrokerOrderType"
    
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
    RaiseError "cOec.AddOrder"
    
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
    RaiseError "cOec.AmendOrder"
    
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
    RaiseError "cOec.CancelOrder"
    
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
    RaiseError "cOec.ParkOrder"
    
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
    RaiseError "cOec.Refresh"
    
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
    RaiseError "cOec.GetAccounts"
    
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
    RaiseError "cOec.GetOrders"
    
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
    RaiseError "cOec.GetFills"
    
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
    RaiseError "cOec.GetPositions"
    
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
    RaiseError "cOec.NextGenesisID"
    
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
    RaiseError "cOec.StartingGenesisIdForOptNav"
    
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
    RaiseError "cOec.IsTradeableSymbol"
    
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
    RaiseError "cOec.OrderTypeAllowed"
    
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
    RaiseError "cOec.TimeInForceAllowed"
    
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
    RaiseError "cOec.DateToDouble"
    
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
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOec.HandleMessage"
    
End Sub

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

    bReturn = False
    If Len(strMessage) > 0 Then
        If (UCase(strMessage) <> "BEGIN") And (UCase(strMessage) <> "END") Then
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
                .Password = m.strPassword
                .SecTypeMask = 31
                .StartingBalance = 0#
                .StartingDate = Date
                .UserName = m.strUserName
            End With
            
            bReturn = True
        End If
    End If
    
    AccountFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOec.AccountFromInfo"
    
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
    Dim astrOrder As cGdArray           ' Order information split out from the message
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim lQuantityFilled As Long         ' Quantity of the order that has filled
    Dim Bars As New cGdBars             ' Bars object

    bReturn = False
    If Len(strMessage) > 0 Then
        If (UCase(strMessage) <> "BEGIN") And (UCase(strMessage) <> "END") Then
            Set astrOrder = New cGdArray
            astrOrder.SplitFields strMessage, vbTab
            
            strGenesisSymbol = GenesisSymbol(astrOrder(3), astrOrder(4), BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                SetBarProperties Bars, strGenesisSymbol
                
                Set Order = New cPtOrder
                With Order
                    .AccountID = g.Broker.AccountIDForNumber(astrOrder(2))
                    .BrokerID = Parse(astrOrder(1), ";", 2)
                    .Buy = (Left(UCase(astrOrder(6)), 3) = "BUY")
                    .ExchangeID = Parse(astrOrder(1), ";", 1)
                    .GenesisOrderID = astrOrder(0)
                    .LimitPrice = GenesisPrice(astrOrder(9), BrokerSym)
                    .Message = astrOrder(15)
                    .OrderType = GenesisOrderType(astrOrder(5))
                    .Quantity = CLng(Val(astrOrder(7)))
                    .Status = GenesisOrderStatus(astrOrder(12))
                    .StatusDate = DateToDouble(astrOrder(13))
                    .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.StatusDate, m.Broker.Broker, strGenesisSymbol, False))
                    .StopPrice = GenesisPrice(astrOrder(8), BrokerSym)
                    .SymbolOrSymbolID = strGenesisSymbol
                    lQuantityFilled = CLng(Val(astrOrder(14)))
                    
                    If .ExchangeID <> .BrokerID Then
                        .PreviousBrokerID = .ExchangeID
                    End If
                    
                    Select Case UCase(astrOrder(10))
                        Case "DAY"
                            .Expiration = .SessionDate * -1&
                        Case "GTC"
                            .Expiration = 0
                        Case "GTD"
                            .Expiration = CLng(Val(astrOrder(11)))
                    End Select
                    
                    If lQuantityFilled > 0 Then
                        If (.Status = eTT_OrderStatus_Working) Then
                            .Status = eTT_OrderStatus_Partial
                        ElseIf (.Status = eTT_OrderStatus_Cancelled) Then
                            .Status = eTT_OrderStatus_BalCancelled
                        End If
                    End If
                End With
                
                bReturn = True
            End If
        End If
    End If
    
    OrderFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOec.OrderFromInfo"
    
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
    Dim astrFill As cGdArray            ' Fill information split out from the message
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim Bars As New cGdBars             ' Bars object

    bReturn = False
    If Len(strMessage) > 0 Then
        If (UCase(strMessage) <> "BEGIN") And (UCase(strMessage) <> "END") Then
            Set astrFill = New cGdArray
            astrFill.SplitFields strMessage, vbTab
            
            strGenesisSymbol = GenesisSymbol(astrFill(3), astrFill(4), BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                SetBarProperties Bars, strGenesisSymbol
                
                Set Fill = New cPtFill
                With Fill
                    .AccountID = g.Broker.AccountIDForNumber(astrFill(2))
                    .BrokerID = astrFill(0)
                    .BrokerOrderID = astrFill(1)
                    .Buy = (Left(UCase(astrFill(5)), 3) = "BUY")
                    .Fees = Val(astrFill(10))
                    .FillDate = DateToDouble(astrFill(8))
                    .Price = GenesisPrice(astrFill(7), BrokerSym)
                    .Quantity = CLng(Val(astrFill(6)))
                    .SessionDate = Bars.SessionDateForTradeTime(.FillDateConverted(""))
                    .SymbolOrSymbolID = strGenesisSymbol
                End With
                
                strGenesisOrderID = ""
                
                bReturn = True
            End If
        End If
    End If
    
    FillFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOec.FillFromInfo"
    
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
    Dim astrPosition As cGdArray        ' Position information split out from the message
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim astrFills As cGdArray           ' Array of fill information
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object
    Dim lPositionDate As Long           ' Date of the position

    bReturn = False
    If Len(strMessage) > 0 Then
        If (UCase(strMessage) <> "BEGIN") And (UCase(strMessage) <> "END") Then
            Set astrPosition = New cGdArray
            astrPosition.SplitFields strMessage, vbTab
            
            strGenesisSymbol = GenesisSymbol(astrPosition(1), astrPosition(2), BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                Set astrFills = New cGdArray
                astrFills.SplitFields astrPosition(7)
                
                lPositionDate = CLng(Val(astrPosition(3)))
            
                Set Position = New cPtPosition
                With Position
                    .AccountID = g.Broker.AccountIDForNumber(astrPosition(0))
                    .SymbolOrSymbolID = strGenesisSymbol
                    .CurrentPosition = CLng(Val(astrPosition(4))) - CLng(Val(astrPosition(5)))
                    .CarriedPosition = CLng(Val(astrPosition(6)))
                    
                    For lIndex = 0 To astrFills.Size - 1
                        Set Fill = New cPtFill
                        If FillFromInfo(Replace(astrFills(lIndex), ",", vbTab), Fill) Then
                            If Fill.SessionDate < lPositionDate Then
                                .CarriedFills.Add Fill
                            Else
                                .CurrentFills.Add Fill
                            End If
                        End If
                    Next lIndex
                End With
                
                bReturn = True
            End If
        End If
    End If
    
    PositionFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOec.PositionFromInfo"
    
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
    RaiseError "cOec.GenesisSymbolList"
    
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
    RaiseError "cOec.TimeZone"
    
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
    RaiseError "cOec.JustConnected"
    
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
    RaiseError "cOec.BrokerPrice"
    
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
    RaiseError "cOec.GenesisPrice"
    
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
    End Select
    
    GenesisOrderType = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOec.GenesisOrderType"
    
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
    RaiseError "cOec.GenesisOrderStatus"
    
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
    RaiseError "cOec.BrokerBuySell"
    
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
    RaiseError "cOec.BrokerTIF"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerOrderString
'' Description: Build an order string to send to the broker
'' Inputs:      Order, New Order?
'' Returns:     Broker Order String
''
'' Fields:      GenesisID, BrokerID, Account, Symbol, Exchange, Order Type,
''              Buy/Sell, Quantity, Stop, Limit, TIF, Expiration
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerOrderString(ByVal Order As cPtOrder, ByVal bNewOrder As Boolean) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrOrder As cGdArray           ' Array of order information to send to broker
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim strBrokerSymbol As String       ' Broker symbol
    
    strReturn = ""
    strBrokerSymbol = BrokerSymbol(Order.Symbol, BrokerSym)
    
    If (Len(strBrokerSymbol) > 0) And (Not BrokerSym Is Nothing) Then
        Set astrOrder = New cGdArray
        astrOrder(0) = Order.GenesisOrderID
        If bNewOrder Then
            astrOrder(1) = ""
        Else
            astrOrder(1) = Order.ExchangeID & ";" & Order.BrokerID
        End If
        astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(3) = BrokerSymbol(Order.Symbol, BrokerSym)
        astrOrder(4) = BrokerSym.BrokerExchange
        astrOrder(5) = BrokerOrderType(Order.OrderType)
        astrOrder(6) = BrokerBuySell(Order.Buy)
        astrOrder(7) = Str(Order.Quantity)
        astrOrder(8) = BrokerPrice(Order.StopPrice, BrokerSym)
        astrOrder(9) = BrokerPrice(Order.LimitPrice, BrokerSym)
        astrOrder(10) = BrokerTIF(Order.TimeInForce)
        If (Order.Expiration <= 0) Then
            astrOrder(11) = ""
        Else
            astrOrder(11) = Format(Order.Expiration, "YYYYMMDD")
        End If
        
        strReturn = astrOrder.JoinFields(vbTab)
    End If
    
    BrokerOrderString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOec.BrokerOrderString"
    
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

    Dim astrInfo As New cGdArray        ' Array of split out information
    Dim Order As cPtOrder               ' Order object
    Dim strTicketNumber As String       ' Current ticket number for the order ID
    Dim PrevOrder As cPtOrder           ' Previous order
    
    If Len(strMessage) > 0 Then
        astrInfo.SplitFields strMessage, vbTab
        
        Set Order = m.Broker.LoadOrder(astrInfo(0), astrInfo(1), astrInfo(2))
        If Order Is Nothing Then
            strTicketNumber = TicketNumberFromID(astrInfo(2))
            If Len(strTicketNumber) > 0 Then
                Set Order = m.Broker.LoadOrder(astrInfo(0), astrInfo(1), strTicketNumber)
            End If
        End If
        If Not Order Is Nothing Then
            m.Broker.DumpDebug "Order with BrokerID = '" & Order.BrokerID & "' and GenesisID = '" & Order.GenesisOrderID & "' exists"
            
            Set PrevOrder = New cPtOrder
            PrevOrder.LoadForBrokerID Order.BrokerID
            Order.PreviousBrokerID = Order.BrokerID
            Order.PreviousOrder = PrevOrder
            Order.BrokerID = astrInfo(3)
            Order.Save
            
            m.Broker.DumpDebug Order.OrderText(True, True, True) & ": " & OrderStatus(Order.Status) & " --> Broker Order Changed from '" & Order.PreviousBrokerID & "'"
            
            OrderCallback Order
            m.Broker.BrokerInfo.AddOrder Order, False
            SendOrderToOptionNav Order, False
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOec.HandleOrderIdChanged"
    
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
    RaiseError "cOec.TicketNumberFromID"

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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOec.Class_Initialize"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOec.Class_Terminate"
    
End Sub
