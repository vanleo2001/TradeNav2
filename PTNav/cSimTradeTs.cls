VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSimTradeTs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cSimTradeTs.cls
'' Description: Routines for handling simulated trading through the Trade Server
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/23/2011   DAJ         Don't try to auto connect if no accounts
'' 06/23/2011   DAJ         Don't try to connect to Salmon if no TradeServer.IP file
'' 06/24/2011   DAJ         Show message if StartSalmon fails, Added NextAccount, fix initial question
'' 06/24/2011   DAJ         Don't allow trading indexes or PFG Forex
'' 06/29/2011   DAJ         Include today's closed orders from DB for refresh
'' 07/28/2011   DAJ         Don't ask for connection upon connect to streaming
'' 08/11/2011   DAJ         Allow non-PFG forex trading for the trade server
'' 08/15/2011   DAJ         Added GetPositionForSymbol stuff
'' 10/21/2011   DAJ         If an order comes in with broker ID = "0", clear it
'' 08/28/2012   DAJ         CarriedFillFromInfo, First Field Begin/End
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 01/24/2013   DAJ         Don't allow Futures Spreads or Non-Genesis Forex for Trade Server
'' 11/14/2014   DAJ         Added support for MIT, On-Close, and On-Open orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private WithEvents m_txtCallback As ctlUniTextBoxXP ' TextBox 'RH changed
Attribute m_txtCallback.VB_VarHelpID = -1
Private WithEvents m_tmrSalmon As Timer
Attribute m_tmrSalmon.VB_VarHelpID = -1

Private Type mPrivate
    Broker As cBroker                   ' Broker base class object
    SimTradeBase As cSimTradeBase       ' Simulated trading base class object

    strTsLogPath As String              ' Trade Server log path
    bTradeServerInitCalled As Boolean   ' Has InitTradeServer been called?
    strDataServiceID As String          ' Customer/Data Service ID
    strMachineID As String              ' Machine ID
    strError As String                  ' Last error returned from DLL

    astrSimTradeFiles As cGdArray       ' Sim Trade files back from the server
    lRequestNum As Long                 ' Request Number
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

Private Property Get TableItem(tbl As cGdTable, ByVal strFieldName As String, ByVal lRecord As Long) As String
    TableItem = tbl.Item(tbl.FieldNum(strFieldName), lRecord)
End Property
Private Property Let TableItem(tbl As cGdTable, ByVal strFieldName As String, ByVal lRecord As Long, ByVal strValue As String)
    tbl.Item(tbl.FieldNum(strFieldName), lRecord) = strValue
End Property

Public Property Get UseSalmon() As Boolean
    UseSalmon = Not FileExist(AddSlash(App.Path) & "TsGclient.FLG")
End Property

Public Property Get InPath() As String
    If UseSalmon Then
        InPath = AddSlash(App.Path) & "TradeServer\In"
    Else
        InPath = AddSlash(App.Path) & "SimTrade\In"
    End If
End Property

Public Property Get OutPath() As String
    If UseSalmon Then
        OutPath = AddSlash(App.Path) & "TradeServer\Out"
    Else
        OutPath = AddSlash(App.Path) & "SimTrade\Out"
    End If
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

'RH changed Textbox arg Public Sub Init(ByVal nBroker As eTT_AccountType, ByVal txtCallback As TextBox, ByVal tmrSalmon As Timer)
Public Sub Init(ByVal nBroker As eTT_AccountType, ByVal txtCallback As ctlUniTextBoxXP, ByVal tmrSalmon As Timer)
On Error GoTo ErrSection:

    Select Case nBroker
        Case eTT_AccountType_SimBroker
            m.Broker.Init Me, eTT_AccountType_SimBroker, AddSlash(App.Path) & "TradeServer"
            m.Broker.IniFile = AddSlash(App.Path) & "TradeServer.INI"
            m.Broker.BrokerName = "Genesis SimBroker"
            m.Broker.ProcessName = ""
            m.Broker.ProcessPath = ""
            m.Broker.ControlID = ""
            m.Broker.ConnectIni = ""
    
    End Select
    
    Set m_txtCallback = txtCallback
    Set m_tmrSalmon = tmrSalmon

    m.strTsLogPath = AddSlash(App.Path) & "TradeServer"

    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Connect) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Disconnect) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccounts) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccountsMode) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ConnectInfo) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ChangePassword) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Refresh) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewActivity) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_BrokerView) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewOnline) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_VerifyPositions) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_AccountDetails) = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.Init"
Resume
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
    RaiseError "cSimTradeTs.ShowConfig"
    
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

    m.strDataServiceID = Str(RI_GetDataServiceID)
    m.strMachineID = RI_GetMachineID

    ShowBrokerConnectionInfo = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.ShowBrokerConnectionInfo"
    
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

    Dim strReturn As String             ' Return value from the dialog
    
    If g.RealTime.ConnectionStatus = eGDConnectionStatus_Connecting Then
        If m.Broker.BrokerInfo.Accounts.Count > 0 Then
            'strReturn = InfBox("Would you like to connect to|" & m.Broker.BrokerName & "?", "?", "+Yes|-No", m.Broker.BrokerName)
            strReturn = "Y"
        Else
            strReturn = "N"
        End If
    Else
        strReturn = "Y"
    End If
    
    If strReturn = "Y" Then
        m.strDataServiceID = Str(RI_GetDataServiceID)
        m.strMachineID = RI_GetMachineID
    End If

    ShowLogin = (strReturn = "Y")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.ShowLogin"
    
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

    m.Broker.HandleConnectionInfo eGDConnectionStatus_Connecting, "", ""
    
    InitTradeServer
    If UseSalmon Then
        If Not ConnectSalmon Then
            m.Broker.HandleConnectionInfo eGDConnectionStatus_Disconnected, "", ""
        End If
    Else
        m.Broker.HandleConnectionInfo eGDConnectionStatus_Connected, "", ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.SendConnect"
    
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

    m.Broker.HandleConnectionInfo eGDConnectionStatus_Disconnecting, "", ""
    
    ResetTradeServer
    If UseSalmon Then
        DisconnectSalmon
    End If
    
    m.Broker.HandleConnectionInfo eGDConnectionStatus_Disconnected, "", ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.SendDisconnect"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.UnloadApp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerSymbol
'' Description: Attempt to convert the Genesis symbol to the broker symbology
'' Inputs:      Genesis Symbol
'' Returns:     Broker Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerSymbol(ByVal strGenesisSymbol As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim bAllowed As Boolean             ' Is the symbol going to be allowed?

    strReturn = ""
    
    If SecurityType(strGenesisSymbol, True) = "I" Then
        bAllowed = ((IsForex(strGenesisSymbol) = True) And (InStr(strGenesisSymbol, "@") = 0))
    Else
        bAllowed = (IsSpreadSymbol(strGenesisSymbol) = False)
    End If
    
    If bAllowed Then
        strReturn = strGenesisSymbol
    End If
    
    BrokerSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.BrokerSymbol"
    
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

    GenesisSymbol = strBrokerSymbol

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.GenesisSymbol"
    
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
    
    Select Case nGenesisOrderType
        Case eTT_OrderType_Market
            strReturn = "MARKET"
        Case eTT_OrderType_Stop
            strReturn = "STOP"
        Case eTT_OrderType_Limit
            strReturn = "LIMIT"
        Case eTT_OrderType_StopWithLimit
            strReturn = "STOP LIMIT"
        Case eTT_OrderType_MarketOnClose
            strReturn = "MOC"
        Case eTT_OrderType_StopCloseOnly
            strReturn = "SCO"
        Case eTT_OrderType_LimitCloseOnly
            strReturn = "LCO"
        Case eTT_OrderType_StopWithLimitCloseOnly
            strReturn = "SWLCO"
    End Select
    
    BrokerOrderType = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.BrokerOrderType"
    
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

    bReturn = False
    If (AddTradeServerOrder(Order) = 0) Then
        Order.ChangeOrderStatus eTT_OrderStatus_Sent, m.Broker.BrokerInfo
        
        CheckSimTradeStatus "", "Sending New " & m.Broker.BrokerName & " Order"
        bReturn = True
    End If
    
    AddOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.AddOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddMultipleOrders
'' Description: Add multiple orders to the simulated trading interface
'' Inputs:      Orders to add
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AddMultipleOrders(ByVal Orders As cGdTree) As Boolean
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrders As New cGdTree         ' Trade Server orders
    
    For lIndex = 1 To Orders.Count
        m.Broker.DumpDebug "AddMultipleOrders(TradeServer): " & Orders(lIndex).OrderText & " (" & Orders(lIndex).GenesisOrderID & ")"
        Orders(lIndex).ChangeOrderStatus eTT_OrderStatus_Sent, m.Broker.BrokerInfo
        tsOrders.Add Orders(lIndex)
    Next lIndex
    
    If tsOrders.Count > 0 Then
        lReturn = AddMultipleTradeServerOrders(tsOrders)
        If lReturn = 0 Then
            CheckSimTradeStatus "", "Sending New " & m.Broker.BrokerName & " Orders"
        End If
    End If
    
    AddMultipleOrders = (lReturn = 0)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.AddMultipleOrders"
    
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
    If (AmendTradeServerOrder(Order) = 0) Then
        Order.ChangeOrderStatus eTT_OrderStatus_AmendPending, m.Broker.BrokerInfo
        
        CheckSimTradeStatus "", "Modifying " & m.Broker.BrokerName & " Order"
        bReturn = True
    End If
    
    AmendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.AmendOrder"
    
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
    If (CancelTradeServerOrder(Order) = 0) Then
        Order.ChangeOrderStatus eTT_OrderStatus_CancelPending, m.Broker.BrokerInfo
        
        CheckSimTradeStatus "", "Cancelling " & m.Broker.BrokerName & " Order"
        bReturn = True
    End If
    
    CancelOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.CancelOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelMultipleOrders
'' Description: Send multiple Cancel Orders to the simulated trading interface
'' Inputs:      Orders to Cancel
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelMultipleOrders(ByVal Orders As cGdTree) As Boolean
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrders As New cGdTree         ' Trade Server orders
    
    For lIndex = 1 To Orders.Count
        Orders(lIndex).ChangeOrderStatus eTT_OrderStatus_CancelPending, m.Broker.BrokerInfo
        tsOrders.Add Orders(lIndex)
    Next lIndex
    
    If tsOrders.Count > 0 Then
        lReturn = CancelMultipleTradeServerOrders(tsOrders)
        If lReturn = 0 Then
            CheckSimTradeStatus "", "Cancelling " & m.Broker.BrokerName & " Orders"
        End If
    End If
    
    CancelMultipleOrders = (lReturn = 0)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.CancelMultipleOrders"
    
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
    If (CancelTradeServerOrder(Order) = 0) Then
        Order.ChangeOrderStatus eTT_OrderStatus_ParkPending, m.Broker.BrokerInfo
        
        CheckSimTradeStatus "", "Parking " & m.Broker.BrokerName & " Order"
        bReturn = True
    End If
    
    ParkOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.ParkOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenExpiredPositions
'' Description: Flattent any expired positions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FlattenExpiredPositions()
On Error GoTo ErrSection:

    m.SimTradeBase.FlattenExpiredPositions m.Broker

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.FlattenExpiredPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenExpiredPosition
'' Description: Flattent the expired position for the given account and symbol
'' Inputs:      Account, Symbol, Auto Trade Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FlattenExpiredPosition(ByVal vAccountNumberOrID As Variant, ByVal strSymbol As String, ByVal lAtID As Long, Optional ByVal lExpirationDate As Long = kNullData)
On Error GoTo ErrSection:

    m.SimTradeBase.FlattenExpiredPosition m.Broker, vAccountNumberOrID, strSymbol, lAtID, lExpirationDate

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSimTradeTs.FlattenExpiredPosition"
    
End Sub

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
    RaiseError "cSimTradeTs.Refresh"
    
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

    m.SimTradeBase.GetAccounts m.Broker

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.GetAccounts"
    
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

    GetTradeServerOrders

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.GetOrders"
    
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

    m.SimTradeBase.GetFills m.Broker, strAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.GetFills"
    
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

    m.SimTradeBase.GetPositions m.Broker, strAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositionForSymbol
'' Description: Get the position for the given symbol/account and send
'' Inputs:      Account, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPositionForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    m.SimTradeBase.GetPositionForSymbol m.Broker, vAccountNumberOrID, vSymbolOrSymbolID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.GetPositionForSymbol"
    
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

    NextGenesisID = m.SimTradeBase.NextGenesisID(m.Broker.IniFile)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.NextGenesisID"
    
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

    StartingGenesisIdForOptNav = m.SimTradeBase.StartingGenesisIdForOptNav

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.StartingGenesisIdForOptNav"
    
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
    RaiseError "cSimTradeTs.IsTradeableSymbol"
    
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
    
    bReturn = (Len(BrokerSymbol(strGenesisSymbol)) > 0)
    If bReturn Then
    End If
    
    IsEnabledSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.IsEnabledSymbol"
    
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

    Dim bReturn As Boolean              ' Return value for the function

    Select Case nOrderType
        Case eTT_OrderType_Market
            bReturn = True
        Case eTT_OrderType_Stop
            bReturn = True
        Case eTT_OrderType_Limit
            bReturn = True
        Case eTT_OrderType_StopWithLimit
            bReturn = True
        Case eTT_OrderType_MarketOnClose
            bReturn = False
        Case eTT_OrderType_StopCloseOnly
            bReturn = False
        Case eTT_OrderType_LimitCloseOnly
            bReturn = False
        Case eTT_OrderType_StopWithLimitCloseOnly
            bReturn = False
        Case eTT_OrderType_MIT
            bReturn = False
        Case eTT_OrderType_MarketOnOpen
            bReturn = False
        Case eTT_OrderType_Adjustment
            bReturn = False
    End Select

    OrderTypeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.OrderTypeAllowed"
    
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

    TimeInForceAllowed = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.TimeInForceAllowed"
    
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
    
    dReturn = 0#
    
    DateToDouble = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.DateToDouble"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming message
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDSimTradeMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    m.SimTradeBase.HandleMessage m.Broker, nType, strMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.HandleMessage"
    
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

    AccountFromInfo = m.SimTradeBase.AccountFromInfo(m.Broker.Broker, strMessage, Acct)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.AccountFromInfo"
    
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

    OrderFromInfo = m.SimTradeBase.OrderFromInfo(m.Broker.Broker, strMessage, Order)
    
    ' 10/21/2011 DAJ: When the Trade Server rejects an order, it sends back a broker order
    ' ID of zero.  If you have multiple rejects in a row, this causes the new order not
    ' to be updated (because Trade Navigator maps the zero broker ID to the previous order).  By
    ' clearing it out instead, we will correctly use the Genesis ID instead of the Broker ID...
    If Not Order Is Nothing Then
        If Order.BrokerID = "0" Then
            Order.BrokerID = ""
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.OrderFromInfo"
    
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

    FillFromInfo = m.SimTradeBase.FillFromInfo(m.Broker.Broker, strMessage, Fill, strGenesisOrderID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.FillFromInfo"
    
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

    CarriedFillFromInfo = m.SimTradeBase.FillFromInfo(m.Broker.Broker, strMessage, Fill, strGenesisOrderID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.CarriedFillFromInfo"
    
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

    PositionFromInfo = m.SimTradeBase.PositionFromInfo(m.Broker.Broker, strMessage, Position)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.PositionFromInfo"
    
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

    Set GenesisSymbolList = Nothing

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.GenesisSymbolList"
    
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

    TimeZone = m.SimTradeBase.TimeZone(strSymbol)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.TimeZone"
    
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
    RaiseError "cSimTradeTs.JustConnected"
    
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

    SymbolInformation = m.SimTradeBase.SymbolInformation(strGenesisSymbol)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.SymbolInformation"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextAccount
'' Description: Determine the next Genesis SimBroker account number
'' Inputs:      None
'' Returns:     Next Genesis SimBroker account number
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextAccount() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim rs As Recordset                 ' Recordset into the database
    Dim strAccount As String            ' Account number to use
        
    strReturn = ""
    For lIndex = 1 To 9999
        strAccount = "GEN" & Format(lIndex, "0000")
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountNumber]='" & strAccount & "';", dbOpenDynaset)
        If (rs.BOF And rs.EOF) Then
            strReturn = strAccount
            Exit For
        End If
    Next lIndex
    
    NextAccount = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.NextAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_tmrSalmon_Timer
'' Description: Handle another file as the salmon timer has gone off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_tmrSalmon_Timer()
On Error GoTo ErrSection:

    Static bWorking As Boolean          ' Are we currently working on a request?
    Dim strFile As String               ' File to work on
    Dim strPath As String               ' Path for the files

    If bWorking = False Then
        bWorking = True
    
        If UseSalmon Then
            strPath = InPath
            
            If m.astrSimTradeFiles.Size > 0 Then
                If Len(m.astrSimTradeFiles(0)) > 0 Then
                    strFile = AddSlash(strPath) & m.astrSimTradeFiles(0)
                    
                    If UCase(FileExt(strFile)) = "GZP" Then
                        DumpTradeServer "Unzipping files out of '" & strFile & "'"
                        ZipExecute "U", strFile, strPath, "*.TRD"
                        ZipExecute "U", strFile, strPath, "*.TXT"
                    ElseIf UCase(FileExt(strFile)) = "TXT" Then
                        DumpTradeServer "Handling '" & strFile & "'"
                        
                    End If
                    
                    KillFile strFile
                End If
    
                m.astrSimTradeFiles.Remove 0
                If m.astrSimTradeFiles.Size = 0 Then
                    frmOnlineBroker.tmrTradeServer.Enabled = False
                End If
            End If
        End If
        
        DeleteOrderFiles
        GetReturnedItems
        
        bWorking = False
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.m_tmrSalmon_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_txtCallback_Change
'' Description: Handle a message from a salmon callback
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_txtCallback_Change()
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message that came back in the callback
    Dim strKey As String                ' Key in the string in the text box
    Dim strValue As String              ' Value in the string in the text box

    strMessage = Trim(m_txtCallback.Text)
    If Len(strMessage) > 0 Then
        DumpTradeServer "Salmon Message Received: '" & strMessage & "'"
        strKey = Parse(strMessage, "=", 1)
        strValue = Parse(strMessage, "=", 2)
        
        Select Case strKey
            Case "0"
                m.astrSimTradeFiles.Add strValue
                m_tmrSalmon.Enabled = True
                
            Case "1"
                Select Case UCase(strValue)
                    Case "CONNECTED"
                        'm.Broker.ConnectionStatus = eGDConnectionStatus_Connected
                        m.Broker.HandleConnectionInfo eGDConnectionStatus_Connected, "", ""
                        
                    Case "DISCONNECTED"
                        'm.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected
                        m.Broker.HandleConnectionInfo eGDConnectionStatus_Disconnected, "", ""
                
                End Select
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.m_txtCallback_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpTradeServer
'' Description: Send a string to the log file for the day
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpTradeServer(ByVal strMessage As String)
On Error Resume Next

#If 0 Then

    Dim fh As Integer                   ' File handle to open file with
    fh = FreeFile
    Open AddSlash(m.strTsLogPath) & Format(Now, "YYYYMMDD") & ".LOG" For Append As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ") - " & strMessage
        Close #fh
    End If

#Else

    Static LogFile As cLogFile
    If LogFile Is Nothing Then
        Set LogFile = New cLogFile
        LogFile.OpenFile AddSlash(m.strTsLogPath) & "*.LOG"
    End If
    LogFile.WriteText strMessage
    
#End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderType
'' Description: Convert a Simulated Trading order type to a Genesis Order Type
'' Inputs:      Simulated Trading Order Type
'' Returns:     Genesis Order Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderType(ByVal strSimOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderType        ' Return value for the function

    Select Case UCase(strSimOrderType)
        Case "MARKET"
            nReturn = eTT_OrderType_Market
        Case "STOP"
            nReturn = eTT_OrderType_Stop
        Case "LIMIT"
            nReturn = eTT_OrderType_Limit
        Case "STOP LIMIT"
            nReturn = eTT_OrderType_StopWithLimit
        Case "MOC"
            nReturn = eTT_OrderType_MarketOnClose
        Case "SCO"
            nReturn = eTT_OrderType_StopCloseOnly
        Case "LCO"
            nReturn = eTT_OrderType_LimitCloseOnly
        Case "SWLCO"
            nReturn = eTT_OrderType_StopWithLimitCloseOnly
    End Select
    
    GenesisOrderType = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.GenesisOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderStatus
'' Description: Convert a Simulated Trading Order Status to a Genesis Order Status
'' Inputs:      Simulated Trading Order Status
'' Returns:     Genesis Order Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderStatus(ByVal strSimOrderStatus As String) As eTT_OrderStatus
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderStatus      ' Return value for the function

    Select Case UCase(strSimOrderStatus)
        Case "FILLED"
            nReturn = eTT_OrderStatus_Filled
        Case "WORKING", "AMENDED"
            nReturn = eTT_OrderStatus_Working
        Case "CANCELLED"
            nReturn = eTT_OrderStatus_Cancelled
        'Case "AMENDED"
        '    nReturn = eTT_OrderStatus_Amended
        Case "EXPIRED"
            nReturn = eTT_OrderStatus_Expired
        Case "ERROR"
            nReturn = eTT_OrderStatus_Error
        Case "REJECTED"
            nReturn = eTT_OrderStatus_Rejected
    End Select
    
    GenesisOrderStatus = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.GenesisOrderStatus"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateFromString
'' Description: Convert YYYYMMDD and HHMMSS to a double value
'' Inputs:      Date (YYYYMMDD) and Time (HHMMSS)
'' Returns:     Date/Time as a double
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DateFromString(ByVal strDate As String, ByVal strTime As String) As Double
On Error GoTo ErrSection:

    Dim dTime As Double                 ' Time as a fraction of a day
    
    If Len(strTime) > 0 Then
        dTime = (Val(Mid(strTime, 1, 2)) * 3600) + (Val(Mid(strTime, 3, 2)) * 60) + Val(Mid(strTime, 5, 2))
        dTime = dTime / 86400#
    End If
    
    If Len(strDate) > 0 Then
        DateFromString = JulFromLong(CLng(strDate)) + dTime
    Else
        DateFromString = dTime
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.DateFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteOrderFiles
'' Description: Delete the order files that came back in the DelOrders.TXT
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteOrderFiles()
On Error GoTo ErrSection:

    Dim astrFiles As New cGdArray       ' Array of filenames
    Dim lIndex As Long                  ' Index into a for loop
    Dim strOutPath As String            ' Output path
    Dim strInPath As String             ' Input path
    
    strOutPath = OutPath
    strInPath = InPath
    
    If UseSalmon Then
        astrFiles.FromFile AddSlash(strInPath) & "DelOrders.TXT"
        
        For lIndex = 0 To astrFiles.Size - 1
            If UCase(FileExt(AddSlash(strOutPath) & astrFiles(lIndex))) = "ORD" Then
                KillFile Replace(UCase(AddSlash(strOutPath) & astrFiles(lIndex)), ".ORD", ".SNT"), True
            Else
                KillFile AddSlash(strOutPath) & astrFiles(lIndex), True
            End If
        Next lIndex
        
        KillFile AddSlash(strInPath) & "DelOrders.TXT"
    Else
        astrFiles.FromFile AddSlash(App.Path) & "FTP\DelOrders.TXT"
        
        For lIndex = 0 To astrFiles.Size - 1
            KillFile AddSlash(strOutPath) & astrFiles(lIndex), True
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.DeleteOrderFiles"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitTradeServer
'' Description: Initialize the trade server interface
'' Inputs:      None
'' Returns:     Return code from the init call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InitTradeServer() As Long
On Error GoTo ErrSection:

    Dim astrParms As New cGdArray       ' Array of parameters to the DLL
    Dim tblDataIn As New cGdTable       ' Table of input data to the DLL
    Dim astrResult As New cGdArray      ' Result array from the DLL
    Dim lReturn As Long                 ' Return from the DLL call
    
    If m.bTradeServerInitCalled = False Then
        ' Create the necessary arrays...
        astrParms.Create eGDARRAY_Strings
        astrResult.Create eGDARRAY_Strings
        
        astrParms.Add "Paper"
        astrParms.Add "Init"
        
        With tblDataIn
            .CreateField eGDARRAY_Strings, 0, "Mode"            ' "Simulated"
            .CreateField eGDARRAY_Strings, 1, "CallBackMode"    ' "Window", "Procedure", "Notice", "None", "Drop"
            .CreateField eGDARRAY_Strings, 2, "CallBack"        '
            .CreateField eGDARRAY_Strings, 3, "MessageID"       '
            .CreateField eGDARRAY_Strings, 4, "WorkPath"        '
            .CreateField eGDARRAY_Strings, 5, "DiagnosticFlags" '
            
            .Item(0, 0) = "Simulated"
            .Item(1, 0) = "None"
            .Item(2, 0) = ""
            .Item(3, 0) = ""
            If UseSalmon Then
                .Item(4, 0) = AddSlash(App.Path) & "TradeServer\"
            Else
                .Item(4, 0) = AddSlash(App.Path) & "SimTrade\"
            End If
            .Item(5, 0) = GetIniFileProperty("DiagnosticFlags", "", "Init", m.Broker.IniFile)
            
            m.Broker.DumpDebug vbTab & "Trade Server Message Sent (Init): " & astrParms.JoinFields(vbTab) & "|" & .GetRecord(0, vbTab)
        End With
        
        DumpTradeServer "Trade Navigator version: " & FormatVersion & "." & Str(App.Revision) & " " & DateFormat(FileDate(App.Path & "\" & App.EXEName & ".EXE"), MM_DD_YYYY, HH_MM, AMPM_UPPER)
        lReturn = gntConfigure(astrParms.ArrayHandle, tblDataIn.TableHandle, ByVal 0, astrResult.ArrayHandle)
        DumpTables "Trade Server Message Sent (Init)", lReturn, , astrResult, , astrParms, tblDataIn
        HandleError lReturn, astrResult
        
        m.bTradeServerInitCalled = True
    End If
    
    InitTradeServer = lReturn

ErrExit:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set astrResult = Nothing
    Exit Function
    
ErrSection:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set astrResult = Nothing
    RaiseError "cSimTradeTs.InitTradeServer"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetTradeServer
'' Description: Reset the trade server interface
'' Inputs:      None
'' Returns:     Return code from the Reset call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ResetTradeServer() As Long
On Error GoTo ErrSection:

    Dim astrParms As New cGdArray       ' Array of parameters to the DLL
    Dim tblDataIn As New cGdTable       ' Table of input data to the DLL
    Dim astrResult As New cGdArray      ' Result array from the DLL
    Dim lReturn As Long                 ' Return from the DLL call

    If m.bTradeServerInitCalled = True Then
        ' Create the necessary arrays...
        astrParms.Create eGDARRAY_Strings
        astrResult.Create eGDARRAY_Strings
        
        astrParms.Add "Paper"
        astrParms.Add "Reset"
    
        m.Broker.DumpDebug vbTab & "Trade Server Message Sent (Reset): " & astrParms.JoinFields(vbTab)
        
        lReturn = gntConfigure(astrParms.ArrayHandle, tblDataIn.TableHandle, ByVal 0, astrResult.ArrayHandle)
        DumpTables "Trade Server Message Sent (Reset)", lReturn, , astrResult, , astrParms
        HandleError lReturn, astrResult
        
        m.bTradeServerInitCalled = False
    End If
    
    ResetTradeServer = lReturn

ErrExit:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set astrResult = Nothing
    Exit Function
    
ErrSection:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set astrResult = Nothing
    RaiseError "cSimTradeTs.ResetTradeServer"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddTradeServerOrder
'' Description: Add an order to the simulated trading interface
'' Inputs:      Order to add
'' Returns:     Return value from the Add Order call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddTradeServerOrder(ByVal Order As cPtOrder) As Long
On Error GoTo ErrSection:

    Dim astrParms As New cGdArray       ' Array of parameters to the DLL
    Dim tblDataIn As New cGdTable       ' Table of input data to the DLL
    Dim tblDataOut As New cGdTable      ' Table of output data from the DLL
    Dim astrResult As New cGdArray      ' Result array from the DLL
    Dim lReturn As Long                 ' Return code from the DLL

    ' Create the parameter array...
    With astrParms
        .Create eGDARRAY_Strings
        .Add "Paper"
        .Add "AddOrder"
    End With

    ' Create the result array...
    astrResult.Create eGDARRAY_Strings
    
    ' Create and fill the input table...
    With tblDataIn
        .CreateField eGDARRAY_Strings, 0, "DataServiceID"   ' Text 16
        .CreateField eGDARRAY_Strings, 1, "MachineID"       ' Text 11
        .CreateField eGDARRAY_Strings, 2, "TraderAccount"   ' Text 20
        .CreateField eGDARRAY_Strings, 3, "OrderType"       ' Text 20
        .CreateField eGDARRAY_Strings, 4, "GenesisSymbol"   ' Text 20
        .CreateField eGDARRAY_Strings, 5, "FeedSymbol"      ' Text 20
        .CreateField eGDARRAY_Strings, 6, "BuyOrSell"       ' 'B' or 'S'
        .CreateField eGDARRAY_Strings, 7, "StopPrice"       ' Float as Text 20
        .CreateField eGDARRAY_Strings, 8, "LimitPrice"      ' Float as Text 20
        .CreateField eGDARRAY_Strings, 9, "Lots"            ' Int as Text
        .CreateField eGDARRAY_Strings, 10, "LinkedOrder"    ' Text 20
        .CreateField eGDARRAY_Strings, 11, "XRefPersist"    ' Int as Text
        .CreateField eGDARRAY_Strings, 12, "GoodTillDate"   ' CCYYMMDD
        .CreateField eGDARRAY_Strings, 13, "DateSent"       ' CCYYMMDD
        .CreateField eGDARRAY_Strings, 14, "TimeSent"       ' HHMMSS
        
        .Item(0, 0) = m.strDataServiceID
        .Item(1, 0) = m.strMachineID
        .Item(2, 0) = g.Broker.AccountNumberForID(Order.AccountID)
        .Item(3, 0) = BrokerOrderType(Order.OrderType)
        .Item(4, 0) = Order.Symbol
        If Order.Buy Then
            .Item(6, 0) = "B"
        Else
            .Item(6, 0) = "S"
        End If
        .Item(7, 0) = Str(Order.StopPrice)
        .Item(8, 0) = Str(Order.LimitPrice)
        .Item(9, 0) = Str(Order.Quantity)
        .Item(10, 0) = ""
        .Item(11, 0) = Order.GenesisOrderID
        If Order.Expiration <> 0 Then
            .Item(12, 0) = Format(Abs(Order.Expiration), "YYYYMMDD")
        Else
            .Item(12, 0) = "0"
        End If
        .Item(13, 0) = Format(CurrentTime, "YYYYMMDD")
        .Item(14, 0) = Format(CurrentTime, "HHMMSS")
        m.Broker.DumpDebug vbTab & "Trade Server Message Sent (AddOrder): " & astrParms.JoinFields(vbTab) & "|" & .GetRecord(0, vbTab)
    End With
    
    lReturn = gntTrade(astrParms.ArrayHandle, tblDataIn.TableHandle, tblDataOut.TableHandle, astrResult.ArrayHandle)
    DumpTables "Trade Server Message Sent (AddOrder)", lReturn, , astrResult, tblDataOut, astrParms, tblDataIn
    HandleError lReturn, astrResult
    
    AddTradeServerOrder = lReturn
    
ErrExit:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    Exit Function
    
ErrSection:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    RaiseError "cSimTradeTs.AddTradeServerOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddMultipleTradeServerOrders
'' Description: Add multiple orders to the simulated trading interface
'' Inputs:      Orders to add
'' Returns:     Return value from the Add Order call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddMultipleTradeServerOrders(ByVal Orders As cGdTree) As Long
On Error GoTo ErrSection:

    Dim astrParms As New cGdArray       ' Array of parameters to the DLL
    Dim tblDataIn As New cGdTable       ' Table of input data to the DLL
    Dim tblDataOut As New cGdTable      ' Table of output data from the DLL
    Dim astrResult As New cGdArray      ' Result array from the DLL
    Dim lReturn As Long                 ' Return code from the DLL
    Dim lIndex As Long                  ' Index into a for loop

    ' Create the parameter array...
    With astrParms
        .Create eGDARRAY_Strings
        .Add "Paper"
        .Add "AddOrder"
    End With

    ' Create the result array...
    astrResult.Create eGDARRAY_Strings
    
    ' Create and fill the input table...
    With tblDataIn
        .CreateField eGDARRAY_Strings, 0, "DataServiceID"   ' Text 16
        .CreateField eGDARRAY_Strings, 1, "MachineID"       ' Text 11
        .CreateField eGDARRAY_Strings, 2, "TraderAccount"   ' Text 20
        .CreateField eGDARRAY_Strings, 3, "OrderType"       ' Text 20
        .CreateField eGDARRAY_Strings, 4, "GenesisSymbol"   ' Text 20
        .CreateField eGDARRAY_Strings, 5, "FeedSymbol"      ' Text 20
        .CreateField eGDARRAY_Strings, 6, "BuyOrSell"       ' 'B' or 'S'
        .CreateField eGDARRAY_Strings, 7, "StopPrice"       ' Float as Text 20
        .CreateField eGDARRAY_Strings, 8, "LimitPrice"      ' Float as Text 20
        .CreateField eGDARRAY_Strings, 9, "Lots"            ' Int as Text
        .CreateField eGDARRAY_Strings, 10, "LinkedOrder"    ' Text 20
        .CreateField eGDARRAY_Strings, 11, "XRefPersist"    ' Int as Text
        .CreateField eGDARRAY_Strings, 12, "GoodTillDate"   ' CCYYMMDD
        .CreateField eGDARRAY_Strings, 13, "DateSent"       ' CCYYMMDD
        .CreateField eGDARRAY_Strings, 14, "TimeSent"       ' HHMMSS
        
        .NumRecords = Orders.Count
        
        For lIndex = 1 To Orders.Count
            m.Broker.DumpDebug "AddMultipleOrders(TradeServer): " & Orders(lIndex).OrderText & " (" & Orders(lIndex).GenesisOrderID & ")"

            .Item(0, lIndex - 1) = m.strDataServiceID
            .Item(1, lIndex - 1) = m.strMachineID
            .Item(2, lIndex - 1) = g.Broker.AccountNumberForID(Orders(lIndex).AccountID)
            .Item(3, lIndex - 1) = BrokerOrderType(Orders(lIndex).OrderType)
            .Item(4, lIndex - 1) = Orders(lIndex).Symbol
            If Orders(lIndex).Buy Then .Item(6, lIndex - 1) = "B" Else .Item(6, lIndex - 1) = "S"
            .Item(7, lIndex - 1) = Str(Orders(lIndex).StopPrice)
            .Item(8, lIndex - 1) = Str(Orders(lIndex).LimitPrice)
            .Item(9, lIndex - 1) = Str(Orders(lIndex).Quantity)
            .Item(10, lIndex - 1) = ""
            .Item(11, lIndex - 1) = Orders(lIndex).GenesisOrderID
            If Orders(lIndex).Expiration <> 0 Then
                .Item(12, lIndex - 1) = Format(Abs(Orders(lIndex).Expiration), "YYYYMMDD")
            Else
                .Item(12, lIndex - 1) = "0"
            End If
            .Item(13, lIndex - 1) = Format(CurrentTime, "YYYYMMDD")
            .Item(14, lIndex - 1) = Format(CurrentTime, "HHMMSS")
            m.Broker.DumpDebug vbTab & "Trade Server Message Sent (AddOrder): " & astrParms.JoinFields(vbTab) & "|" & .GetRecord(lIndex - 1, vbTab)
        Next lIndex
    End With
    
    lReturn = gntTrade(astrParms.ArrayHandle, tblDataIn.TableHandle, tblDataOut.TableHandle, astrResult.ArrayHandle)
    DumpTables "Trade Server Message Sent (AddOrder)", lReturn, , astrResult, tblDataOut, astrParms, tblDataIn
    HandleError lReturn, astrResult
    
    AddMultipleTradeServerOrders = lReturn
    
ErrExit:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    Exit Function
    
ErrSection:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    RaiseError "cSimTradeTs.AddMultipleTradeServerOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AmendTradeServerOrder
'' Description: Amend an order to the simulated trading interface
'' Inputs:      Order to amend
'' Returns:     Return value from the Amend Order call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AmendTradeServerOrder(ByVal Order As cPtOrder) As Long
On Error GoTo ErrSection:

    Dim astrParms As New cGdArray       ' Array of parameters to the DLL
    Dim tblDataIn As New cGdTable       ' Table of input data to the DLL
    Dim tblDataOut As New cGdTable      ' Table of output data from the DLL
    Dim astrResult As New cGdArray      ' Result array from the DLL
    Dim lReturn As Long                 ' Return code from the DLL

    ' Create the parameter array...
    With astrParms
        .Create eGDARRAY_Strings
        .Add "Paper"
        .Add "AmendOrder"
    End With

    ' Create the result array...
    astrResult.Create eGDARRAY_Strings
    
    ' Create and fill the input table...
    With tblDataIn
        .CreateField eGDARRAY_Strings, 0, "OrderID"         ' Text 20
        .CreateField eGDARRAY_Strings, 1, "DataServiceID"   ' Text 16
        .CreateField eGDARRAY_Strings, 2, "MachineID"       ' Text 11
        .CreateField eGDARRAY_Strings, 3, "OrderType"       ' Text 20
        .CreateField eGDARRAY_Strings, 4, "TraderAccount"   ' Text 20
        .CreateField eGDARRAY_Strings, 5, "StopPrice"       ' Float as Text 20
        .CreateField eGDARRAY_Strings, 6, "LimitPrice"      ' Float as Text 20
        .CreateField eGDARRAY_Strings, 7, "Lots"            ' Int as Text
        .CreateField eGDARRAY_Strings, 8, "LinkedOrder"     ' Text 20
        .CreateField eGDARRAY_Strings, 9, "GoodTillDate"    ' CCYYMMDD
        .CreateField eGDARRAY_Strings, 10, "XRefPersist"     ' Int as Text
    
        .Item(0, 0) = Order.BrokerID
        .Item(1, 0) = m.strDataServiceID
        .Item(2, 0) = m.strMachineID
        .Item(3, 0) = BrokerOrderType(Order.OrderType)
        .Item(4, 0) = g.Broker.AccountNumberForID(Order.AccountID)
        .Item(5, 0) = Str(Order.StopPrice)
        .Item(6, 0) = Str(Order.LimitPrice)
        .Item(7, 0) = Str(Order.Quantity)
        .Item(8, 0) = ""
        If Order.Expiration <> 0 Then
            .Item(9, 0) = Format(Abs(Order.Expiration), "YYYYMMDD")
        Else
            .Item(9, 0) = "0"
        End If
        .Item(10, 0) = Order.GenesisOrderID
        m.Broker.DumpDebug vbTab & "Trade Server Message Sent (AmendOrder): " & astrParms.JoinFields(vbTab) & "|" & .GetRecord(0, vbTab)
    End With
    
    lReturn = gntTrade(astrParms.ArrayHandle, tblDataIn.TableHandle, tblDataOut.TableHandle, astrResult.ArrayHandle)
    DumpTables "Trade Server Message Sent (AmendOrder)", lReturn, , astrResult, tblDataOut, astrParms, tblDataIn
    HandleError lReturn, astrResult
    
    AmendTradeServerOrder = lReturn
    
ErrExit:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    Exit Function
    
ErrSection:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    RaiseError "cSimTradeTs.AmendTradeServerOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelTradeServerOrder
'' Description: Send a Cancel Order to the simulated trading interface
'' Inputs:      Order
'' Returns:     Return value from the Cancel Order call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CancelTradeServerOrder(ByVal Order As cPtOrder) As Long
On Error GoTo ErrSection:

    Dim astrParms As New cGdArray       ' Array of parameters to the DLL
    Dim tblDataIn As New cGdTable       ' Table of input data to the DLL
    Dim tblDataOut As New cGdTable      ' Table of output data from the DLL
    Dim astrResult As New cGdArray      ' Result array from the DLL
    Dim lReturn As Long                 ' Return code from the DLL

    ' Create the parameter array...
    With astrParms
        .Create eGDARRAY_Strings
        .Add "Paper"
        .Add "CancelOrder"
    End With

    ' Create the result array...
    astrResult.Create eGDARRAY_Strings
    
    ' Create and fill the input table...
    With tblDataIn
        .CreateField eGDARRAY_Strings, 0, "OrderID"         ' Text 20
        .CreateField eGDARRAY_Strings, 1, "DataServiceID"   ' Text 16
        .CreateField eGDARRAY_Strings, 2, "MachineID"       ' Text 11
        .CreateField eGDARRAY_Strings, 3, "TraderAccount"   ' Text 20
        .CreateField eGDARRAY_Strings, 4, "XRefPersist"     ' Int as Text
    
        .Item(0, 0) = Order.BrokerID
        .Item(1, 0) = m.strDataServiceID
        .Item(2, 0) = m.strMachineID
        .Item(3, 0) = g.Broker.AccountNumberForID(Order.AccountID)
        .Item(4, 0) = Order.GenesisOrderID
        m.Broker.DumpDebug vbTab & "Trade Server Message Sent (CancelOrder): " & astrParms.JoinFields(vbTab) & "|" & .GetRecord(0, vbTab)
    End With
    
    lReturn = gntTrade(astrParms.ArrayHandle, tblDataIn.TableHandle, tblDataOut.TableHandle, astrResult.ArrayHandle)
    DumpTables "Trade Server Message Sent (CancelOrder)", lReturn, , astrResult, tblDataOut, astrParms, tblDataIn
    HandleError lReturn, astrResult
    
    CancelTradeServerOrder = lReturn
    
ErrExit:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    Exit Function
    
ErrSection:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    RaiseError "cSimTradeTs.CancelTradeServerOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelMultipleTradeServerOrders
'' Description: Send multiple Cancel Orders to the simulated trading interface
'' Inputs:      Orders to Cancel
'' Returns:     Return value from the Cancel Order call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CancelMultipleTradeServerOrders(ByVal Orders As cGdTree) As Long
On Error GoTo ErrSection:

    Dim astrParms As New cGdArray       ' Array of parameters to the DLL
    Dim tblDataIn As New cGdTable       ' Table of input data to the DLL
    Dim tblDataOut As New cGdTable      ' Table of output data from the DLL
    Dim astrResult As New cGdArray      ' Result array from the DLL
    Dim lReturn As Long                 ' Return code from the DLL
    Dim Order As New cPtOrder           ' Temporary order object
    Dim lIndex As Long                  ' Index into a for loop

    ' Create the parameter array...
    With astrParms
        .Create eGDARRAY_Strings
        .Add "Paper"
        .Add "CancelOrder"
    End With

    ' Create the result array...
    astrResult.Create eGDARRAY_Strings
    
    ' Create and fill the input table...
    With tblDataIn
        .CreateField eGDARRAY_Strings, 0, "OrderID"         ' Text 20
        .CreateField eGDARRAY_Strings, 1, "DataServiceID"   ' Text 16
        .CreateField eGDARRAY_Strings, 2, "MachineID"       ' Text 11
        .CreateField eGDARRAY_Strings, 3, "TraderAccount"   ' Text 20
        .CreateField eGDARRAY_Strings, 4, "XRefPersist"     ' Int as Text
        
        .NumRecords = Orders.Count
        For lIndex = 1 To Orders.Count
            Set Order = Orders(lIndex)
            m.Broker.DumpDebug "CancelMultipleOrders(TradeServer): " & Order.OrderText & " (" & Order.GenesisOrderID & ")"
            
            .Item(0, lIndex - 1) = Str(Order.BrokerID)
            .Item(1, lIndex - 1) = m.strDataServiceID
            .Item(2, lIndex - 1) = m.strMachineID
            .Item(3, lIndex - 1) = g.Broker.AccountNumberForID(Order.AccountID)
            .Item(4, lIndex - 1) = Order.GenesisOrderID
            m.Broker.DumpDebug vbTab & "Trade Server Message Sent (CancelOrder): " & astrParms.JoinFields(vbTab) & "|" & .GetRecord(lIndex - 1, vbTab)
        Next lIndex
    End With
    
    lReturn = gntTrade(astrParms.ArrayHandle, tblDataIn.TableHandle, tblDataOut.TableHandle, astrResult.ArrayHandle)
    DumpTables "Trade Server Message Sent (CancelOrder)", lReturn, , astrResult, tblDataOut, astrParms, tblDataIn
    HandleError lReturn, astrResult
    
    CancelMultipleTradeServerOrders = lReturn

ErrExit:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    Exit Function
    
ErrSection:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    RaiseError "cSimTradeTs.CancelMultipleTradeServerOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelAllTradeServerOrders
'' Description: Send a Cancel All Orders to the simulated trading interface
'' Inputs:      Account Number
'' Returns:     Return value from the Cancel All Orders call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CancelAllTradeServerOrders(ByVal strAccountNumber As String) As Long
On Error GoTo ErrSection:

    Dim astrParms As New cGdArray       ' Array of parameters to the DLL
    Dim tblDataIn As New cGdTable       ' Table of input data to the DLL
    Dim tblDataOut As New cGdTable      ' Table of output data from the DLL
    Dim astrResult As New cGdArray      ' Result array from the DLL
    Dim lReturn As Long                 ' Return code from the DLL

    m.Broker.DumpDebug "CancelAllOrders(TradeServer): " & strAccountNumber

    ' Create the parameter array...
    With astrParms
        .Create eGDARRAY_Strings
        .Add "Paper"
        .Add "CancelAll"
    End With

    ' Create the result array...
    astrResult.Create eGDARRAY_Strings
    
    ' Create and fill the input table...
    With tblDataIn
        .CreateField eGDARRAY_Strings, 0, "DataServiceID"   ' Text 16
        .CreateField eGDARRAY_Strings, 1, "MachineID"       ' Text 11
        .CreateField eGDARRAY_Strings, 2, "TraderAccount"   ' Text 20
    
        .Item(0, 0) = m.strDataServiceID
        .Item(1, 0) = m.strMachineID
        .Item(2, 0) = strAccountNumber
        m.Broker.DumpDebug vbTab & "Trade Server Message Sent (CancelAll): " & astrParms.JoinFields(vbTab) & "|" & .GetRecord(0, vbTab)
    End With
    
    lReturn = gntTrade(astrParms.ArrayHandle, tblDataIn.TableHandle, tblDataOut.TableHandle, astrResult.ArrayHandle)
    DumpTables "Trade Server Message Sent (CancelAll)", lReturn, , astrResult, tblDataOut, astrParms, tblDataIn
    HandleError lReturn, astrResult
    
    CancelAllTradeServerOrders = lReturn
    
ErrExit:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    Exit Function
    
ErrSection:
    Set astrParms = Nothing
    Set tblDataIn = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    RaiseError "cSimTradeTs.CancelAllTradeServerOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTradeServerOrders
'' Description: Get the orders from the Simulated Trading server
'' Inputs:      None
'' Returns:     Return value from the GetInfo call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetTradeServerOrders(Optional ByVal bWithHistory As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim astrParms As New cGdArray       ' Array of parameters to the DLL
    Dim astrDesc As New cGdArray        ' Array of descriptions
    Dim tblDataIn As New cGdTable       ' Table of input data to the DLL
    Dim tblDataOut As New cGdTable      ' Table of output data from the DLL
    Dim astrResult As New cGdArray      ' Result array from the DLL
    Dim lReturn As Long                 ' Return value from the DLL call

    ' Create the necessary arrays...
    astrParms.Create eGDARRAY_Strings
    astrDesc.Create eGDARRAY_Strings
    astrResult.Create eGDARRAY_Strings
    
    ' Fill in the parameters array...
    astrParms.Add "Paper"
    astrParms.Add "Orders"
    If bWithHistory Then
        astrParms.Add "Y"
    Else
        astrParms.Add "N"
    End If
    
    ' Create and fill the input table...
    With tblDataIn
        .CreateField eGDARRAY_Strings, 0, "MachineID"       ' Text 11
        .CreateField eGDARRAY_Strings, 1, "DataServiceID"   ' Text 16
        .CreateField eGDARRAY_Strings, 2, "TraderAccount"   ' Text 20
    
        .Item(0, 0) = m.strMachineID
        .Item(1, 0) = m.strDataServiceID
        .Item(2, 0) = "" 'strAccountNumber
        m.Broker.DumpDebug vbTab & "Trade Server Message Sent (GetOrders): " & astrParms.JoinFields(vbTab) & "|" & .GetRecord(0, vbTab)
    End With
    
    lReturn = gntTrade(astrParms.ArrayHandle, tblDataIn.TableHandle, tblDataOut.TableHandle, astrResult.ArrayHandle)
    DumpTables "Trade Server Message Sent (GetOrders)", lReturn, , astrResult, tblDataOut, astrParms, tblDataIn
    
    If lReturn = 0 Then
        CheckSimTradeStatus "", "Refreshing Orders and Fills"
    Else
        HandleError lReturn, astrResult
    End If
    
    GetTradeServerOrders = lReturn

ErrExit:
    Set astrParms = Nothing
    Set astrDesc = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    Exit Function
    
ErrSection:
    Set astrParms = Nothing
    Set astrDesc = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    RaiseError "cSimTradeTs.GetTradeServerOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetReturnedItems
'' Description: Get information from the Simulated Trading DLL
'' Inputs:      Optional CallBack ID
'' Returns:     Return value of the DLL call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetReturnedItems(Optional ByVal lCBID As Long = 0&) As Long
On Error GoTo ErrSection:

    Dim astrParms As New cGdArray       ' Array of parameters to the DLL
    Dim astrDesc As New cGdArray        ' Array of descriptions
    Dim tblDataOut As New cGdTable      ' Table of output data from the DLL
    Dim astrResult As New cGdArray      ' Result array from the DLL
    Dim lReturn As Long                 ' Return value from the DLL call
    
    ' Create the necessary arrays...
    astrParms.Create eGDARRAY_Strings
    astrDesc.Create eGDARRAY_Strings
    astrResult.Create eGDARRAY_Strings
    
    astrParms.Add "Paper"
    astrParms.Add "CallBack"
    
    If lCBID = 0& Then
        astrParms.Add "Next"
    Else
        astrParms.Add "CBID"
        astrParms.Add Str(lCBID)
    End If
    
    lReturn = 0&
    Do While lReturn = 0
        lReturn = gntGetInfo(astrParms.ArrayHandle, astrDesc.ArrayHandle, tblDataOut.TableHandle, astrResult.ArrayHandle)
        DumpTables "Trade Server Message Received", lReturn, astrDesc, astrResult, tblDataOut
        HandleTradeServerMessage lReturn, astrDesc, tblDataOut, astrResult
        If lCBID <> 0& Then Exit Do
    Loop
    
    GetReturnedItems = lReturn

ErrExit:
    Set astrParms = Nothing
    Set astrDesc = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    Exit Function
    
ErrSection:
    Set astrParms = Nothing
    Set astrDesc = Nothing
    Set tblDataOut = Nothing
    Set astrResult = Nothing
    RaiseError "cSimTradeTs.GetReturnedItems"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleTradeServerMessage
'' Description: Handle the information returned from the DLL
'' Inputs:      Code of the CallBack, Arrays to store/retrieve information
'' Returns:     Return value of the callback
''
'' Order:       Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date, At ID, Session Date, Exchange ID, Status Date,
''              Date Sent, Date Exch Recd, Message
'' Fill:        Broker ID, Genesis ID, Fill ID, Account, Symbol, Fill Date,
''              B/S, Fill Quantity, Fill Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HandleTradeServerMessage(ByVal lCode As Long, astrDesc As cGdArray, tblDataOut As cGdTable, astrResult As cGdArray, Optional ByVal bDumpDebug As Boolean = True) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Order As New cPtOrder           ' Temporary Order object for filling in
    Dim Fill As New cPtFill             ' Temporary Fill object for filling in
    Dim strPersist As String            ' Persist XREF
    Dim nStatus As eTT_OrderStatus      ' Order status
    Dim bFillExisted As Boolean         ' Did the fill we just got already exist?
    Dim bContinue As Boolean            ' Do we want to continue?
    Dim astrToAdd As cGdArray           ' Array of information to join together
    Dim bRefresh As Boolean             ' Was this a refresh?
    Dim astrOrders As cGdArray          ' Array of closed orders
    
    Set astrToAdd = New cGdArray
    astrToAdd.Create eGDARRAY_Strings

    If lCode = 0& Then
        bRefresh = (UCase(astrDesc(1)) = "FILLS")
        
        Select Case UCase(astrDesc(4))
            Case "ORDER"
                If bRefresh Then
                    HandleMessage eGDSimTradeMessageType_RefreshOrder, "BEGIN"
                    Set astrOrders = m.SimTradeBase.GetSnapshotOrdersFromDb(m.Broker.Broker)
                    For lIndex = 0 To astrOrders.Size - 1
                        HandleMessage eGDSimTradeMessageType_RefreshOrder, astrOrders(lIndex)
                    Next lIndex
                End If
                
                For lIndex = 0 To tblDataOut.NumRecords - 1
                    m.Broker.DumpDebug vbTab & "Trade Server Message Received (" & astrDesc(4) & "): " & tblDataOut.GetRecord(lIndex, vbTab)
                    
                    astrToAdd.Clear
                    astrToAdd(0) = TableItem(tblDataOut, "OrderID", lIndex)
                    astrToAdd(1) = TableItem(tblDataOut, "XRefPersist", lIndex)
                    astrToAdd(2) = ""
                    astrToAdd(3) = TableItem(tblDataOut, "TraderAccount", lIndex)
                    astrToAdd(4) = TableItem(tblDataOut, "GenesisSymbol", lIndex)
                    astrToAdd(5) = Str(GenesisOrderType(TableItem(tblDataOut, "OrderType", lIndex)))
                    astrToAdd(6) = Str(CLng((UCase(TableItem(tblDataOut, "BuyOrSell", lIndex)) = "B")))
                    astrToAdd(7) = TableItem(tblDataOut, "Lots", lIndex)
                    astrToAdd(8) = TableItem(tblDataOut, "LimitPrice", lIndex)
                    astrToAdd(9) = TableItem(tblDataOut, "StopPrice", lIndex)
                    astrToAdd(10) = Str(JulFromLong(ValOfText(TableItem(tblDataOut, "GoodTillDate", lIndex))))
                    astrToAdd(11) = Str(GenesisOrderStatus(TableItem(tblDataOut, "Status", lIndex)))
                    astrToAdd(12) = ""
                    If TableItem(tblDataOut, "DateSent", lIndex) <> "-999999" Then
                        astrToAdd(13) = Str(DateFromString(TableItem(tblDataOut, "DateSent", lIndex), TableItem(tblDataOut, "TimeSent", lIndex)))
                    End If
                    astrToAdd(14) = ""
                    astrToAdd(15) = ""
                    astrToAdd(16) = TableItem(tblDataOut, "ExchangeOrderID", lIndex)
                    If TableItem(tblDataOut, "StatusDate", lIndex) <> "-999999" Then
                        astrToAdd(17) = Str(DateFromString(TableItem(tblDataOut, "StatusDate", lIndex), TableItem(tblDataOut, "StatusTime", lIndex)))
                    End If
                    If TableItem(tblDataOut, "DateSent", lIndex) <> "-999999" Then
                        astrToAdd(18) = Str(DateFromString(TableItem(tblDataOut, "DateSent", lIndex), TableItem(tblDataOut, "TimeSent", lIndex)))
                    End If
                    If TableItem(tblDataOut, "DateExchRecd", lIndex) <> "-999999" Then
                        astrToAdd(19) = Str(DateFromString(TableItem(tblDataOut, "DateExchRecd", lIndex), TableItem(tblDataOut, "TimeExchRecd", lIndex)))
                    End If
                    astrToAdd(20) = TableItem(tblDataOut, "NonExecReason", lIndex)
                    
                    If UCase(astrToAdd(20)) <> "NO ORDERS FOUND" Then
                        If bRefresh Then
                            HandleMessage eGDSimTradeMessageType_RefreshOrder, astrToAdd.JoinFields(vbTab)
                        Else
                            HandleMessage eGDSimTradeMessageType_Order, astrToAdd.JoinFields(vbTab)
                        End If
                    End If
                Next lIndex
            
                If bRefresh Then
                    HandleMessage eGDSimTradeMessageType_RefreshOrder, "END"
                End If
                
            Case "FILL"
                For lIndex = 0 To tblDataOut.NumRecords - 1
                    m.Broker.DumpDebug vbTab & "Trade Server Message Received (" & astrDesc(4) & "): " & tblDataOut.GetRecord(lIndex, vbTab)
                    
                    astrToAdd.Clear
                    astrToAdd(0) = TableItem(tblDataOut, "OrderID", lIndex)
                    astrToAdd(1) = TableItem(tblDataOut, "XRefPersist", lIndex)
                    astrToAdd(2) = TableItem(tblDataOut, "FillID", lIndex)
                    astrToAdd(3) = TableItem(tblDataOut, "TraderAccount", lIndex)
                    astrToAdd(4) = TableItem(tblDataOut, "GenesisSymbol", lIndex)
                    astrToAdd(5) = Str(DateFromString(TableItem(tblDataOut, "DateFilled", lIndex), TableItem(tblDataOut, "TimeFilled", lIndex)))
                    astrToAdd(6) = Str(CLng(UCase(TableItem(tblDataOut, "BuyOrSell", lIndex)) = "B"))
                    astrToAdd(7) = TableItem(tblDataOut, "Lots", lIndex)
                    astrToAdd(8) = TableItem(tblDataOut, "Price", lIndex)
                    
                    HandleMessage eGDSimTradeMessageType_Fill, astrToAdd.JoinFields(vbTab)
                Next lIndex
                
        End Select
    End If
    
ErrExit:
    Set Order = Nothing
    Set Fill = Nothing
    Exit Function
    
ErrSection:
    Set Order = Nothing
    Set Fill = Nothing
    RaiseError "cSimTradeTs.HandleTradeServerMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpTables
'' Description: Dump debug information
'' Inputs:      Title, Return Code, Arrays and Tables to dump
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpTables(ByVal strTitle As String, ByVal lReturn As Long, Optional astrDesc As cGdArray = Nothing, Optional astrResult As cGdArray = Nothing, Optional tblDataOut As cGdTable = Nothing, Optional astrParms As cGdArray = Nothing, Optional tblDataIn As cGdTable = Nothing)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    ' Dump the title...
    DumpTradeServer strTitle
    DumpTradeServer "---------------"
    
    ' Dump the parms array if supplied...
    If Not astrParms Is Nothing Then
        If astrParms.Size > 0 Then
            DumpTradeServer vbTab & "Parms: "
            DumpTradeServer vbTab & vbTab & astrParms.JoinFields(vbTab)
        End If
    End If
    
    ' Dump the description array if supplied...
    If Not astrDesc Is Nothing Then
        If astrDesc.Size > 0 Then
            DumpTradeServer vbTab & "Desc: "
            DumpTradeServer vbTab & vbTab & astrDesc.JoinFields(vbTab)
        End If
    End If
    
    ' Dump the input table if supplied and no error...
    If Not tblDataIn Is Nothing And lReturn = 0 Then
        DumpTradeServer vbTab & "Data In: "
        For lIndex = 0 To tblDataIn.NumRecords - 1
            DumpTradeServer vbTab & vbTab & tblDataIn.GetRecord(lIndex, vbTab)
        Next lIndex
    End If
    
    ' Dump the output table if supplied and no error...
    If Not tblDataOut Is Nothing And lReturn = 0 Then
        DumpTradeServer vbTab & "Data Out: "
        For lIndex = 0 To tblDataOut.NumRecords - 1
            DumpTradeServer vbTab & vbTab & tblDataOut.GetRecord(lIndex, vbTab)
        Next lIndex
    End If
    
    ' Dump the result array if supplied and there was an error...
    If Not astrResult Is Nothing And lReturn <> 0 Then
        DumpTradeServer vbTab & "Result: "
        For lIndex = 0 To astrResult.Size - 1
            DumpTradeServer vbTab & vbTab & astrResult(lIndex)
        Next lIndex
    End If

    DumpTradeServer "---------------"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.DumpTables"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleError
'' Description: Handle an error when it happens
'' Inputs:      Return code, Result array with error information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleError(ByVal lReturnCode As Long, ByVal astrResult As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bStart As Boolean               ' Whether to start dumping or not
    
    If lReturnCode <> 0 Then
        bStart = False
        For lIndex = 0 To astrResult.Size - 1
            If Left(Trim(UCase(astrResult(lIndex))), 8) = "CONTEXT:" Then
                bStart = True
            ElseIf Left(Trim(UCase(astrResult(lIndex))), 16) = "EXECUTION TRACE:" Then
                Exit For
            ElseIf bStart = True Then
                g.ActivityLogs.AddActivityLogEntry m.Broker.Broker, Trim(astrResult(lIndex))
                m.strError = Trim(astrResult(lIndex))
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.HandleError"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectSalmon
'' Description: Start the connection to the Trade Server via Salmon client
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ConnectSalmon() As Boolean
On Error GoTo ErrSection:

    Dim astrParms As cGdArray           ' Parameters to send to the salmon client
    Dim astrIpFile As cGdArray          ' IP file
    Dim lReturn As Long                 ' Return from the salmon call
    
    Set astrParms = New cGdArray
    astrParms.Create eGDARRAY_Strings
    Set astrIpFile = New cGdArray
    astrIpFile.Create eGDARRAY_Strings
    
    astrIpFile.FromFile AddSlash(App.Path) & "Provided\TradeServer.IP"
    
    If astrIpFile.Size > 0 Then
        astrParms(0) = astrIpFile(0)
        astrParms(1) = Str(frmOnlineBroker.txtSalmonCallbackTs.hWnd)
        astrParms(2) = m.strDataServiceID
        astrParms(3) = AddSlash(App.Path) & "TradeServer\Out"
        astrParms(4) = AddSlash(App.Path) & "TradeServer\In"
        astrParms(5) = m.strMachineID
        
        If Not DirExist(astrParms(3)) Then
            MakeDir astrParms(3)
        End If
        If Not DirExist(astrParms(4)) Then
            MakeDir astrParms(4)
        End If
        
        lReturn = StartSalmonTS(astrParms.ArrayHandle)
        DumpTradeServer "StartSalmonTS(" & astrParms.JoinFields(",") & ") = " & Str(lReturn)
        If lReturn <> 0 Then
            InfBox "Error initializing the Genesis SimBroker", "!", , "Genesis SimBroker"
        End If
    Else
        lReturn = -1&
        InfBox "IP address could not be found for the Genesis SimBroker", "!", , "Genesis SimBroker"
    End If
    
    ConnectSalmon = (lReturn = 0)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.ConnectSalmon"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisconnectSalmon
'' Description: Stop the connection to the Trade Server via Salmon client
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DisconnectSalmon() As Boolean
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return from the salmon call
    
    lReturn = StopSalmonTS
    DumpTradeServer "StopSalmonTS = " & Str(lReturn)
    
    DisconnectSalmon = (lReturn = 0)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.DisconnectSalmon"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendSalmon
'' Description: Send stuff to the Trade Server via Salmon client
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SendSalmon() As Boolean
On Error GoTo ErrSection:
    
    Dim astrParms As cGdArray           ' Parameters for the salmon call
    Dim strPath As String               ' Path for the output files
    Dim strGzpFile As String            ' GZP file to send to salmon
    Dim lReturn As Long                 ' Return from the salmon call
    Dim astrFiles As cGdArray           ' Array of matching files
    Dim lIndex As Long                  ' Index into a for loop
    
    strPath = OutPath
    strGzpFile = AddSlash(App.Path) & "TradeServer\Orders_" & m.strDataServiceID & "_" & m.strMachineID & "_" & Str(m.lRequestNum) & ".GZP"
    
    Set astrFiles = New cGdArray
    astrFiles.GetMatchingFiles AddSlash(strPath) & "* /i=ORD,ACK", True, True
    If astrFiles.Size > 0 Then
        For lIndex = 0 To astrFiles.Size - 1
            If FileExist(strGzpFile) Then
                ZipExecute "A", strGzpFile, "", astrFiles(lIndex)
            Else
                ZipExecute "C", strGzpFile, "", astrFiles(lIndex)
            End If
            
            If UCase(FileExt(astrFiles(lIndex))) = "ORD" Then
                RenameFile astrFiles(lIndex), Replace(UCase(astrFiles(lIndex)), ".ORD", ".SNT")
            End If
        Next lIndex
        
        Set astrParms = New cGdArray
        astrParms.Create eGDARRAY_Strings
        
        astrParms(0) = strGzpFile
        
        m.lRequestNum = m.lRequestNum + 1&
        lReturn = SendSalmonTS(astrParms.ArrayHandle)
        DumpTradeServer "SendSalmonTS(" & astrParms.JoinFields(",") & ") = " & Str(lReturn)
    End If
    
    SendSalmon = (lReturn = 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeTs.SendSalmon"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckSimTradeStatus
'' Description: Do a download so that we can see if any new information about
''              the simulated trading comes down
'' Inputs:      Message, Status Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckSimTradeStatus(Optional ByVal strMsg As String = "", Optional ByVal strStatusMsg As String = "")
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return from the question to the user

    If UseSalmon Then
        SendSalmon
    Else
        strReturn = "Y"
        If Len(strMsg) > 0 Then
            strReturn = InfBox(strMsg & "|(This requires a connection to Genesis)", "?", "+Yes|-No", "Account Status")
        End If
            
        If strReturn = "Y" Then
            If Len(strStatusMsg) = 0 Then
                frmStatus.SetTitle "Checking SimTrade Order Status"
            Else
                frmStatus.SetTitle strStatusMsg
            End If
            GetNYTime
        End If
        
        If frmStatus.Status = eStatus_Completed Then frmStatus.AddDetail "Finished"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.CheckSimTradeStatus"
    
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
    Set m.SimTradeBase = New cSimTradeBase
    Set m.astrSimTradeFiles = New cGdArray
    m.astrSimTradeFiles.Create eGDARRAY_Strings
    
    Set m_txtCallback = Nothing
    Set m_tmrSalmon = Nothing
    
    m.bTradeServerInitCalled = False
    m.lRequestNum = 0&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.Class_Initialize"
    
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
    Set m.SimTradeBase = Nothing
    Set m.astrSimTradeFiles = Nothing

    Set m_txtCallback = Nothing
    Set m_tmrSalmon = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeTs.Class_Terminate"
    
End Sub
