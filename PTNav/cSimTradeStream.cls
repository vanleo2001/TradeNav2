VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSimTradeStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cSimTradeStream.cls
'' Description: Routines for handling simulated trading through Trade Navigator
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/24/2011   DAJ         Added NextAccount
'' 06/24/2011   DAJ         Don't allow trading indexes
'' 08/02/2011   DAJ         When parking or cancelling all, optionally include auto trade orders
'' 08/11/2011   DAJ         Allow forex trading for the trade server
'' 08/15/2011   DAJ         Added GetPositionForSymbol stuff
'' 09/16/2011   DAJ         Don't fill an order at a price of zero
'' 02/14/2012   DAJ         Added multi-leg order support
'' 03/01/2012   DAJ         Added MIT as an order type
'' 05/15/2012   DAJ         Fix for fill quantity on single leg order
'' 08/16/2012   DAJ         Set the snapshot flag on a SimTrade order when it gets sent
'' 08/28/2012   DAJ         CarriedFillFromInfo, First Field Begin/End
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 03/19/2013   DAJ         When cancelling all SimTrade orders, don't always confirm each cancel
'' 08/28/2014   DAJ         Added profiling for CheckQuickFillOrders
'' 08/29/2014   DAJ         Moved session date calculation from CheckQuickFillOrders into the order object
'' 09/04/2014   DAJ         Pulled SimTrade conversions out of trade objects
'' 10/23/2014   DAJ         Renamed mDataNav.SpreadComponents to mDataNav.SpreadComponentsForBars
'' 11/14/2014   DAJ         Added support for MIT, On-Close, and On-Open orders
'' 09/14/2015   DAJ         Get new copy of order before parking in ParkAllWorkingOrders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Broker As cBroker                   ' Broker base class object
    SimTradeBase As cSimTradeBase       ' Simulated trading base class object

    astrLoggedNotValid As cGdArray      ' Array of symbol for which not valid time message was logged
    strStopHit As String                ' List of demo orders for which the stop has been hit
    strMitHit As String                 ' List of demo orders for which the MIT has been hit
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

    Select Case nBroker
        Case eTT_AccountType_SimStream
            m.Broker.Init Me, eTT_AccountType_SimStream, AddSlash(App.Path) & "SimTrade"
            m.Broker.IniFile = AddSlash(App.Path) & "SimTrade.INI"
            m.Broker.BrokerName = "Trade Navigator"
            m.Broker.ProcessName = ""
            m.Broker.ProcessPath = ""
            m.Broker.ControlID = ""
            m.Broker.ConnectIni = ""
    
        Case eTT_AccountType_SimReplay
            m.Broker.Init Me, eTT_AccountType_SimReplay, AddSlash(App.Path) & "StreamReplay"
            m.Broker.IniFile = AddSlash(App.Path) & "StreamReplay.INI"
            m.Broker.BrokerName = "Stream Replay"
            m.Broker.ProcessName = ""
            m.Broker.ProcessPath = ""
            m.Broker.ControlID = ""
            m.Broker.ConnectIni = ""
    
    End Select

    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Connect) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Disconnect) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccounts) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccountsMode) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ConnectInfo) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ChangePassword) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Refresh) = (g.nReplaySession = 0)
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewActivity) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_BrokerView) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewOnline) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_VerifyPositions) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_AccountDetails) = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeStream.Init"

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
    RaiseError "cSimTradeStream.ShowConfig"
    
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

    ShowBrokerConnectionInfo = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.ShowBrokerConnectionInfo"
    
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

    ShowLogin = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.ShowLogin"
    
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

    If m.Broker.Broker = eTT_AccountType_SimStream Then
        If g.RealTime.ConnectionStatus = eGDConnectionStatus_Disconnected Then
            If InfBox("This requires a connection to streaming.|Would you like to connect now?", "?", "+Yes|-No", "Connect to Streaming") = "Y" Then
                If Not ProcessIsBusy Then
                    g.nReplaySession = 0
                    g.RealTime.Init True
                Else
                    m.Broker.HandleConnectionInfo eGDConnectionStatus_Disconnected, "", ""
                End If
            Else
                m.Broker.HandleConnectionInfo eGDConnectionStatus_Disconnected, "", ""
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeStream.SendConnect"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeStream.SendDisconnect"
    
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
    RaiseError "cSimTradeStream.UnloadApp"
    
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

    strReturn = ""
    If (IsForex(strGenesisSymbol) = True) Or (SecurityType(strGenesisSymbol, True) <> "I") Then
        strReturn = strGenesisSymbol
    End If
    
    BrokerSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.BrokerSymbol"
    
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
    RaiseError "cSimTradeStream.GenesisSymbol"
    
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
        Case eTT_OrderType_MIT
            strReturn = "MIT"
    End Select
    
    BrokerOrderType = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.BrokerOrderType"
    
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

    m.Broker.DumpDebug "AddOrder: " & Order.OrderText(True, True, True)
    If Order.Status <> eTT_OrderStatus_Sent Then
        ' DAJ 08/16/2012: Tim is running into issues with automated trading where an
        ' order is being sent just before a refresh.  The refresh isn't picking up the
        ' sent order because the snapshot flag isn't set yet...
        Order.IsSnapshot = True
        Order.ChangeOrderStatus eTT_OrderStatus_Sent, m.Broker.BrokerInfo
    End If
    
    AddOrder = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.AddOrder"
    
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

    m.Broker.DumpDebug "AmendOrder: " & Order.OrderText(True, True, True)
    If Order.Status <> eTT_OrderStatus_AmendPending Then
        Order.ChangeOrderStatus eTT_OrderStatus_AmendPending, m.Broker.BrokerInfo
    End If
    
    AmendOrder = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.AmendOrder"
    
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

    m.Broker.DumpDebug "CancelOrder: " & Order.OrderText(True, True, True)
    If (g.bStarting = True) Or (g.bUnloading = True) Or (g.RealTime.Active = False) Then
        Order.ChangeOrderStatus eTT_OrderStatus_Cancelled, m.Broker.BrokerInfo
    ElseIf Order.Status <> eTT_OrderStatus_CancelPending Then
        Order.ChangeOrderStatus eTT_OrderStatus_CancelPending, m.Broker.BrokerInfo
    End If
    
    CancelOrder = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.CancelOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelAllWorkingOrders
'' Description: Cancel all working orders
'' Inputs:      Include automated trading orders?, Confirm Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelAllWorkingOrders(Optional ByVal bIncludeAutoTradeOrders As Boolean = True, Optional ByVal bConfirm As Boolean = True)
On Error GoTo ErrSection:

    Dim Orders As cPtOrders             ' Collection of working orders
    Dim lIndex As Long                  ' Index into a for loop
    
    m.Broker.DumpDebug "Cancelling All Working Orders"
    Set Orders = m.Broker.BrokerInfo.Orders.MakeCopy
    If Not Orders Is Nothing Then
        For lIndex = Orders.Count To 1 Step -1
            If (bIncludeAutoTradeOrders = True) Or (Orders(lIndex).AutoTradeItemID = 0) Then
                If IsOpenOrder(Orders(lIndex).Status) Then
                    CancelOrderFromOrder Orders(lIndex), "SimTradeStream", False, bConfirm
                End If
            End If
        Next lIndex
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeStream.CancelAllWorkingOrders"
    
End Sub

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

    m.Broker.DumpDebug "CancelOrder: " & Order.OrderText(True, True, True)
    If g.bUnloading Then
        Order.ChangeOrderStatus eTT_OrderStatus_Parked, m.Broker.BrokerInfo
    ElseIf Order.Status <> eTT_OrderStatus_ParkPending Then
        Order.ChangeOrderStatus eTT_OrderStatus_ParkPending, m.Broker.BrokerInfo
    End If
    
    ParkOrder = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.ParkOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkAllWorkingOrders
'' Description: Park all working orders
'' Inputs:      Include automated trading orders?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ParkAllWorkingOrders(Optional ByVal bIncludeAutoTradeOrders As Boolean = True)
On Error GoTo ErrSection:

    Dim Orders As cPtOrders             ' Collection of working orders
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cPtOrder               ' Order from the broker info object
    
    m.Broker.DumpDebug "Parking All Working Orders"
    Set Orders = m.Broker.BrokerInfo.Orders.MakeCopy
    If Not Orders Is Nothing Then
        For lIndex = Orders.Count To 1 Step -1
            If (bIncludeAutoTradeOrders = True) Or (Orders(lIndex).AutoTradeItemID = 0) Then
                ' DAJ 09/14/2015: Try to get the order out of the BrokerInfo object again here in case
                ' it changed since we got the collection above ( e.g. a Triggered-By order already being
                ' parked because the Triggering order was parked )...
                Set Order = m.Broker.BrokerInfo.OrderForAccessID(Orders(lIndex).AccountID, Orders(lIndex).OrderID)
                If Not Order Is Nothing Then
                    If IsOpenOrder(Order.Status) Then
                        ParkOrderFromOrder Order, "SimTradeStream"
                    End If
                End If
            End If
        Next lIndex
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeStream.ParkAllWorkingOrders"
    
End Sub

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
    RaiseError "cSimTradeStream.FlattenExpiredPositions"
    
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
    RaiseError "cSimTradeStream.FlattenExpiredPosition"
    
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
    RaiseError "cSimTradeStream.Refresh"
    
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
    RaiseError "cSimTradeStream.GetAccounts"
    
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

    Dim astrOrders As cGdArray          ' Array of orders from the database
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrOrders = m.SimTradeBase.GetSnapshotOrdersFromDb(m.Broker.Broker, strAccount)
    
    HandleMessage eGDSimTradeMessageType_RefreshOrder, "BEGIN"
    For lIndex = 0 To astrOrders.Size - 1
        HandleMessage eGDSimTradeMessageType_RefreshOrder, astrOrders(lIndex)
    Next lIndex
    HandleMessage eGDSimTradeMessageType_RefreshOrder, "END"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeStream.GetOrders"
    
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
    RaiseError "cSimTradeStream.GetFills"
    
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
    RaiseError "cSimTradeStream.GetPositions"
    
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
    RaiseError "cSimTradeStream.GetPositionForSymbol"
    
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
    RaiseError "cSimTradeStream.NextGenesisID"
    
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
    RaiseError "cSimTradeStream.StartingGenesisIdForOptNav"
    
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
    RaiseError "cSimTradeStream.IsTradeableSymbol"
    
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
    RaiseError "cSimTradeStream.IsEnabledSymbol"
    
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
            bReturn = True
        Case eTT_OrderType_MarketOnOpen
            bReturn = False
        Case eTT_OrderType_Adjustment
            bReturn = False
    End Select

    OrderTypeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.OrderTypeAllowed"
    
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
    RaiseError "cSimTradeStream.TimeInForceAllowed"
    
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
    RaiseError "cSimTradeStream.DateToDouble"
    
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
    RaiseError "cSimTradeStream.HandleMessage"
    
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
    RaiseError "cSimTradeStream.AccountFromInfo"
    
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

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.OrderFromInfo"
    
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
    RaiseError "cSimTradeStream.FillFromInfo"
    
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
    RaiseError "cSimTradeStream.CarriedFillFromInfo"
    
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
    RaiseError "cSimTradeStream.PositionFromInfo"
    
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
    RaiseError "cSimTradeStream.GenesisSymbolList"
    
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
    RaiseError "cSimTradeStream.TimeZone"
    
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
    RaiseError "cSimTradeStream.JustConnected"
    
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
    RaiseError "cSimTradeStream.SymbolInformation"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckQuickFillOrders
'' Description: Check quick fill orders if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckQuickFillOrders()
On Error GoTo ErrSection:

gdStartProfile 520
gdStartProfile 521

    Dim lAccountID As Long              ' Account ID for the demo account
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Order As cPtOrder               ' Temporary order object
    Dim Fill As cPtFill                 ' Temporary fill object
    Dim Bars As New cGdBars             ' Temporary bars object
    Dim bFill As Boolean                ' Fill the order?
    Dim dLast As Double                 ' Last price from the bars
    Dim dBid As Double                  ' Bid price from the bars
    Dim dAsk As Double                  ' Ask price from the bars
    Dim dFillPrice As Double            ' Fill price to take
    Dim dExchangeTime As Double         ' Current date/time in exchange time
    Dim dFeedTime As Double             ' Current feed time
    'Dim nAcctType As eTT_AccountType    ' Account type for the order
    Dim bAllowTrading As Boolean        ' Do we want to allow trading?
    Dim Orders As cPtOrders             ' Collection of working orders
    Dim lSymbolID As Long               ' Symbol ID for the order
    Dim vSymbolOrSymbolID As Variant    ' Symbol or Symbol ID for the order
    Dim lPos As Long                    ' Position in an array
    Dim bDump As Boolean                ' Do extra logging?
    Dim OrderLeg As cOrderLeg           ' Order leg object
    Dim bBuy As Boolean                 ' Is the order a buy?
    Dim Bars2 As cGdBars                ' Bars for the lead contract of a calendar spread
    Dim astrComponents As cGdArray      ' Components for a calendar spread
    Dim dFillPrice2 As Double           ' Fill price for the lead contract of a calendar spread
    
    If g.nReplayAccountID = 0 Then
        bAllowTrading = (g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected)
    ElseIf FormIsLoaded("frmReplay") Then
        bAllowTrading = (g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected) And (frmReplay.IsPlaying = True)
    Else
        bAllowTrading = False
    End If
    
gdStopProfile 521
    If bAllowTrading Then
gdStartProfile 522
        Set Orders = m.Broker.BrokerInfo.Orders
gdStopProfile 522
        If Not Orders Is Nothing Then
gdStartProfile 523
            bDump = m.Broker.BrokerInfo.HasPendingOrders
gdStopProfile 523
            
            For lIndex = Orders.Count To 1 Step -1
                ' DAJ 08/06/2007: Only need to check open (including pending) orders...
                If IsOpenOrder(Orders(lIndex).Status) Then
gdStartProfile 524
                    If bDump Then
                        m.Broker.DumpDebug "Pending Orders Exist: '" & Orders(lIndex).OrderText(True, True, True) & "' - " & OrderStatus(Orders(lIndex).Status)
                    End If
                    
                    'SetBarProperties Bars, Orders(lIndex).Symbol
                    dFeedTime = g.RealTime.FeedTime(Orders(lIndex).Symbol)
                    'nAcctType = Orders(lIndex).Broker ' g.Broker.AccountTypeForNumber(Orders(lIndex).AccountID)
gdStopProfile 524

gdStartProfile 525
                    Select Case Orders(lIndex).Status
                        Case eTT_OrderStatus_Sent, eTT_OrderStatus_AmendPending
                            If (Orders(lIndex).NumberOfLegs = 1) And (SecurityType(Orders(lIndex).Symbol) = "FO") And (Orders(lIndex).OrderType <> eTT_OrderType_Limit) Then
                                Orders(lIndex).SetSessionDateForFeedTime dFeedTime
                                Orders(lIndex).Status = eTT_OrderStatus_Rejected
                                Orders(lIndex).StatusDate = dFeedTime
                                If Orders(lIndex).OrderDate = 0# Then Orders(lIndex).OrderDate = dFeedTime
                                Orders(lIndex).Message = "Only Limit orders are allowed for future options"
                                Orders(lIndex).BrokerID = "DBO" & Orders(lIndex).GenesisOrderID
                            Else
                                Orders(lIndex).SetSessionDateForFeedTime dFeedTime
                                Orders(lIndex).Status = eTT_OrderStatus_Working
                                Orders(lIndex).StatusDate = dFeedTime
                                If Orders(lIndex).OrderDate = 0# Then Orders(lIndex).OrderDate = dFeedTime
                                Orders(lIndex).BrokerID = "DBO" & Orders(lIndex).GenesisOrderID
                                Orders(lIndex).ExchangeID = "EBO" & Orders(lIndex).GenesisOrderID
                                Orders(lIndex).DateSent = Orders(lIndex).OrderDate
                                Orders(lIndex).DateHostRecd = dFeedTime
                                Orders(lIndex).DateExchRecd = dFeedTime
                                Orders(lIndex).DateExchAckn = dFeedTime
                            End If

                            HandleMessage eGDSimTradeMessageType_Order, m.SimTradeBase.OrderToSimTradeString(Orders(lIndex))
                        
                        Case eTT_OrderStatus_CancelPending
                            Orders(lIndex).Status = eTT_OrderStatus_Cancelled
                            Orders(lIndex).StatusDate = dFeedTime
                            
                            HandleMessage eGDSimTradeMessageType_Order, m.SimTradeBase.OrderToSimTradeString(Orders(lIndex))
                        
                        Case eTT_OrderStatus_ParkPending
                            Orders(lIndex).Status = eTT_OrderStatus_Parked
                            Orders(lIndex).StatusDate = dFeedTime
                            
                            HandleMessage eGDSimTradeMessageType_Order, m.SimTradeBase.OrderToSimTradeString(Orders(lIndex))
                        
                        Case eTT_OrderStatus_Partial, eTT_OrderStatus_Working
                            If TradeBarsExist(Orders(lIndex).SymbolOrSymbolID) Then
gdStartProfile 526
                                Set Bars = New cGdBars
                                If Orders(lIndex).NumberOfLegs > 1 Then
                                    Set Bars = GetTradeBars(Orders(lIndex).SpreadSymbol).MakeCopy
                                    bBuy = (Orders(lIndex).DebitCredit = eGDDebitCredit_Debit)
                                Else
                                    Set Bars = GetTradeBars(Orders(lIndex).SymbolOrSymbolID).MakeCopy
                                    bBuy = Orders(lIndex).Buy
                                End If
                                dExchangeTime = ConvertTimeZone(dFeedTime, TimeZone(Orders(lIndex).Symbol), Bars.Prop(eBARS_ExchangeTimeZoneInf))
gdStopProfile 526
                                If Bars.IsValidTradingTime(dExchangeTime) = True Then
gdStartProfile 527
                                    ' Remove the symbol from the "logged not valid" collection since
                                    ' the time is valid once again (this way we will log it again when
                                    ' and if the trading time goes invalid again)...
                                    If m.astrLoggedNotValid.BinarySearch(Orders(lIndex).Symbol, lPos) = True Then
                                        m.astrLoggedNotValid.Remove lPos
                                    End If
                                    
                                    dLast = Bars(eBARS_Close, Bars.Size - 1)
                                    dBid = Bars(eBARS_Bid, Bars.Size - 1)
                                    'If (dBid = kNullData) Or (dBid = 0#) Then dBid = dLast
                                    dAsk = Bars(eBARS_Ask, Bars.Size - 1)
                                    'If (dAsk = kNullData) Or (dAsk = 0#) Then dAsk = dLast
                                    If (dLast <> kNullData) Or (dBid <> kNullData) Or (dAsk <> kNullData) Then
                                        bFill = False
gdStartProfile 528
                                        Select Case Orders(lIndex).OrderType
                                            Case eTT_OrderType_Market
                                                dFillPrice = MarketPrice(bBuy, Bars)
                                                If dFillPrice <> kNullData Then
                                                    bFill = True
                                                    m.Broker.DumpDebug "Filling Market Order " & Orders(lIndex).BrokerID & ": Bid=" & Str(dBid) & ", Ask=" & Str(dAsk) & ", Last=" & Str(dLast) & ", Fill Price=" & Str(dFillPrice)
                                                End If
                                                
                                            Case eTT_OrderType_Stop
                                                If InStr(m.strStopHit, "," & Orders(lIndex).BrokerID & ",") = 0 Then
                                                    If StopPriceHit(Orders(lIndex).StopPrice, bBuy, Bars, Orders(lIndex).OrderDate) Then
                                                        m.strStopHit = m.strStopHit & "," & Orders(lIndex).BrokerID & ","
                                                        m.Broker.DumpDebug "Stop Price Hit for " & Orders(lIndex).BrokerID & ": Bid=" & Str(dBid) & ", Ask=" & Str(dAsk) & ", Last=" & Str(dLast)
                                                    End If
                                                Else
                                                    dFillPrice = MarketPrice(bBuy, Bars)
                                                    If dFillPrice <> kNullData Then
                                                        bFill = True
                                                        m.Broker.DumpDebug "Filling Stop Order " & Orders(lIndex).BrokerID & ": Bid=" & Str(dBid) & ", Ask=" & Str(dAsk) & ", Last=" & Str(dLast) & ", Fill Price=" & Str(dFillPrice)
                                                    End If
                                                End If
                                            
                                            Case eTT_OrderType_Limit
                                                If InStr(Bars.Prop(eBARS_Symbol), " ") = 0 Then
                                                    dFillPrice = LimitOrderFilled(Orders(lIndex).LimitPrice, bBuy, Bars)
                                                Else
                                                    dFillPrice = LimitOrderFilledOptions(Orders(lIndex).LimitPrice, bBuy, Bars)
                                                End If
                                                If dFillPrice <> kNullData Then
                                                    bFill = True
                                                    m.Broker.DumpDebug "Filling Limit Order " & Orders(lIndex).BrokerID & ": Bid=" & Str(dBid) & ", Ask=" & Str(dAsk) & ", Last=" & Str(dLast) & ", Fill Price=" & Str(dFillPrice)
                                                End If
                                            
                                            Case eTT_OrderType_StopWithLimit
                                                If InStr(m.strStopHit, "," & Orders(lIndex).BrokerID & ",") = 0 Then
                                                    If StopPriceHit(Orders(lIndex).StopPrice, bBuy, Bars, Orders(lIndex).OrderDate) Then
                                                        m.strStopHit = m.strStopHit & "," & Orders(lIndex).BrokerID & ","
                                                        m.Broker.DumpDebug "Stop Price Hit for " & Orders(lIndex).BrokerID & ": Bid=" & Str(dBid) & ", Ask=" & Str(dAsk) & ", Last=" & Str(dLast)
                                                    End If
                                                Else
                                                    If InStr(Bars.Prop(eBARS_Symbol), " ") = 0 Then
                                                        dFillPrice = LimitOrderFilled(Orders(lIndex).LimitPrice, bBuy, Bars)
                                                    Else
                                                        dFillPrice = LimitOrderFilledOptions(Orders(lIndex).LimitPrice, bBuy, Bars)
                                                    End If
                                                    If dFillPrice <> kNullData Then
                                                        bFill = True
                                                        m.strStopHit = Replace(m.strStopHit, "," & Orders(lIndex).BrokerID & ",", "")
                                                        m.Broker.DumpDebug "Filling Stop with Limit Order " & Orders(lIndex).BrokerID & ": Bid=" & Str(dBid) & ", Ask=" & Str(dAsk) & ", Last=" & Str(dLast) & ", Fill Price=" & Str(dFillPrice)
                                                    End If
                                                End If
                                                
                                            Case eTT_OrderType_MIT
                                                If InStr(m.strMitHit, "," & Orders(lIndex).BrokerID & ",") = 0 Then
                                                    If MITPriceHit(Orders(lIndex).MitPrice, bBuy, Bars) Then
                                                        m.strMitHit = m.strMitHit & "," & Orders(lIndex).BrokerID & ","
                                                        m.Broker.DumpDebug "MIT Price Hit for " & Orders(lIndex).BrokerID & ": Bid=" & Str(dBid) & ", Ask=" & Str(dAsk) & ", Last=" & Str(dLast)
                                                    End If
                                                Else
                                                    dFillPrice = MarketPrice(bBuy, Bars)
                                                    If dFillPrice <> kNullData Then
                                                        bFill = True
                                                        m.Broker.DumpDebug "Filling MIT Order " & Orders(lIndex).BrokerID & ": Bid=" & Str(dBid) & ", Ask=" & Str(dAsk) & ", Last=" & Str(dLast) & ", Fill Price=" & Str(dFillPrice)
                                                    End If
                                                End If
                                                                                                
                                        End Select
gdStopProfile 528
                                        
                                        If bFill = True Then
                                            If g.RealTime.RtDataAvailable(Orders(lIndex).SymbolOrSymbolID, ePRD_Days + 1) = True Then
                                                Orders(lIndex).Status = eTT_OrderStatus_Filled
                                                Orders(lIndex).StatusDate = dFeedTime
                                                
                                                HandleMessage eGDSimTradeMessageType_Order, m.SimTradeBase.OrderToSimTradeString(Orders(lIndex))
                                                
                                                If IsSpreadSymbol(Orders(lIndex).Symbol) Then
                                                    Set astrComponents = SpreadComponentsForBars(Bars)
                                                    If astrComponents.Size > 0 Then
                                                        Set Bars2 = GetTradeBars(astrComponents(0))
                                                        If Not Bars2 Is Nothing Then
                                                            Set OrderLeg = Orders(lIndex).OrderLegs(1)
                                                    
                                                            dFillPrice2 = MarketPrice(bBuy, Bars2)
                                                            
                                                            Set Fill = New cPtFill
                                                            Fill.AccountID = Orders(lIndex).AccountID
                                                            Fill.AutoTradingItemID = Orders(lIndex).AutoTradeItemID
                                                            Fill.BrokerID = "DFO" & Orders(lIndex).GenesisOrderID
                                                            Fill.BrokerOrderID = Orders(lIndex).BrokerID
                                                            Fill.Buy = Orders(lIndex).Buy
                                                            Fill.FillDate = dFeedTime
                                                            Fill.IsManual = False
                                                            Fill.IsSnapshot = True
                                                            Fill.OrderID = Orders(lIndex).OrderID
                                                            Fill.Price = dFillPrice
                                                            Fill.Quantity = Orders(lIndex).Quantity * OrderLeg.Multiplier
                                                            Fill.SessionDate = Bars.SessionDateForTradeTime(ConvertTimeZone(dFeedTime, "NY", Bars.Prop(eBARS_ExchangeTimeZoneInf)))
                                                            Fill.SymbolOrSymbolID = Orders(lIndex).SymbolOrSymbolID
                                                        
                                                            HandleMessage eGDSimTradeMessageType_SpreadFill, m.SimTradeBase.FillToSimTradeString(Fill)
                                                            
                                                            Set Fill = New cPtFill
                                                            Fill.AccountID = Orders(lIndex).AccountID
                                                            Fill.AutoTradingItemID = Orders(lIndex).AutoTradeItemID
                                                            Fill.BrokerID = "DFO" & Orders(lIndex).GenesisOrderID & "_1"
                                                            Fill.BrokerOrderID = Orders(lIndex).BrokerID
                                                            Fill.Buy = Orders(lIndex).Buy
                                                            Fill.FillDate = dFeedTime
                                                            Fill.IsManual = False
                                                            Fill.IsSnapshot = True
                                                            Fill.OrderID = Orders(lIndex).OrderID
                                                            Fill.Price = dFillPrice2
                                                            Fill.Quantity = Orders(lIndex).Quantity
                                                            Fill.SessionDate = Bars2.SessionDateForTradeTime(ConvertTimeZone(dFeedTime, "NY", Bars2.Prop(eBARS_ExchangeTimeZoneInf)))
                                                            Fill.SymbolOrSymbolID = Bars2.SymbolOrSymbolID
                                                    
                                                            HandleMessage eGDSimTradeMessageType_Fill, m.SimTradeBase.FillToSimTradeString(Fill)
                                                        
                                                            Set Fill = New cPtFill
                                                            Fill.AccountID = Orders(lIndex).AccountID
                                                            Fill.AutoTradingItemID = Orders(lIndex).AutoTradeItemID
                                                            Fill.BrokerID = "DFO" & Orders(lIndex).GenesisOrderID & "_2"
                                                            Fill.BrokerOrderID = Orders(lIndex).BrokerID
                                                            Fill.Buy = Not Orders(lIndex).Buy
                                                            Fill.FillDate = dFeedTime
                                                            Fill.IsManual = False
                                                            Fill.IsSnapshot = True
                                                            Fill.OrderID = Orders(lIndex).OrderID
                                                            Fill.Price = dFillPrice2 - dFillPrice
                                                            Fill.Quantity = Orders(lIndex).Quantity
                                                            Fill.SessionDate = Bars2.SessionDateForTradeTime(ConvertTimeZone(dFeedTime, "NY", Bars2.Prop(eBARS_ExchangeTimeZoneInf)))
                                                            Fill.SymbolOrSymbolID = astrComponents(1)
                                                    
                                                            HandleMessage eGDSimTradeMessageType_Fill, m.SimTradeBase.FillToSimTradeString(Fill)
                                                        End If
                                                    End If
                                                    
                                                ElseIf Orders(lIndex).NumberOfLegs = 1 Then
                                                    Set OrderLeg = Orders(lIndex).OrderLegs(1)
                                                    
                                                    Set Fill = New cPtFill
                                                    Fill.AccountID = Orders(lIndex).AccountID
                                                    Fill.AutoTradingItemID = Orders(lIndex).AutoTradeItemID
                                                    Fill.BrokerID = "DFO" & Orders(lIndex).GenesisOrderID
                                                    Fill.BrokerOrderID = Orders(lIndex).BrokerID
                                                    Fill.Buy = Orders(lIndex).Buy
                                                    Fill.FillDate = dFeedTime
                                                    Fill.IsManual = False
                                                    Fill.IsSnapshot = True
                                                    Fill.OrderID = Orders(lIndex).OrderID
                                                    Fill.Price = dFillPrice
                                                    Fill.Quantity = Orders(lIndex).Quantity * OrderLeg.Multiplier
                                                    Fill.SessionDate = Bars.SessionDateForTradeTime(ConvertTimeZone(dFeedTime, "NY", Bars.Prop(eBARS_ExchangeTimeZoneInf)))
                                                    Fill.SymbolOrSymbolID = Orders(lIndex).SymbolOrSymbolID
                                                
                                                    HandleMessage eGDSimTradeMessageType_Fill, m.SimTradeBase.FillToSimTradeString(Fill)
                                                    
                                                Else
                                                    For lIndex2 = 1 To Orders(lIndex).NumberOfLegs
                                                        Set OrderLeg = Orders(lIndex).OrderLegs(lIndex2)
                                                        If Not OrderLeg Is Nothing Then
                                                            Set Bars = GetTradeBars(OrderLeg.SymbolOrSymbolID)
                                                            If Not Bars Is Nothing Then
                                                                If OrderLeg.IsBuy Then
                                                                    dFillPrice = Bars(eBARS_Ask, Bars.Size - 1)
                                                                Else
                                                                    dFillPrice = Bars(eBARS_Bid, Bars.Size - 1)
                                                                End If
                                                            
                                                                Set Fill = New cPtFill
                                                                Fill.AccountID = Orders(lIndex).AccountID
                                                                Fill.AutoTradingItemID = Orders(lIndex).AutoTradeItemID
                                                                Fill.BrokerID = "DFO" & Orders(lIndex).GenesisOrderID & "_" & Str(lIndex2)
                                                                Fill.BrokerOrderID = Orders(lIndex).BrokerID
                                                                Fill.Buy = OrderLeg.IsBuy
                                                                Fill.FillDate = dFeedTime
                                                                Fill.IsManual = False
                                                                Fill.IsSnapshot = True
                                                                Fill.OrderID = Orders(lIndex).OrderID
                                                                Fill.Price = dFillPrice
                                                                Fill.Quantity = Orders(lIndex).Quantity * OrderLeg.Multiplier
                                                                Fill.SessionDate = Bars.SessionDateForTradeTime(ConvertTimeZone(dFeedTime, "NY", Bars.Prop(eBARS_ExchangeTimeZoneInf)))
                                                                Fill.SymbolOrSymbolID = OrderLeg.SymbolOrSymbolID
                                                            
                                                                HandleMessage eGDSimTradeMessageType_Fill, m.SimTradeBase.FillToSimTradeString(Fill)
                                                            End If
                                                        End If
                                                    Next lIndex2
                                                End If
                                            End If
                                        End If
                                    End If
                                ElseIf Orders(lIndex).SessionDate <> 0 Then
                                    ' Only log the invalid trading time message once...
                                    If m.astrLoggedNotValid.BinarySearch(Orders(lIndex).Symbol, lPos) = False Then
                                        m.Broker.DumpDebug "Not valid trading time for " & Orders(lIndex).Symbol & ": " & DateFormat(dExchangeTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
                                        m.Broker.DumpDebug "Expiration: " & Str(Orders(lIndex).Expiration) & " ; Exchange Time: " & DateFormat(dExchangeTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & " ; End Session: " & DateFormat((Abs(Orders(lIndex).SessionDate) + (Bars.Prop(eBARS_DefaultEndTime) / 1440#)), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
                                        
                                        m.astrLoggedNotValid.Add Orders(lIndex).Symbol, lPos
                                    End If
                                    
                                    ' If this is a "day" order, then we need to expire the order if it is still
                                    ' working at the end of the session date of the order...
                                    If Orders(lIndex).Expiration < 0 Then
                                        If dExchangeTime > (Abs(Orders(lIndex).SessionDate) + (Bars.Prop(eBARS_DefaultEndTime) / 1440#)) Then
                                            Orders(lIndex).Status = eTT_OrderStatus_Expired
                                            Orders(lIndex).StatusDate = dFeedTime
                                            
                                            HandleMessage eGDSimTradeMessageType_Order, m.SimTradeBase.OrderToSimTradeString(Orders(lIndex))
                                        End If
                                        
                                    ' Otherwise, if this is a GTD order, we need to expire the order if it is
                                    ' still working at the end of the expiration session...
                                    ElseIf Orders(lIndex).Expiration > 0 Then
                                        If dExchangeTime > (Abs(Orders(lIndex).Expiration) + (Bars.Prop(eBARS_DefaultEndTime) / 1440#)) Then
                                            Orders(lIndex).Status = eTT_OrderStatus_Expired
                                            Orders(lIndex).StatusDate = dFeedTime
                                            
                                            HandleMessage eGDSimTradeMessageType_Order, m.SimTradeBase.OrderToSimTradeString(Orders(lIndex))
                                        End If
                                    End If
gdStopProfile 527
                                End If
                            End If
                    End Select
gdStopProfile 525
                End If
            Next lIndex
        End If
    End If
    
gdStopProfile 520
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeStream.CheckQuickFillOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextAccount
'' Description: Determine the next simulated stream account number
'' Inputs:      None
'' Returns:     Next Simulated Stream account number
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
        strAccount = "SIM" & Format(lIndex, "0000")
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
    RaiseError "cSimTradeStream.NextAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MarketPrice
'' Description: Determine the market price for the given Bars based on the side
'' Inputs:      Buy?, Bars
'' Returns:     Market Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MarketPrice(ByVal bIsBuy As Boolean, Bars As cGdBars) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function

    dReturn = kNullData
    If IsPitFuture(Bars.Prop(eBARS_Symbol)) Then
        If Bars(eBARS_Close, Bars.Size - 1) <> kNullData Then
            dReturn = Bars(eBARS_Close, Bars.Size - 1)
        ElseIf bIsBuy = True Then
            If Bars(eBARS_Ask, Bars.Size - 1) <> kNullData Then
                dReturn = Bars(eBARS_Ask, Bars.Size - 1)
            End If
        Else
            If Bars(eBARS_Bid, Bars.Size - 1) <> kNullData Then
                dReturn = Bars(eBARS_Bid, Bars.Size - 1)
            End If
        End If
    Else
        If bIsBuy = True Then
            If Bars(eBARS_Ask, Bars.Size - 1) <> kNullData Then
                dReturn = Bars(eBARS_Ask, Bars.Size - 1)
            ElseIf Bars(eBARS_Close, Bars.Size - 1) <> kNullData Then
                dReturn = Bars(eBARS_Close, Bars.Size - 1)
            End If
        Else
            If Bars(eBARS_Bid, Bars.Size - 1) <> kNullData Then
                dReturn = Bars(eBARS_Bid, Bars.Size - 1)
            ElseIf Bars(eBARS_Close, Bars.Size - 1) <> kNullData Then
                dReturn = Bars(eBARS_Close, Bars.Size - 1)
            End If
        End If
    End If
    
    MarketPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.MarketPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BidSidePrice
'' Description: Determine the prices on the bid side of the market
'' Inputs:      Bars, Bid Price, Bid Time, Last Tick Price, Last Tick Time
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BidSidePrice(ByVal Bars As cGdBars, dBidPrice As Double, dBidTime As Double, dTickPrice As Double, dTickTime As Double)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol for the Bars
    
    strSymbol = Bars.Prop(eBARS_Symbol)

    If (g.RealTime.SalmonIsRunning = True) And (InStr(strSymbol, "|") = 0) Then
        dBidPrice = g.RealTime.LastKnownPrice(Bars.SymbolOrSymbolID, -1, False, dBidTime)
        dTickPrice = g.RealTime.LastKnownPrice(Bars.SymbolOrSymbolID, 0, False, dTickTime)
    Else
        dBidTime = 0
        dBidPrice = Bars(eBARS_Bid, Bars.Size - 1)
        dTickTime = 0
        dTickPrice = Bars(eBARS_Close, Bars.Size - 1)
        
        ' For pit-session futures, if the last is valid then we don't want to believe the
        ' bid price because it could be quite old...
        If IsPitFuture(Bars.Prop(eBARS_Symbol)) And (dTickPrice <> kNullData) Then
            dBidPrice = kNullData
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeStream.BidSidePrice"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AskSidePrice
'' Description: Determine the prices on the ask side of the market
'' Inputs:      Bars, Ask Price, Ask Time, Last Tick Price, Last Tick Time
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AskSidePrice(ByVal Bars As cGdBars, dAskPrice As Double, dAskTime As Double, dTickPrice As Double, dTickTime As Double)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol for the Bars
    
    strSymbol = Bars.Prop(eBARS_Symbol)

    If (g.RealTime.SalmonIsRunning = True) And (InStr(strSymbol, "|") = 0) Then
        dAskPrice = g.RealTime.LastKnownPrice(Bars.SymbolOrSymbolID, 1, False, dAskTime)
        dTickPrice = g.RealTime.LastKnownPrice(Bars.SymbolOrSymbolID, 0, False, dTickTime)
    Else
        dAskTime = 0
        dAskPrice = Bars(eBARS_Ask, Bars.Size - 1)
        dTickTime = 0
        dTickPrice = Bars(eBARS_Close, Bars.Size - 1)
        
        ' For pit-session futures, if the last is valid then we don't want to believe the
        ' ask price because it could be quite old...
        If IsPitFuture(Bars.Prop(eBARS_Symbol)) And (dTickPrice <> kNullData) Then
            dAskPrice = kNullData
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeStream.AskSidePrice"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StopPriceHit
'' Description: Determine if the stop price for an order has been hit
'' Inputs:      Stop Price, Buy?, Bars, Order Time
'' Returns:     True if Hit, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function StopPriceHit(ByVal dStopPrice As Double, ByVal bIsBuy As Boolean, Bars As cGdBars, ByVal dOrderTime As Double) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim dAskTime As Double              ' Time of the last ask
    Dim dAskPrice As Double             ' Price of the last ask
    Dim dBidTime As Double              ' Time of the last bid
    Dim dBidPrice As Double             ' Price of the last bid
    Dim dTickTime As Double             ' Time of the last tick
    Dim dTickPrice As Double            ' Price of the last tick
    
    dOrderTime = ConvertTimeZone(dOrderTime, TimeZone(Bars.Prop(eBARS_Symbol)), Bars.Prop(eBARS_ExchangeTimeZoneInf))

    bReturn = False
    If bIsBuy = True Then
        BidSidePrice Bars, dBidPrice, dBidTime, dTickPrice, dTickTime
        
        If g.RealTime.SalmonIsRunning Then
            If (dBidPrice <> kNullData) And (dBidTime >= dOrderTime) Then
                bReturn = (dBidPrice >= dStopPrice)
            End If
            If bReturn = False Then
                If (dTickPrice <> kNullData) And (dTickTime >= dOrderTime) Then
                    bReturn = (dTickPrice >= dStopPrice)
                End If
            End If
        Else
            If (dBidPrice <> kNullData) And (dBidTime >= dTickTime) Then
                bReturn = (dBidPrice >= dStopPrice)
            ElseIf dTickPrice <> kNullData Then
                bReturn = (dTickPrice >= dStopPrice)
            End If
        End If
    Else
        AskSidePrice Bars, dAskPrice, dAskTime, dTickPrice, dTickTime
            
        If g.RealTime.SalmonIsRunning Then
            If (dAskPrice <> kNullData) And (dAskTime >= dOrderTime) Then
                bReturn = (dAskPrice <= dStopPrice)
            End If
            If bReturn = False Then
                If (dTickPrice <> kNullData) And (dTickTime >= dOrderTime) Then
                    bReturn = (dTickPrice <= dStopPrice)
                End If
            End If
        Else
            If (dAskPrice <> kNullData) And (dAskTime >= dTickTime) Then
                bReturn = (dAskPrice <= dStopPrice)
            ElseIf dTickPrice <> kNullData Then
                bReturn = (dTickPrice <= dStopPrice)
            End If
        End If
    End If
    
    StopPriceHit = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.StopPriceHit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LimitOrderFilled
'' Description: Determine if the Limit price for an order has been hit
'' Inputs:      Limit Price, Buy?, Bars
'' Returns:     Fill Price or kNullData if not filled
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LimitOrderFilled(ByVal dLimitPrice, ByVal bIsBuy As Boolean, Bars As cGdBars) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim dAskTime As Double              ' Time of the last ask
    Dim dAskPrice As Double             ' Price of the last ask
    Dim dBidTime As Double              ' Time of the last bid
    Dim dBidPrice As Double             ' Price of the last bid
    Dim dTickTime As Double             ' Time of the last tick
    Dim dTickPrice As Double            ' Price of the last tick

    dReturn = kNullData
    If bIsBuy Then
        AskSidePrice Bars, dAskPrice, dAskTime, dTickPrice, dTickTime
            
        If (dAskPrice <> kNullData) And (dAskTime >= dTickTime) Then
            If dAskPrice <= dLimitPrice Then
                dReturn = dAskPrice
            End If
        ElseIf dTickPrice <> kNullData Then
            If dTickPrice < dLimitPrice Then
                dReturn = dTickPrice + Bars.TickMove ' dLimitPrice
            End If
        End If
    Else
        BidSidePrice Bars, dBidPrice, dBidTime, dTickPrice, dTickTime
            
        If (dBidPrice <> kNullData) And (dBidTime >= dTickTime) Then
            If dBidPrice >= dLimitPrice Then
                dReturn = dBidPrice
            End If
        ElseIf dTickPrice <> kNullData Then
            If dTickPrice > dLimitPrice Then
                dReturn = dTickPrice - Bars.TickMove ' dLimitPrice
            End If
        End If
    End If
    
    LimitOrderFilled = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.LimitOrderFilled"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LimitOrderFilledOptions
'' Description: Determine if the Limit price for an option order has been hit
'' Inputs:      Limit Price, Buy?, Bars
'' Returns:     Fill Price or kNullData if not filled
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LimitOrderFilledOptions(ByVal dLimitPrice, ByVal bIsBuy As Boolean, Bars As cGdBars) As Double
On Error GoTo ErrSection:

    Dim dBid As Double                  ' Bid price for the option
    Dim dAsk As Double                  ' Ask price for the option
    Dim dReturn As Double               ' Return value for the function
    Dim dMarket As Double               ' Market price
    Dim nFillMethod As eGDOptionFill    ' Method for filling an option order
    
    nFillMethod = g.Broker.OptionFillMethod
    dBid = Bars(eBARS_Bid, Bars.Size - 1)
    dAsk = Bars(eBARS_Ask, Bars.Size - 1)
    dReturn = kNullData
    
    If (nFillMethod = eGDOptionFill_BidOrAsk) Or (dBid = kNullData) Or (dAsk = kNullData) Then
        dReturn = LimitOrderFilled(dLimitPrice, bIsBuy, Bars)
    Else
        If nFillMethod = eGDOptionFill_Midpoint Then
            dMarket = Bars.RoundToPrice((dBid + dAsk) / 2)
        Else
            If bIsBuy Then
                dMarket = Bars.RoundToPrice(dBid + ((dAsk - dBid) * 2 / 3))
            Else
                dMarket = Bars.RoundToPrice(dBid + ((dAsk - dBid) / 3))
            End If
        End If
    
        If bIsBuy Then
            If dMarket <= dLimitPrice Then
                dReturn = dLimitPrice
            End If
        Else
            If dMarket >= dLimitPrice Then
                dReturn = dLimitPrice
            End If
        End If
    End If
    
    LimitOrderFilledOptions = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.LimitOrderFilledOptions"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MITPriceHit
'' Description: Determine if the MIT price for an order has been hit
'' Inputs:      MIT Price, Buy?, Bars
'' Returns:     True if Hit, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MITPriceHit(ByVal dMITPrice As Double, ByVal bIsBuy As Boolean, Bars As cGdBars) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim dAskTime As Double              ' Time of the last ask
    Dim dAskPrice As Double             ' Price of the last ask
    Dim dBidTime As Double              ' Time of the last bid
    Dim dBidPrice As Double             ' Price of the last bid
    Dim dTickTime As Double             ' Time of the last tick
    Dim dTickPrice As Double            ' Price of the last tick

    bReturn = False
    If bIsBuy = True Then
        BidSidePrice Bars, dBidPrice, dBidTime, dTickPrice, dTickTime
        
        If (dBidPrice <> kNullData) And (dBidTime >= dTickTime) Then
            bReturn = (dBidPrice <= dMITPrice)
        ElseIf dTickPrice <> kNullData Then
            bReturn = (dTickPrice <= dMITPrice)
        End If
    Else
        AskSidePrice Bars, dAskPrice, dAskTime, dTickPrice, dTickTime
        
        If (dAskPrice <> kNullData) And (dAskTime >= dTickTime) Then
            bReturn = (dAskPrice >= dMITPrice)
        ElseIf dTickPrice <> kNullData Then
            bReturn = (dTickPrice >= dMITPrice)
        End If
    End If
    
    MITPriceHit = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeStream.MITPriceHit"
    
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
    Set m.SimTradeBase = New cSimTradeBase
    
    Set m.astrLoggedNotValid = New cGdArray
    m.astrLoggedNotValid.Create eGDARRAY_Strings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeStream.Class_Initialize"
    
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
    Set m.astrLoggedNotValid = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeStream.Class_Terminate"
    
End Sub

