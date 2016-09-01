VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cCattle.cls
'' Description: Main controller for Cattle stuff
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/04/2012   DAJ         Turnkey administration / Split from broker stand-alone
'' 06/05/2012   DAJ         Finished off the Turnkey administration stuff
'' 06/07/2012   DAJ         Reconnect code
'' 06/11/2012   DAJ         Make Turnkey work with all brokers
'' 06/25/2012   DAJ         Visible Columns mode for Turnkey Admin, Feedyard hidden columns
'' 08/21/2012   DAJ         Added FcmAccountNumber for accounts
'' 09/11/2012   DAJ         New Turnkey Enums, Server DB redesign, Assign parts of fills
'' 09/14/2012   DAJ         Positions, Visible Lot Column by Genesis Customer not Feedyard Customer
'' 09/20/2012   DAJ         Changed DeleteTrades to DeleteTrade (only delete one trade)
'' 10/16/2012   DAJ         Send password on a connect message
'' 10/22/2012   DAJ         Pass TradeNav version and Machine ID on connect call
'' 01/30/2013   DAJ         Live/Demo/Test modes for Turnkey
'' 11/15/2013   DAJ         Allow for adding/editing feedyard/lots/customers
'' 11/20/2013   DAJ         Reports
'' 11/26/2013   DAJ         Exposed feedyards collection
'' 12/03/2013   DAJ         Default visible lot columns; Turnkey mode
'' 12/04/2013   DAJ         Detail Options
'' 12/18/2013   DAJ         Trim password before sending off to Turnkey servers
'' 01/03/2014   DAJ         "Either Feedyard" mode
'' 01/31/2014   DAJ         Manage feedyards and feedyard customers
'' 02/25/2014   DAJ         Rations/Ingredients; Commissions on fills
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
'' 03/10/2014   DAJ         Pass the FeedYardID in when deleting a feedyard customer
'' 03/13/2014   DAJ         Allow brokers to associate carried fill information
'' 03/14/2014   DAJ         Added support for Boolean lot column type; Dump DLL version on connect
'' 03/17/2014   DAJ         Renamed Turnkey to Cattle for admin stuff
'' 04/15/2014   DAJ         New lot column administration; Feedyard Customers;
''                          Formatting for "link" fields
'' 04/21/2014   DAJ         Map customers and feedyards back to one enum
'' 05/15/2014   DAJ         Implemented DeletePosition callback from server
'' 05/22/2014   DAJ         Renamed frmTurnkeyFeedYardCustomer to frmFeedYardCustomer; Renamed
''                          frmTurnkeyFeedYard to frmFeedYard; Renamed frmTurnkeyManage to frmCattleManage
'' 05/22/2014   DAJ         Renamed cTurnkey to cCattle; Renamed frmTurnkeyEditor to frmCattleEditor;
''                          Renamed frmTurnkey to frmLots; Renamed g.Turnkey to g.Cattle
'' 05/30/2014   DAJ         Utilized new accounts object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    FeedYards As cGdTree                ' Collection of feedyards
    VisibleCustomers As cGdTree         ' Collection of visible feedyard customers
    FeedyardCustomers As cGdTree        ' Collection of feedyard customers
    Lots As cGdTree                     ' Lots for the customer
    Positions As cGdTree                ' Collection of positions per lot
    LotColumns As cGdTree               ' Collection of lot columns
    LotColumnMap As cGdTree             ' Map of lot column key-value name to ID
    LotColumnCategories As cGdTree      ' Collection of lot column categories
    LotColumnSubCategories As cGdTree   ' Collection of lot column subcategories
    DetailOptions As cGdTree            ' Collection of lot column detail options
    Accounts As cAccounts               ' Collection of accounts
    Ingredients As cGdTree              ' Collection of ingredients
    Rations As cGdTree                  ' Collection of rations
    FillQuantities As cFillQuantities   ' Collection of fill quantities
    
    nStatus As eGDConnectionStatus      ' Connection status
    bTurnkey As Boolean                 ' User doing Turnkey ( as opposed to Cattle Navigator )?
    nType As eGDCattleCustomerType      ' Customer type
    strAppMailTarget As String          ' App-mail target based on broker
    strProcessName As String            ' Name of the process to check
    strProcessPath As String            ' Full path of the process to start
    strControlID As String              ' Control ID
    bReconnect As Boolean               ' Do we want to reconnect?
    strMode As String                   ' (L)ive or (D)emo mode?
    bCanEditLots As Boolean             ' Can the user edit the lots?
    lRequestCounter As Long             ' Request counter
    
    WaitAddCustomer As cGdTree          ' Collection of new customers waiting to be sent
End Type
Private m As mPrivate

Public Property Get ConnectionStatus() As eGDConnectionStatus
    ConnectionStatus = m.nStatus
End Property
Private Property Let ConnectionStatus(ByVal nNewStatus As eGDConnectionStatus)
    If nNewStatus <> m.nStatus Then
        m.nStatus = nNewStatus
        
        If m.nStatus = eGDConnectionStatus_Disconnected Then
            UnloadApp m.bReconnect
        End If
    End If
End Property

Public Property Get Lots() As cGdTree
    Set Lots = m.Lots
End Property

Public Property Get Customers() As cGdTree
    Set Customers = m.VisibleCustomers
End Property

Public Property Get FeedyardCustomers() As cGdTree
    Set FeedyardCustomers = m.FeedyardCustomers
End Property

Public Property Get Mode() As String
    Mode = m.strMode
End Property

Public Property Get CanEditLots() As Boolean
    CanEditLots = m.bCanEditLots
End Property

Public Property Get FeedYards() As cGdTree
    Set FeedYards = m.FeedYards
End Property

Public Property Get Ingredients() As cGdTree
    Set Ingredients = m.Ingredients
End Property

Public Property Get Rations() As cGdTree
    Set Rations = m.Rations
End Property

Public Property Get LotColumns() As cGdTree
    Set LotColumns = m.LotColumns
End Property

Public Property Get LotColumnMap() As cGdTree
    Set LotColumnMap = m.LotColumnMap
End Property

Public Property Get LotColumnCategories() As cGdTree
    Set LotColumnCategories = m.LotColumnCategories
End Property

Public Property Get LotColumnSubCategories() As cGdTree
    Set LotColumnSubCategories = m.LotColumnSubCategories
End Property

Public Property Get FillQuantities() As cFillQuantities
    Set FillQuantities = m.FillQuantities
End Property

Public Property Get Turnkey() As Boolean
    Turnkey = m.bTurnkey
End Property
Public Property Let Turnkey(ByVal bTurnkey As Boolean)
    m.bTurnkey = bTurnkey
End Property

Private Property Get PositionsKey(ByVal strFeedYardLotID As String, ByVal strBrokerAccountID As String, ByVal strSymbol As String)
    PositionsKey = strFeedYardLotID & "|" & strBrokerAccountID & "|" & strSymbol
End Property

Private Property Get GenesisID() As Long
    GenesisID = (g.lDataServiceID \ 1000) Mod 1000000
End Property

Public Property Get SelectedFeedYard() As cBrokerMessage
    If FormIsLoaded("frmLots") Then
        If m.FeedYards.Exists(Str(frmLots.SelectedFeedYard)) Then
            Set SelectedFeedYard = m.FeedYards(Str(frmLots.SelectedFeedYard))
        Else
            Set SelectedFeedYard = Nothing
        End If
    Else
        Set SelectedFeedYard = Nothing
    End If
End Property

Public Property Get GridValue(Grid As VSFlexGrid, ByVal Row As Long, ByVal Col As Long, ByVal LotColumn As cLotColumn) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the property

    If UCase(LotColumn.Format) = "BOOLEAN" Then
        strReturn = BoolToString(CheckedCell(Grid, Row, Col))
    ElseIf (UCase(LotColumn.Format) = "TEXT") Or (UCase(LotColumn.Format) = "LINK") Then
        strReturn = Grid.TextMatrix(Row, Col)
    Else
        strReturn = Str(ValOfText(Grid.TextMatrix(Row, Col)))
    End If
    
    GridValue = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCattle.GridValue.Get"
    
End Property
Public Property Let GridValue(Grid As VSFlexGrid, ByVal Row As Long, ByVal Col As Long, ByVal LotColumn As cLotColumn, ByVal strValue As String)
On Error GoTo ErrSection:

    If UCase(LotColumn.Format) = "BOOLEAN" Then
        CheckedCell(Grid, Row, Col) = StringToBool(strValue)
        Grid.Cell(flexcpPictureAlignment, Row, Col) = flexAlignCenterCenter
    ElseIf (UCase(LotColumn.Format) = "DATE") And (Val(strValue) = 0) Then
        Grid.TextMatrix(Row, Col) = ""
    ElseIf (Len(LotColumn.DisplayFormat) > 0) And (Len(strValue) > 0) Then
        Grid.TextMatrix(Row, Col) = Format(Val(strValue), LotColumn.DisplayFormat)
    Else
        Grid.TextMatrix(Row, Col) = strValue
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCattle.GridValue.Let"
    
End Property

Public Property Get Accounts() As cAccounts
    Set Accounts = m.Accounts
End Property

Public Function BoolToString(ByVal bValue As Boolean) As String
    If bValue Then
        BoolToString = "1"
    Else
        BoolToString = "0"
    End If
End Function
Public Function StringToBool(ByVal strValue As String) As Boolean
    StringToBool = (strValue = "1")
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ProductName
'' Description: Determine the product name based on enablements
'' Inputs:      None
'' Returns:     Product Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ProductName() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    If Turnkey Then
        strReturn = kTurnkeyCompanyName
    Else
        strReturn = "Cattle Navigator"
    End If
    
    ProductName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.ProductName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IconName
'' Description: Determine the icon name based on enablements
'' Inputs:      None
'' Returns:     Icon Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IconName(Optional ByVal bWhiteBackground As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    If Turnkey Then
        If bWhiteBackground Then
            strReturn = "kHedgeLincW"
        Else
            strReturn = "kHedgeLinc"
        End If
    Else
        strReturn = "kCattle"
    End If
    
    IconName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IconName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTurnkeyUser
'' Description: Determine if the user is entitled for Turnkey ( Live )
'' Inputs:      None
'' Returns:     True if entitled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsTurnkeyUser() As Boolean
On Error GoTo ErrSection:

    IsTurnkeyUser = IsTurnkeyLiveUser Or IsTurnkeyDemoUser Or IsTurnkeyTestUser

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IsTurnkeyUser"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTurnkeyLiveUser
'' Description: Determine if the user is entitled for Turnkey ( Live )
'' Inputs:      None
'' Returns:     True if entitled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsTurnkeyLiveUser() As Boolean
On Error GoTo ErrSection:

    IsTurnkeyLiveUser = g.AppBridge.HasModule("TURNKEY") Or FileExist(AddSlash(g.strAppPath) & "Turnkey.FLG")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IsTurnkeyLiveUser"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTurnkeyDemoUser
'' Description: Determine if the user is entitled for Turnkey ( Demo )
'' Inputs:      None
'' Returns:     True if entitled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsTurnkeyDemoUser() As Boolean
On Error GoTo ErrSection:

    IsTurnkeyDemoUser = g.AppBridge.HasModule("TKEYDEMO") Or FileExist(AddSlash(g.strAppPath) & "TkeyDemo.FLG")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IsTurnkeyDemoUser"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTurnkeyTestUser
'' Description: Determine if the user is entitled for Turnkey ( Test )
'' Inputs:      None
'' Returns:     True if entitled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsTurnkeyTestUser() As Boolean
On Error GoTo ErrSection:

    IsTurnkeyTestUser = g.AppBridge.HasModule("TKEYTEST") Or FileExist(AddSlash(g.strAppPath) & "TkeyTest.FLG")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IsTurnkeyTestUser"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTurnkeyAdminUser
'' Description: Determine if the user is entitled for Turnkey administration
'' Inputs:      None
'' Returns:     True if entitled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsTurnkeyAdminUser() As Boolean
On Error GoTo ErrSection:

    IsTurnkeyAdminUser = g.AppBridge.HasModule("TKEYADM") Or FileExist(AddSlash(g.strAppPath) & "TkeyAdm.FLG")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IsTurnkeyAdminUser"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsCattleUser
'' Description: Determine if the user is entitled for Cattle ( Live )
'' Inputs:      None
'' Returns:     True if entitled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsCattleUser() As Boolean
On Error GoTo ErrSection:

    IsCattleUser = IsCattleLiveUser Or IsCattleDemoUser Or IsCattleTestUser

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IsCattleUser"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsCattleLiveUser
'' Description: Determine if the user is entitled for Cattle ( Live )
'' Inputs:      None
'' Returns:     True if entitled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsCattleLiveUser() As Boolean
On Error GoTo ErrSection:

    IsCattleLiveUser = g.AppBridge.HasModule("CTTLLIVE") Or FileExist(AddSlash(g.strAppPath) & "CttlLive.FLG")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IsCattleLiveUser"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsCattleDemoUser
'' Description: Determine if the user is entitled for Cattle ( Demo )
'' Inputs:      None
'' Returns:     True if entitled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsCattleDemoUser() As Boolean
On Error GoTo ErrSection:

    IsCattleDemoUser = g.AppBridge.HasModule("CTTLDEMO") Or FileExist(AddSlash(g.strAppPath) & "CttlDemo.FLG")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IsCattleDemoUser"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsCattleTestUser
'' Description: Determine if the user is entitled for Cattle ( Test )
'' Inputs:      None
'' Returns:     True if entitled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsCattleTestUser() As Boolean
On Error GoTo ErrSection:

    IsCattleTestUser = g.AppBridge.HasModule("CTTLTEST") Or FileExist(AddSlash(g.strAppPath) & "CttlTest.FLG")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IsCattleTestUser"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Connect
'' Description: Connect to the Genesis Turnkey server
'' Inputs:      From Admin Form?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Connect(Optional ByVal bAdmin As Boolean = False)
On Error GoTo ErrSection:

    Dim bLiveUser As Boolean            ' Is this a 'Turnkey Live' user?
    Dim bDemoUser As Boolean            ' Is this a 'Turnkey Demo' user?
    Dim bTestUser As Boolean            ' Is this a 'Turnkey Test' user?

    If (ConnectionStatus <> eGDConnectionStatus_Connected) And (ConnectionStatus <> eGDConnectionStatus_Connecting) Then
        If bAdmin Then
            bLiveUser = IsTurnkeyLiveUser Or IsCattleLiveUser
            bDemoUser = IsTurnkeyDemoUser Or IsCattleDemoUser
            bTestUser = IsTurnkeyTestUser Or IsCattleTestUser
        ElseIf Turnkey Then
            bLiveUser = IsTurnkeyLiveUser
            bDemoUser = IsTurnkeyDemoUser
            bTestUser = IsTurnkeyTestUser
        Else
            bLiveUser = IsCattleLiveUser
            bDemoUser = IsCattleDemoUser
            bTestUser = IsCattleTestUser
        End If
                
        If bLiveUser And bDemoUser And bTestUser Then
            m.strMode = InfBox("Do you want to login to Live, Demo, or Test server?", "?", "+Live|-Demo|Test", ProductName)
        ElseIf bLiveUser And bDemoUser Then
            m.strMode = InfBox("Do you want to login to Live or Demo server?", "?", "+Live|-Demo", ProductName)
        ElseIf bLiveUser And bTestUser Then
            m.strMode = InfBox("Do you want to login to Live or Test server?", "?", "+Live|-Test", ProductName)
        ElseIf bDemoUser And bTestUser Then
            m.strMode = InfBox("Do you want to login to Demo or Test server?", "?", "+Demo|-Test", ProductName)
        ElseIf bLiveUser Then
            m.strMode = "L"
        ElseIf bDemoUser Then
            m.strMode = "D"
        ElseIf bTestUser Then
            m.strMode = "T"
        End If
        
        StartStandAlone
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.Connect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disconnect
'' Description: Disconnect from the Genesis Turnkey server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Disconnect()
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    
    m.bReconnect = False
    SendTurnkeyMessage eGDCattleMessage_Disconnect, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.Disconnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnloadApp
'' Description: Tell the stand-alone application to unload
'' Inputs:      Reconnect?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UnloadApp(ByVal bReconnect As Boolean)
On Error GoTo ErrSection:
    
    m.bReconnect = bReconnect And (FormIsLoaded("frmLots") Or FormIsLoaded("frmCattleAdmin"))
    SendTurnkeyMessageString eGDCattleMessage_UnloadApp, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UnloadApp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFeedyards
'' Description: Get the feedyards that the user can see
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetFeedyards()
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "GenesisID", Str(GenesisID)
    
    SendTurnkeyMessage eGDCattleMessage_GetFeedYards, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetFeedyards"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetCustomers
'' Description: Get the customers that the user can see
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetCustomers(ByVal lFeedYardID As Long)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "GenesisID", Str(GenesisID)
    turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
        
    SendTurnkeyMessage eGDCattleMessage_GetCustomers, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFeedyardCustomers
'' Description: Get the all of the customers for the given feedyard
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetFeedyardCustomers(ByVal lFeedYardID As Long)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
        
    SendTurnkeyMessage eGDCattleMessage_GetFeedyardCustomers, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetFeedyardCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLots
'' Description: Get the lots for the given feed yard
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetLots(ByVal lFeedYardID As Long)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
        
    SendTurnkeyMessage eGDCattleMessage_GetLots, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetLots"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Get the accounts for the given feed yard
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAccounts(ByVal lFeedYardID As Long)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
        
    SendTurnkeyMessage eGDCattleMessage_GetAccounts, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Get the orders for the given feed yard
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetOrders(ByVal lFeedYardID As Long)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
        
    SendTurnkeyMessage eGDCattleMessage_GetOrders, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllFills
'' Description: Get all of the fills for the given feed yard
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAllFills(ByVal lFeedYardID As Long)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
        
    SendTurnkeyMessage eGDCattleMessage_GetAllFills, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetAllFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAssociatedFills
'' Description: Get the fill associations for the given feed yard
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAssociatedFills(ByVal lFeedYardID As Long)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
        
    SendTurnkeyMessage eGDCattleMessage_GetAssociatedFills, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetAssociatedFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTrades
'' Description: Get the trades for the given feed yard
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetTrades(ByVal lFeedYardID As Long)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
        
    SendTurnkeyMessage eGDCattleMessage_GetTrades, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetTrades"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositions
'' Description: Get the positions for the given feed yard
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPositions(ByVal lFeedYardID As Long)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
        
    SendTurnkeyMessage eGDCattleMessage_GetPositions, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetVisibleLots
'' Description: Get the visible lots for the given feed yard
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetVisibleLots(ByVal lFeedYardID As Long)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
        
    SendTurnkeyMessage eGDCattleMessage_GetVisibleLots, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetVisibleLots"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetGenesisCustomers
'' Description: Get a list of all known Genesis customers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetGenesisCustomers()
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    
    SendTurnkeyMessage eGDCattleMessage_GetGenesisCustomers, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetGenesisCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllFeedYards
'' Description: Get a list of all known feed yards
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAllFeedyards()
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    
    SendTurnkeyMessage eGDCattleMessage_GetAllFeedYards, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetAllFeedYards"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllCustomers
'' Description: Get a list of all known customers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAllCustomers()
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    
    SendTurnkeyMessage eGDCattleMessage_GetAllCustomers, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetAllCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetVisibleFeedYards
'' Description: Get a list of all visible feed yards for the given Genesis ID
'' Inputs:      Genesis ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetVisibleFeedyards(ByVal strGenesisID As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "ID", strGenesisID
    
    SendTurnkeyMessage eGDCattleMessage_GetVisibleFeedYards, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetVisibleFeedYards"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetVisibleCustomers
'' Description: Get a list of all visible customers for the given Genesis ID
'' Inputs:      Genesis ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetVisibleCustomers(ByVal strGenesisID As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "ID", strGenesisID
    
    SendTurnkeyMessage eGDCattleMessage_GetVisibleCustomers, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetVisibleCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllLotColumns
'' Description: Get a list of all columns for the lots grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAllLotColumns(ByVal lFeedYardID As Long)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    If lFeedYardID > -1& Then
        Set turnkeyMessage = New cBrokerMessage
        turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
        
        SendTurnkeyMessage eGDCattleMessage_GetAllLotColumns, turnkeyMessage
    Else
        SendTurnkeyMessageString eGDCattleMessage_GetAllLotColumns, ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetAllLotColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetVisibleLotColumns
'' Description: Get a list of all visible lot columns for the feedyard customer
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetVisibleLotColumns()
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim FeedYard As cBrokerMessage      ' Currently selected feedyard
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "GenesisID", Str(GenesisID)
    
    Select Case m.nType
        Case eGDCattleCustomerType_TurnkeyCustomer, eGDCattleCustomerType_CattleNavCustomer
            turnkeyMessage.Add "Type", Str(eGDCattleCustomerType_TurnkeyCustomer)
        
        Case eGDCattleCustomerType_TurnkeyFeedYard, eGDCattleCustomerType_CattleNavFeedYard, eGDCattleCustomerType_EitherFeedYard
            turnkeyMessage.Add "Type", Str(eGDCattleCustomerType_TurnkeyFeedYard)
        
        Case Else
            turnkeyMessage.Add "Type", Str(m.nType)
            
    End Select
        
    Set FeedYard = SelectedFeedYard
    If Not FeedYard Is Nothing Then
        turnkeyMessage.Add "FeedYardID", FeedYard("ID")
        turnkeyMessage.Add "FeedYardSource", FeedYard("FeedYardSource")
    End If
    
    SendTurnkeyMessage eGDCattleMessage_GetVisibleLotColumns, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetVisibleLotColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllLotColumnsAdmin
'' Description: Get a list of all columns for the lots grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAllLotColumnsAdmin()
On Error GoTo ErrSection:

    SendTurnkeyMessageString eGDCattleMessage_GetAllLotColumnsAdmin, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetAllLotColumnsAdmin"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetVisibleLotColumnsAdmin
'' Description: Get a list of all visible lot columns for the feedyard customer
'' Inputs:      Genesis ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetVisibleLotColumnsAdmin(ByVal strGenesisID As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "ID", strGenesisID
    
    SendTurnkeyMessage eGDCattleMessage_GetVisibleLotColumnsAdmin, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetVisibleLotColumnsAdmin"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDefaultVisibleLotColumnsAdmin
'' Description: Get a list of all default visible lot columns for the feedyard customer
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetDefaultVisibleLotColumnsAdmin()
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    
    SendTurnkeyMessage eGDCattleMessage_GetDefaultVisibleLotColumnsAdmin, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetDefaultVisibleLotColumnsAdmin"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLotColumnCategories
'' Description: Get lot column categories from the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetLotColumnCategories()
On Error GoTo ErrSection:

    SendTurnkeyMessageString eGDCattleMessage_GetLotColumnCategories, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetLotColumnCategories"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLotColumnSubCategories
'' Description: Get lot column subcategories from the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetLotColumnSubCategories()
On Error GoTo ErrSection:

    SendTurnkeyMessageString eGDCattleMessage_GetLotColumnSubCategories, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetLotColumnCategories"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLotContentDetailsForLot
'' Description: Get lot content details for a given lot ID
'' Inputs:      Feed Yard Lot ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetLotContentDetailsForLot(ByVal strFeedYardLotID As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "FeedYardLotID", strFeedYardLotID
    
    SendTurnkeyMessage eGDCattleMessage_GetLotContentDetails, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetLotContentDetailsForLot"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLotContentDetailsForYard
'' Description: Get lot content details for a given lot ID
'' Inputs:      Feed Yard Lot ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetLotContentDetailsForYard(ByVal strFeedyardID As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "FeedYardID", strFeedyardID
    
    SendTurnkeyMessage eGDCattleMessage_GetLotContentDetails, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetLotContentDetailsForYard"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDetailOptions
'' Description: Get lot content detail options
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetDetailOptions(ByVal strFeedyardID As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "FeedYardID", strFeedyardID
    
    SendTurnkeyMessage eGDCattleMessage_GetDetailOptions, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetDetailOptions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFeedyards
'' Description: Send a message to the Turnkey server to add the given feedyards
'' Inputs:      Feedyards
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddFeedyards(ByVal strFeedyards As String)
On Error GoTo ErrSection:

    SendTurnkeyMessageString eGDCattleMessage_AddFeedYards, strFeedyards

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.AddFeedyards"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddCustomers
'' Description: Send a message to the Turnkey server to add the given customers
'' Inputs:      Customers
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddCustomers(ByVal strCustomers As String)
On Error GoTo ErrSection:

    SendTurnkeyMessageString eGDCattleMessage_AddCustomers, strCustomers

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.AddCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddLots
'' Description: Send a message to the Turnkey server to add the given lots
'' Inputs:      Lots
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddLots(ByVal strLots As String)
On Error GoTo ErrSection:

    SendTurnkeyMessageString eGDCattleMessage_AddLots, strLots

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.AddLots"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateVisibleFeedyards
'' Description: Update visible feedyards for Genesis customers
'' Inputs:      Genesis ID, Visible Feedyards
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateVisibleFeedyards(ByVal strGenesisID As String, ByVal strVisibleFeedyards As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "GenesisID", strGenesisID
    turnkeyMessage.Add "FeedyardIds", strVisibleFeedyards
    
    SendTurnkeyMessage eGDCattleMessage_UpdateVisibleFeedYards, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateVisibleFeedyards"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateVisibleCustomers
'' Description: Update visible customers for Genesis customers
'' Inputs:      Genesis ID, Visible Customers
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateVisibleCustomers(ByVal strGenesisID As String, ByVal strVisibleCustomers As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "GenesisID", strGenesisID
    turnkeyMessage.Add "CustomerIds", strVisibleCustomers
    
    SendTurnkeyMessage eGDCattleMessage_UpdateVisibleCustomers, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateVisibleCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateGenesisCustomer
'' Description: Update the Genesis customer
'' Inputs:      Genesis ID, Account Number, Name, Password, Type, Can Edit Lots?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateGenesisCustomer(ByVal strGenesisID As String, ByVal strAccountNumber As String, ByVal strName As String, ByVal strPassword As String, ByVal nCustomerType As eGDCattleCustomerType, ByVal bCanEditLots As Boolean)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    If Len(strGenesisID) > 0 Then
        turnkeyMessage.Add "GenesisID", strGenesisID
    End If
    turnkeyMessage.Add "Account", strAccountNumber
    turnkeyMessage.Add "Name", strName
    turnkeyMessage.Add "Password", strPassword
    turnkeyMessage.Add "Type", Str(nCustomerType)
    
    If bCanEditLots Then
        turnkeyMessage.Add "CanEditLots", "1"
    Else
        turnkeyMessage.Add "CanEditLots", "0"
    End If
    
    SendTurnkeyMessage eGDCattleMessage_UpdateGenesisCustomer, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateGenesisCustomer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateVisibleLotColumns
'' Description: Update visible lot columns for a feedyard customer
'' Inputs:      Genesis ID, Feedyard ID, Visible Lot Columns
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateVisibleLotColumns(ByVal strGenesisID As String, ByVal strFeedyardID As String, ByVal strVisibleLotColumns As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "GenesisID", strGenesisID
    turnkeyMessage.Add "FeedYardID", strFeedyardID
    turnkeyMessage.Add "LotColumnIds", strVisibleLotColumns
    
    SendTurnkeyMessage eGDCattleMessage_UpdateVisibleLotColumns, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateVisibleLotColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateDefaultVisibleLotColumns
'' Description: Update default visible lot columns for a feedyard customer
'' Inputs:      Customer Type, Feedyard Source, Visible Lot Columns
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateDefaultVisibleLotColumns(ByVal strCustomerType As String, ByVal strFeedyardSource As String, ByVal strDefaultVisibleLotColumns As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "Type", strCustomerType
    turnkeyMessage.Add "FeedYardSource", strFeedyardSource
    turnkeyMessage.Add "LotColumnIds", strDefaultVisibleLotColumns
    
    SendTurnkeyMessage eGDCattleMessage_UpdateDefaultVisibleLotColumns, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateDefaultVisibleLotColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateAccount
'' Description: Send the given account information to the Genesis Turnkey server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateAccount(ByVal accountMessage As cBrokerMessage)
On Error GoTo ErrSection:

    SendTurnkeyMessage eGDCattleMessage_UpdateAccount, accountMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveAccount
'' Description: Remove the given account information from the Genesis Turnkey server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveAccount(ByVal accountMessage As cBrokerMessage)
On Error GoTo ErrSection:

    SendTurnkeyMessage eGDCattleMessage_RemoveAccount, accountMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.RemoveAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateOrder
'' Description: Send the given order information to the Genesis Turnkey server
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateOrder(ByVal orderMessage As cBrokerMessage)
On Error GoTo ErrSection:

    SendTurnkeyMessage eGDCattleMessage_UpdateOrder, orderMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFill
'' Description: Send the given fill information to the Genesis Turnkey server
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateFill(ByVal FillMessage As cBrokerMessage)
On Error GoTo ErrSection:

    SendTurnkeyMessage eGDCattleMessage_UpdateFill, FillMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AssociateFill
'' Description: Send the given fill association information to the Genesis Turnkey server
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AssociateFill(ByVal FillMessage As cBrokerMessage)
On Error GoTo ErrSection:

    SendTurnkeyMessage eGDCattleMessage_AssociateFill, FillMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.AssociateFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdatePosition
'' Description: Send the given position information to the Genesis Turnkey server
'' Inputs:      Position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdatePosition(ByVal positionMessage As cBrokerMessage)
On Error GoTo ErrSection:

    SendTurnkeyMessage eGDCattleMessage_UpdatePosition, positionMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdatePosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateLotContentDetails
'' Description: Update the lot content details
'' Inputs:      Details, Feed Yard ID, Feed Yard Lot ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateLotContentDetails(ByVal Details As cGdTree, ByVal strFeedyardID As String, ByVal strFeedYardLotID As String)
On Error GoTo ErrSection:

    Dim Detail As cBrokerMessage        ' Detail object
    Dim lIndex As Long                  ' Index into a for loop
    
    SendTurnkeyMessageString eGDCattleMessage_UpdateLotContentDetails, "Begin" & vbTab & "FeedYardID=" & Str(strFeedyardID) & vbTab & "FeedYardLotID=" & Str(strFeedYardLotID)
    For lIndex = 1 To Details.Count
        Set Detail = Details(lIndex)
        
        Detail.Add "FeedYardLotID", strFeedYardLotID
        SendTurnkeyMessage eGDCattleMessage_UpdateLotContentDetails, Detail
    Next lIndex
    SendTurnkeyMessageString eGDCattleMessage_UpdateLotContentDetails, "End"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateLotContentDetails"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToString
'' Description: Convert the given order to a string
'' Inputs:      Order message, FCM Account Number
'' Returns:     Order string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderToString(ByVal turnkeyMessage As cBrokerMessage, Optional ByVal strFcmAccountNumber As String = "") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim accountMessage As cBrokerMessage ' Account message from the collection
    Dim lNumLegs As Long                ' Number of legs
    Dim astrLeg As cGdArray             ' Leg information broken out
    Dim lIndex As Long                  ' Index into a for loop
    Dim lOrderQuantity As Long          ' Order quantity
    Dim lLegMultiplier As Long          ' Leg Multiplier
    Dim astrLegs As cGdArray            ' Array of leg strings
    Dim strStopPrice As String          ' Stop Price to display
    Dim strLimitPrice As String         ' Limit Price to display

    If Len(strFcmAccountNumber) = 0 Then
        strFcmAccountNumber = m.Accounts.FcmNumberForId(turnkeyMessage("BrokerAccountID"), turnkeyMessage("FcmAccountNumber"))
    End If
        
    lNumLegs = CLng(Val(turnkeyMessage("NumberOfLegs")))
    lOrderQuantity = CLng(Val(turnkeyMessage("Quantity")))
    
    Set astrLegs = New cGdArray
    astrLegs.Create eGDARRAY_Strings, lNumLegs
    
    For lIndex = 1 To lNumLegs
        Set astrLeg = New cGdArray
        astrLeg.SplitFields turnkeyMessage("Leg1"), ","
        
        If astrLeg(1) = "1" Then
            strReturn = "Buy "
        Else
            strReturn = "Sell "
        End If
        
        lLegMultiplier = CLng(Val(astrLeg(3)))
        strReturn = strReturn & Str(lOrderQuantity * lLegMultiplier) & " " & astrLeg(2)
        
        astrLegs(lIndex - 1) = strReturn
    Next lIndex
    
    If lNumLegs = 1 Then
        strReturn = astrLegs(0)
    Else
        strReturn = "(" & astrLegs.JoinFields(";") & ")"
    End If
    
    Select Case UCase(turnkeyMessage("Type"))
        Case "MARKET", Str(eTT_OrderType_Market)
            strReturn = strReturn & " at Market"
        Case "STOP", Str(eTT_OrderType_Stop)
            strStopPrice = g.AppBridge.PriceDisplay(Val(turnkeyMessage("StopPrice")), turnkeyMessage("Symbol"))
            strReturn = strReturn & " at " & strStopPrice & " STOP"
        Case "LIMIT", Str(eTT_OrderType_Limit)
            strLimitPrice = g.AppBridge.PriceDisplay(Val(turnkeyMessage("LimitPrice")), turnkeyMessage("Symbol"))
            strReturn = strReturn & " at " & strLimitPrice & " LIMIT"
        Case "STOPWITHLIMIT", Str(eTT_OrderType_StopWithLimit)
            strStopPrice = g.AppBridge.PriceDisplay(Val(turnkeyMessage("StopPrice")), turnkeyMessage("Symbol"))
            strLimitPrice = g.AppBridge.PriceDisplay(Val(turnkeyMessage("LimitPrice")), turnkeyMessage("Symbol"))
            strReturn = strReturn & " at " & strStopPrice & " STOP with a " & strLimitPrice & " LIMIT"
    End Select

    If Len(strFcmAccountNumber) > 0 Then
        strReturn = strReturn & " in account " & strFcmAccountNumber
    End If
    
    Select Case UCase(turnkeyMessage("TimeInForce"))
        Case "DAY", Str(eTT_TimeInForce_Day)
            strReturn = strReturn & " Day"
        Case "GTC", Str(eTT_TimeInForce_GTC)
            strReturn = strReturn & " GTC"
        Case "GTD", Str(eTT_TimeInForce_GTD)
            strReturn = strReturn & " Expires " & DateFormat(Val(turnkeyMessage("ExpirationDate")), MM_DD_YYYY)
    End Select
    
    strReturn = strReturn & " (" & g.BrokerEnums.OrderStatusToString(CLng(Val(turnkeyMessage("Status")))) & ")"
    
    OrderToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.OrderToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillToString
'' Description: Convert the given fill to a string
'' Inputs:      Fill message, FCM Account Number, Include Date
'' Returns:     Fill string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillToString(ByVal turnkeyMessage As cBrokerMessage, Optional ByVal strFcmAccountNumber As String = "", Optional ByVal bIncludeDate As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim accountMessage As cBrokerMessage ' Account message from the collection
    Dim strPrice As String              ' Price to display

    If Len(strFcmAccountNumber) = 0 Then
        strFcmAccountNumber = m.Accounts.FcmNumberForId(turnkeyMessage("BrokerAccountID"), turnkeyMessage("FcmAccountNumber"))
    End If
    
    If turnkeyMessage("IsBuy") = "1" Then
        strReturn = "Bought "
    Else
        strReturn = "Sold "
    End If
    
    strPrice = g.AppBridge.PriceDisplay(Val(turnkeyMessage("Price")), turnkeyMessage("Symbol"))
    
    strReturn = strReturn & turnkeyMessage("Quantity") & " " & turnkeyMessage("Symbol") & " at " & strPrice
    
    If Len(strFcmAccountNumber) > 0 Then
        strReturn = strReturn & " in account " & strFcmAccountNumber
    End If
    
    If bIncludeDate Then
        strReturn = strReturn & " at " & DateFormat(Val(turnkeyMessage("FillTime")), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER, True)
    End If
    
    FillToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.FillToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeToString
'' Description: Convert the given trade to a string
'' Inputs:      Trade message, FCM Account Number, Show Symbol?
'' Returns:     Trade string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TradeToString(ByVal turnkeyMessage As cBrokerMessage, Optional ByVal strFcmAccountNumber As String = "", Optional ByVal bShowSymbol As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim accountMessage As cBrokerMessage ' Account message from the collection
    Dim strPrice As String              ' Price to display

    If Len(strFcmAccountNumber) = 0 Then
        strFcmAccountNumber = m.Accounts.FcmNumberForId(turnkeyMessage("BrokerAccountID"), turnkeyMessage("FcmAccountNumber"))
    End If
    
    If turnkeyMessage("IsBuy") = "1" Then
        strReturn = "Bought "
    Else
        strReturn = "Sold "
    End If
    
    strPrice = g.AppBridge.PriceDisplay(Val(turnkeyMessage("EntryFillPrice")), turnkeyMessage("Symbol"))
    
    strReturn = strReturn & turnkeyMessage("Quantity") & " "
    If bShowSymbol Then
        strReturn = strReturn & turnkeyMessage("Quantity") & " " & turnkeyMessage("Symbol")
    End If
    strReturn = strReturn & " at " & strPrice
    
    If Len(strFcmAccountNumber) > 0 Then
        strReturn = strReturn & " in account " & strFcmAccountNumber
    End If
    
    strReturn = strReturn & " at " & DateFormat(Val(turnkeyMessage("EntryFillTime")), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER, True)
    
    If Len(turnkeyMessage("ExitAssociatedFillID")) > 0 Then
        If turnkeyMessage("IsBuy") = "1" Then
            strReturn = strReturn & "; Sold "
        Else
            strReturn = strReturn & "; Bought "
        End If
        
        strPrice = g.AppBridge.PriceDisplay(Val(turnkeyMessage("ExitFillPrice")), turnkeyMessage("Symbol"))
        
        strReturn = strReturn & turnkeyMessage("Quantity") & " "
        If bShowSymbol Then
            strReturn = strReturn & turnkeyMessage("Quantity") & " " & turnkeyMessage("Symbol")
        End If
        strReturn = strReturn & " at " & strPrice
        
        If Len(strFcmAccountNumber) > 0 Then
            strReturn = strReturn & " in account " & strFcmAccountNumber
        End If
        
        strReturn = strReturn & " at " & DateFormat(Val(turnkeyMessage("ExitFillTime")), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER, True)
    End If
    
    TradeToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.TradeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeEntryToString
'' Description: Convert the entry portion of the given trade to a string
'' Inputs:      Trade message, FCM Account Number, Show Symbol?, Show Date?
'' Returns:     Trade Entry string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TradeEntryToString(ByVal turnkeyMessage As cBrokerMessage, Optional ByVal strFcmAccountNumber As String = "", Optional ByVal bShowSymbol As Boolean = True, Optional ByVal bShowDate As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim accountMessage As cBrokerMessage ' Account message from the collection
    Dim strPrice As String              ' Price to display

    If Len(strFcmAccountNumber) = 0 Then
        strFcmAccountNumber = m.Accounts.FcmNumberForId(turnkeyMessage("BrokerAccountID"), turnkeyMessage("FcmAccountNumber"))
    End If
    
    If turnkeyMessage("IsBuy") = "1" Then
        strReturn = "Bought "
    Else
        strReturn = "Sold "
    End If
    
    strPrice = g.AppBridge.PriceDisplay(Val(turnkeyMessage("EntryFillPrice")), turnkeyMessage("Symbol"))
    
    strReturn = strReturn & turnkeyMessage("Quantity") & " "
    If bShowSymbol Then
        strReturn = strReturn & turnkeyMessage("Quantity") & " " & turnkeyMessage("Symbol")
    End If
    strReturn = strReturn & " at " & strPrice
    
    If Len(strFcmAccountNumber) > 0 Then
        strReturn = strReturn & " in account " & strFcmAccountNumber
    End If
    
    If bShowDate Then
        strReturn = strReturn & " at " & DateFormat(Val(turnkeyMessage("EntryFillTime")), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER, True)
    End If
    
    TradeEntryToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.TradeEntryToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeExitToString
'' Description: Convert the exit portion of the given trade to a string
'' Inputs:      Trade message, FCM Account Number, Show Symbol?, Show Date?
'' Returns:     Trade Exit string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TradeExitToString(ByVal turnkeyMessage As cBrokerMessage, Optional ByVal strFcmAccountNumber As String = "", Optional ByVal bShowSymbol As Boolean = True, Optional ByVal bShowDate As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim accountMessage As cBrokerMessage ' Account message from the collection
    Dim strPrice As String              ' Price to display

    strReturn = ""
    If Len(turnkeyMessage("ExitAssociatedFillID")) > 0 Then
        If Len(strFcmAccountNumber) = 0 Then
            strFcmAccountNumber = m.Accounts.FcmNumberForId(turnkeyMessage("BrokerAccountID"), turnkeyMessage("FcmAccountNumber"))
        End If
        
        If turnkeyMessage("IsBuy") = "1" Then
            strReturn = strReturn & "Sold "
        Else
            strReturn = strReturn & "Bought "
        End If
        
        strPrice = g.AppBridge.PriceDisplay(Val(turnkeyMessage("ExitFillPrice")), turnkeyMessage("Symbol"))
        
        strReturn = strReturn & turnkeyMessage("Quantity") & " "
        If bShowSymbol Then
            strReturn = strReturn & turnkeyMessage("Quantity") & " " & turnkeyMessage("Symbol")
        End If
        strReturn = strReturn & " at " & strPrice
        
        If Len(strFcmAccountNumber) > 0 Then
            strReturn = strReturn & " in account " & strFcmAccountNumber
        End If
        
        If bShowDate Then
            strReturn = strReturn & " at " & DateFormat(Val(turnkeyMessage("ExitFillTime")), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER, True)
        End If
    End If
    
    TradeExitToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.TradeExitToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionToString
'' Description: Convert the given position to a string
'' Inputs:      Position message, FCM Account number
'' Returns:     Position string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionToString(ByVal turnkeyMessage As cBrokerMessage, Optional ByVal strFcmAccountNumber As String = "") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim accountMessage As cBrokerMessage ' Account message from the collection
    Dim lPosition As Long               ' Current position for the symbol

    If Len(strFcmAccountNumber) = 0 Then
        strFcmAccountNumber = m.Accounts.FcmNumberForId(turnkeyMessage("BrokerAccountID"), turnkeyMessage("FcmAccountNumber"))
    End If
    
    lPosition = CLng(Val(turnkeyMessage("Quantity")))
    If lPosition = 0 Then
        strReturn = "Flat"
    ElseIf lPosition > 0 Then
        strReturn = "Long " & Str(lPosition)
    Else
        strReturn = "Short " & Str(Abs(lPosition))
    End If
    
    strReturn = strReturn & " " & turnkeyMessage("Symbol")

    If Len(strFcmAccountNumber) > 0 Then
        strReturn = strReturn & " in account " & strFcmAccountNumber
    End If
    
    PositionToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.PositionToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LotIdForNumber
'' Description: Determine the Lot ID for the given Lot number
'' Inputs:      Lot Number
'' Returns:     Lot ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LotIdForNumber(ByVal strLotNumber As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim turnkeyMsg As cBrokerMessage    ' Lot information from the collection
    
    strReturn = ""
    
    For lIndex = 1 To m.Lots.Count
        Set turnkeyMsg = m.Lots(lIndex)
        If turnkeyMsg("Number") = strLotNumber Then
            strReturn = turnkeyMsg("FeedYardLotID")
            Exit For
        End If
    Next lIndex
    
    LotIdForNumber = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.LotIdForNumber"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Send a string to the log file for the day
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DumpDebug(ByVal strMessage As String)
On Error Resume Next

#If 0 Then

    Dim fh As Integer                   ' File handle to open file with
    fh = FreeFile
    Open AddSlash(g.strAppPath) & "Turnkey\TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ") - " & strMessage
        Close #fh
    End If

#Else

    Static LogFile As cLogFile
    If LogFile Is Nothing Then
        Set LogFile = New cLogFile
        LogFile.OpenFile AddSlash(g.strAppPath) & "Turnkey\TN*.LOG"
    End If
    LogFile.WriteText strMessage

#End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LotNumberForID
'' Description: Determine the Lot Number for the given Lot ID
'' Inputs:      Lot ID
'' Returns:     Lot Number
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LotNumberForID(ByVal strFeedYardLotID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim Lot As cBrokerMessage           ' Lot
    
    strReturn = ""
    If m.Lots.Exists(strFeedYardLotID) Then
        Set Lot = m.Lots(strFeedYardLotID)
        strReturn = Lot("Number")
    End If
    
    LotNumberForID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.LotNumberForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LotDisplayForID
'' Description: Determine the Lot display for the given Lot ID
'' Inputs:      Lot ID
'' Returns:     Lot Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LotDisplayForID(ByVal strFeedYardLotID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim Lot As cBrokerMessage           ' Lot
    
    strReturn = ""
    If m.Lots.Exists(strFeedYardLotID) Then
        Set Lot = m.Lots(strFeedYardLotID)
        strReturn = LotDisplay(Lot)
    End If
    
    LotDisplayForID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.LotDisplayForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LotDisplay
'' Description: Determine the Lot display for the given Lot
'' Inputs:      Lot
'' Returns:     Lot Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LotDisplay(ByVal Lot As cBrokerMessage) As String
On Error GoTo ErrSection:

    LotDisplay = Lot("Number") & " (" & Lot("OwnerName") & ")"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.LotDisplay"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadLotsCombo
'' Description: Load the lots combo with the valid values
'' Inputs:      Combo Box, Feed Yard Lot ID, Initial Entry
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadLotsCombo(cboLots As ComboBox, Optional ByVal strFeedYardLotID As String = "", Optional ByVal strInitialEntry As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim cattleMessage As cBrokerMessage ' Cattle message
    Dim lListIndex As Long              ' List index to select
    
    lListIndex = 0&
    
    cboLots.Clear
    
    If (Len(strFeedYardLotID) = 0) And (FormIsLoaded("frmLots") = True) Then
        Set cattleMessage = frmLots.SelectedLot
        If Not cattleMessage Is Nothing Then
            strFeedYardLotID = cattleMessage("FeedYardLotID")
        End If
    End If
    
    If Len(strInitialEntry) > 0 Then
        cboLots.AddItem strInitialEntry
        cboLots.ItemData(cboLots.NewIndex) = -1&
    End If
    
    For lIndex = 1 To m.Lots.Count
        Set cattleMessage = m.Lots(lIndex)
        cboLots.AddItem cattleMessage("Number") & " (" & cattleMessage("OwnerName") & ")"
        cboLots.ItemData(cboLots.NewIndex) = CLng(Val(cattleMessage("FeedYardLotID")))
        
        If strFeedYardLotID = cattleMessage("FeedYardLotID") Then
            lListIndex = lIndex
        End If
    Next lIndex
    
    cboLots.ListIndex = lListIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.LoadLotsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Position
'' Description: Determine the position for the given information
'' Inputs:      Feed Yard Lot ID, Account Number, Broker, Genesis Symbol
'' Returns:     Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Position(ByVal strFeedYardLotID As String, ByVal strAccountNumber As String, ByVal nBroker As eTT_AccountType, ByVal strGenesisSymbol As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim strBrokerAccountID As String    ' Broker Account ID
    Dim strPositionKey As String        ' Key into the positions collection
    Dim turnkeyMessage As cBrokerMessage ' Position message
    
    lReturn = 0&
    strBrokerAccountID = m.Accounts.AccountIdForBrokerNumber(strAccountNumber, nBroker)
    strPositionKey = PositionsKey(strFeedYardLotID, strBrokerAccountID, strGenesisSymbol)
    
    If m.Positions.Exists(strPositionKey) Then
        Set turnkeyMessage = m.Positions(strPositionKey)
        lReturn = CLng(Val(turnkeyMessage("Quantity")))
    End If
    
    Position = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.Position"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DetailOptions
'' Description: Get the detail options for the given lot column ID
'' Inputs:      Lot Column ID
'' Returns:     Detail Options
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DetailOptions(ByVal strLotColumnID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    strReturn = ""
    If m.DetailOptions.Exists(strLotColumnID) Then
        strReturn = m.DetailOptions(strLotColumnID)
    End If
    
    DetailOptions = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.DetailOptions"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextRequestID
'' Description: Get the next request ID
'' Inputs:      None
'' Returns:     Next Request ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextRequestID() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = g.strMachineID & "-" & Str(m.lRequestCounter)
    
    If m.lRequestCounter = 1000000 Then
        m.lRequestCounter = 1
    Else
        m.lRequestCounter = m.lRequestCounter + 1
    End If
    
    SetIniFileProperty "Counter", m.lRequestCounter, "Turnkey", g.strIniFile
    
    NextRequestID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.NextRequestID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFeedYard
'' Description: Update the given feed yard
'' Inputs:      Feed Yard
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateFeedYard(Optional ByVal FeedYard As cBrokerMessage = Nothing)
On Error GoTo ErrSection:

    Dim yardCustomer As cBrokerMessage  ' Feedyard customer
    
    If FeedYard Is Nothing Then
        Set FeedYard = New cBrokerMessage
    End If
        
    If frmFeedYard.ShowMe(FeedYard) Then
        FeedYard.Add "RequestID", NextRequestID
        
        ' If this is a new feedyard, create a default feed yard customer...
        If Len(FeedYard("ID")) = 0 Then
            Set yardCustomer = New cBrokerMessage
            yardCustomer.Add "Number", "1"
            yardCustomer.Add "Name", FeedYard("Name")
            yardCustomer.Add "FeedYardName", FeedYard("Name")
            m.WaitAddCustomer.Add yardCustomer, FeedYard("RequestID")
        End If
            
        AddFeedyards FeedYard.ToString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateFeedYard"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFeedYard
'' Description: Remove the given feed yard
'' Inputs:      Feed Yard
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveFeedYard(ByVal FeedYard As cBrokerMessage)
On Error GoTo ErrSection:

    If InfBox("Are you sure that you want to remove|'" & FeedYard("Name") & "'?", "?", "+Yes|-No", "Confirmation") = "Y" Then
        FeedYard.Add "RequestID", NextRequestID
        FeedYard.Add "Deleted", "1"
        
        AddFeedyards FeedYard.ToString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.RemoveFeedYard"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFeedYardCustomer
'' Description: Allow the user to update a customer
'' Inputs:      Customer
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateFeedYardCustomer(Optional ByVal FeedYardCustomer As cBrokerMessage = Nothing)
On Error GoTo ErrSection:

    Dim lSelectedFeedYardID As Long     ' Selected feed yard ID

    lSelectedFeedYardID = -1&
    If FormIsLoaded("frmLots") Then
        lSelectedFeedYardID = frmLots.SelectedFeedYard
    End If

    If lSelectedFeedYardID <> -1& Then
        If FeedYardCustomer Is Nothing Then
            Set FeedYardCustomer = New cBrokerMessage
        End If

        If frmFeedYardCustomer.ShowMe(FeedYardCustomer) Then
            FeedYardCustomer.Add "FeedYardID", Str(lSelectedFeedYardID)
            FeedYardCustomer.Add "RequestID", NextRequestID
            
            AddCustomers FeedYardCustomer.ToString
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateFeedYardCustomer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFeedYardCustomer
'' Description: Remove the given feed yard customer
'' Inputs:      Feed Yard Customer
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveFeedYardCustomer(ByVal FeedYardCustomer As cBrokerMessage)
On Error GoTo ErrSection:

    Dim lSelectedFeedYardID As Long     ' Selected feed yard ID

    lSelectedFeedYardID = -1&
    If FormIsLoaded("frmLots") Then
        lSelectedFeedYardID = frmLots.SelectedFeedYard
    End If

    If lSelectedFeedYardID <> -1& Then
        If InfBox("Are you sure that you want to remove|'" & FeedYardCustomer("Name") & "'?", "?", "+Yes|-No", "Confirmation") = "Y" Then
            FeedYardCustomer.Add "RequestID", NextRequestID
            FeedYardCustomer.Add "FeedYardID", Str(lSelectedFeedYardID)
            FeedYardCustomer.Add "Deleted", "1"
            
            AddCustomers FeedYardCustomer.ToString
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.RemoveFeedYardCustomer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetRations
'' Description: Ask the server for a list of rations for the feedyard
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetRations(Optional ByVal lFeedYardID As Long = -1&)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    If lFeedYardID = -1& Then
        If FormIsLoaded("frmLots") Then
            lFeedYardID = frmLots.SelectedFeedYard
        End If
    End If
    
    If lFeedYardID <> -1& Then
        Set turnkeyMessage = New cBrokerMessage
        turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
            
        SendTurnkeyMessage eGDCattleMessage_GetRations, turnkeyMessage
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetRations"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateRation
'' Description: Allow the user to update a ration
'' Inputs:      Ration
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateRation(Ration As cBrokerMessage)
On Error GoTo ErrSection:

    Ration.Add "RequestID", NextRequestID
    SendTurnkeyMessage eGDCattleMessage_UpdateRation, Ration

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateRation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetIngredients
'' Description: Ask the server for a list of ingredients for the feedyard
'' Inputs:      Feed Yard ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetIngredients(Optional ByVal lFeedYardID As Long = -1&)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    If lFeedYardID = -1& Then
        If FormIsLoaded("frmLots") Then
            lFeedYardID = frmLots.SelectedFeedYard
        End If
    End If
    
    If lFeedYardID <> -1& Then
        Set turnkeyMessage = New cBrokerMessage
        turnkeyMessage.Add "FeedYardID", Str(lFeedYardID)
            
        SendTurnkeyMessage eGDCattleMessage_GetIngredients, turnkeyMessage
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.GetIngredients"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateIngredient
'' Description: Allow the user to update an ingredient
'' Inputs:      Ration
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateIngredient(ByVal Ingredient As cBrokerMessage)
On Error GoTo ErrSection:

    Ingredient.Add "RequestID", NextRequestID
    SendTurnkeyMessage eGDCattleMessage_UpdateIngredient, Ingredient

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.UpdateIngredient"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RationNameExists
'' Description: Does the given ration name exist?
'' Inputs:      Name to Check, ID to ignore
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RationNameExists(ByVal strName As String, Optional ByVal strIgnoreID As String = "") As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Ration As cBrokerMessage        ' Ration object
    
    bReturn = False
    For lIndex = 1 To m.Rations.Count
        Set Ration = m.Rations(lIndex)
        If Ration("RationName") = strName Then
            bReturn = ((Len(strIgnoreID) = 0) Or (Ration("ID") <> strIgnoreID))
            Exit For
        End If
    Next lIndex
    
    RationNameExists = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.RationNameExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IngredientNameForID
'' Description: Determine the ingredient name for the ID
'' Inputs:      Ingredient ID
'' Returns:     Ingredient Name ( Blank if ID doesn't exist )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IngredientNameForID(ByVal strID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim Ingredient As cBrokerMessage    ' Ingredient object
    
    strReturn = ""
    If m.Ingredients.Exists(strID) Then
        Set Ingredient = m.Ingredients(strID)
        strReturn = Ingredient("Ingredient")
    End If
    
    IngredientNameForID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IngredientNameForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IngredientIDForName
'' Description: Determine the ingredient ID for the name
'' Inputs:      Ingredient Name
'' Returns:     Ingredient ID ( Blank if ID doesn't exist )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IngredientIDForName(ByVal strName As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim Ingredient As cBrokerMessage    ' Ingredient object
    Dim lIndex As Long                  ' Index into a for loop
    
    strReturn = ""
    For lIndex = 1 To m.Ingredients.Count
        Set Ingredient = m.Ingredients(lIndex)
        If UCase(Ingredient("Ingredient")) = UCase(strName) Then
            strReturn = Ingredient("ID")
            Exit For
        End If
    Next lIndex
    
    IngredientIDForName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IngredientIDForName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IngredientUsedInRations
'' Description: Determine if the ingredient ID exists in a ration
'' Inputs:      Ingredient ID
'' Returns:     True if exists in ration, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IngredientUsedInRations(ByVal strIngredientID As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Ration As cBrokerMessage        ' Ration object
    
    bReturn = False
    For lIndex = 1 To m.Rations.Count
        Set Ration = m.Rations(lIndex)
        
        If InStr("|" & Ration("IngredientID") & "|", "|" & strIngredientID & "|") > 0 Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    IngredientUsedInRations = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IngredientUsedInRations"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IngredientUsedInDetails
'' Description: Determine if the ingredient ID exists in lot content details
'' Inputs:      Ingredient Name
'' Returns:     True if exists in details, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IngredientUsedInDetails(ByVal strIngredientName As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strColumnID As String           ' Lot column ID
    Dim strDetailOptions As String      ' Detail options
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    If m.LotColumnMap.Exists("Ingredient") Then
        strColumnID = m.LotColumnMap("Ingredient")
        strDetailOptions = m.DetailOptions(strColumnID)
        
        If Len(strDetailOptions) > 0 Then
            If InStr("|" & strDetailOptions & "|", "|" & strIngredientName & "|") <> 0 Then
                bReturn = True
            End If
        End If
    End If
    
    IngredientUsedInDetails = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IngredientUsedInDetails"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveIngredientFromRations
'' Description: Remove the given ingredient ID from any rations
'' Inputs:      Ingredient Ids ( Pipe delimited )
'' Returns:     True if removed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemoveIngredientFromRations(ByVal strIngredientIds As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Ration As cBrokerMessage        ' Ration object
    Dim astrIngredient As cGdArray      ' Ingredient information split out into an array
    Dim astrPoundsFed As cGdArray       ' Pounds Fed information split out into an array
    Dim astrPctMarkup As cGdArray       ' Percent Markup informaion split out into an array
    Dim bRemoved As Boolean             ' Was the ingredient removed from the ration?
    
    bReturn = False
    For lIndex = 1 To m.Rations.Count
        Set Ration = m.Rations(lIndex)
        
        Set astrIngredient = New cGdArray
        Set astrPoundsFed = New cGdArray
        Set astrPctMarkup = New cGdArray
        
        astrIngredient.SplitFields Ration("IngredientID"), "|"
        astrPoundsFed.SplitFields Ration("PoundsFed"), "|"
        astrPctMarkup.SplitFields Ration("PercentMarkup"), "|"
        
        bRemoved = False
        For lIndex2 = astrIngredient.Size - 1 To 0 Step -1
            If InStr("|" & strIngredientIds & "|", "|" & astrIngredient(lIndex2) & "|") > 0 Then
                bRemoved = True
                
                astrIngredient.Remove lIndex2
                astrPoundsFed.Remove lIndex2
                astrPctMarkup.Remove lIndex2
            End If
        Next lIndex2
        
        If bRemoved = True Then
            bReturn = True
            
            Ration("IngredientID") = astrIngredient.JoinFields("|")
            Ration("PoundsFed") = astrPoundsFed.JoinFields("|")
            Ration("PercentMarkup") = astrPctMarkup.JoinFields("|")
            
            UpdateRation Ration
        End If
    Next lIndex
    
    RemoveIngredientFromRations = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.RemoveIngredientFromRations"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcTotalForDetails
'' Description: Calculate the total value for the given detail values
'' Inputs:      Details, Lot Column
'' Returns:     Total
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CalcTotalForDetails(ByVal Details As cGdTree, ByVal LotColumn As cLotColumn) As String
On Error GoTo ErrSection:
    
    Dim strReturn As String             ' Return value for the function
    Dim Detail As cBrokerMessage        ' Detail object
    Dim lIndex As Long                  ' Index into a for loop
    Dim dTotal As Double                ' Total of the values in the details
    Dim strText As String               ' Text for the field
    Dim bAllMatch As Boolean            ' Do all the values match?
    Dim lCount As Long                  ' Count of details for column
    Dim dFirst As Double                ' First value
    Dim dLast As Double                 ' Last value
    Dim strFirst As String              ' First value
    Dim strLast As String               ' Last value
    Dim bAllChecked As Boolean          ' Are all the booleans checked?
    Dim bAllUnchecked As Boolean        ' Are all the booleans unchecked?

    If UCase(LotColumn.Format) = "TEXT" Then
        bAllMatch = True
        lCount = 0
        strText = ""
        strFirst = ""
        strLast = ""
        
        For lIndex = 1 To Details.Count
            Set Detail = Details(lIndex)
            
            If Detail("LotColumnID") = Str(LotColumn.ID) Then
                lCount = lCount + 1
                If Len(strText) = 0 Then
                    strText = Detail("Value")
                ElseIf UCase(Detail("Value")) <> UCase(strText) Then
                    bAllMatch = False
                End If
                
                strLast = Detail("Value")
                If Len(strFirst) = 0 Then
                    strFirst = strLast
                End If
            End If
        Next lIndex
        
        If lCount = 0 Then
            strReturn = ""
        ElseIf bAllMatch Then
            strReturn = strText
        Else
            Select Case LotColumn.AggregateMode
                Case 0, 1:  ' Total, Average
                    strReturn = "Mixed"
            
                Case 2:     ' First
                    strReturn = strFirst
                    
                Case 3:     ' Last
                    strReturn = strLast
            
            End Select
        End If
    
    ElseIf UCase(LotColumn.Format) = "BOOLEAN" Then
        bAllChecked = True
        bAllUnchecked = True
        strFirst = ""
        strLast = ""
        
        For lIndex = 1 To Details.Count
            Set Detail = Details(lIndex)
            
            If Detail("LotColumnID") = Str(LotColumn.ID) Then
                strLast = Detail("Value")
                If Len(strFirst) = 0 Then
                    strFirst = strLast
                End If
                
                If strLast = "1" Then
                    bAllUnchecked = False
                Else
                    bAllChecked = False
                End If
            End If
        Next lIndex
        
        Select Case LotColumn.AggregateMode
            Case 0:     ' Total
                If bAllChecked Then
                    strReturn = "1"
                Else
                    strReturn = "0"
                End If
                
            Case 1:     ' Average
                If bAllUnchecked Then
                    strReturn = "1"
                Else
                    strReturn = "0"
                End If
        
            Case 2:     ' First
                strReturn = strFirst
            
            Case 3:     ' Last
                strReturn = strLast
        
        End Select
        
    Else
        dTotal = 0
        lCount = 0
        dFirst = kNullData
        
        For lIndex = 1 To Details.Count
            Set Detail = Details(lIndex)
            
            If Detail("LotColumnID") = Str(LotColumn.ID) Then
                lCount = lCount + 1
                dLast = Val(Detail("Value"))
                dTotal = dTotal + dLast
                
                If dFirst = kNullData Then
                    dFirst = dLast
                End If
            End If
        Next lIndex
        
        If lCount = 0 Then
            strReturn = ""
        Else
            Select Case LotColumn.AggregateMode
                Case 0:     ' Total
                    strReturn = Str(dTotal)
                    
                Case 1:     ' Average
                    strReturn = Str(dTotal / lCount)
                    
                Case 2:     ' First
                    strReturn = Str(dFirst)
                    
                Case 3:     ' Last
                    strReturn = Str(dLast)
                    
            End Select
        End If
    End If
    
    CalcTotalForDetails = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTurkey.CalcTotalForDetails"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IngredientList
'' Description: Build a list of ingredients from the collection
'' Inputs:      None
'' Returns:     Array of Ingredients
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IngredientList() As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array to return
    Dim lIndex As Long                  ' Index into a for loop
    Dim Ingredient As cBrokerMessage    ' Ingredient
    
    Set astrReturn = New cGdArray
    For lIndex = 1 To m.Ingredients.Count
        Set Ingredient = m.Ingredients(lIndex)
        astrReturn.Add Ingredient("Ingredient")
    Next lIndex
    
    astrReturn.Sort
    
    Set IngredientList = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattle.IngredientList"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TurnkeyMessageReceived
'' Description: Handle a Turnkey message from app-mail
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TurnkeyMessageReceived(ByVal nCattleMessageType As eGDCattleMessage, ByVal strTurnkeyMessage As String)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the message
    Dim turnkeyMessage As cBrokerMessage ' Turnkey message
    Dim yardCustomer As cBrokerMessage  ' Feedyard customer
    Dim strEndFlag As String            ' Flag on the end call
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim Account As cAccount             ' Account object

    If nCattleMessageType <> eGDCattleMessage_Heartbeat Then
        DumpDebug "Message Received from Turnkey (" & g.CattleEnums.CattleMessageTypeToString(nCattleMessageType) & "): '" & strTurnkeyMessage & "'"
    End If

    strFirstField = Parse(strTurnkeyMessage, vbTab, 1)
    Select Case nCattleMessageType
        Case eGDCattleMessage_AppLoaded
            SendConnect
        
        Case eGDCattleMessage_AppUnloaded
            If m.bReconnect = True Then
                StartStandAlone
            Else
                m.bReconnect = True
            End If
        
        Case eGDCattleMessage_ConnectionStatus
            Set turnkeyMessage = New cBrokerMessage
            turnkeyMessage.FromString strTurnkeyMessage
            
            If Len(turnkeyMessage("Error")) > 0 Then
                m.bReconnect = False
                g.AppBridge.AddDialogMessage turnkeyMessage("Error"), "!", , "Turnkey Error"
            End If
            
            Select Case UCase(turnkeyMessage("Status"))
                Case "DISCONNECTED"
                    ConnectionStatus = eGDConnectionStatus_Disconnected
                Case "DISCONNECTING"
                    ConnectionStatus = eGDConnectionStatus_Disconnecting
                Case "CONNECTING"
                    ConnectionStatus = eGDConnectionStatus_Connecting
                Case "CONNECTED"
                    ConnectionStatus = eGDConnectionStatus_Connected
            End Select
            
            If FormIsLoaded("frmLots") Then
                frmLots.Status = ConnectionStatus
            End If
            If FormIsLoaded("frmCattleAdmin") Then
                frmCattleAdmin.Status = ConnectionStatus
            End If
        
        Case eGDCattleMessage_FeedYard
            If UCase(strFirstField) = "BEGIN" Then
                m.FeedYards.Clear
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strTurnkeyMessage
                m.FeedYards.Add turnkeyMessage, turnkeyMessage("ID")
                
                If m.WaitAddCustomer.Exists(turnkeyMessage("RequestID")) = True Then
                    Set yardCustomer = m.WaitAddCustomer(turnkeyMessage("RequestID"))
                    
                    yardCustomer.Add "RequestID", NextRequestID
                    yardCustomer.Add "FeedYardID", turnkeyMessage("ID")
                    
                    g.Cattle.AddCustomers yardCustomer.ToString
                    
                    m.WaitAddCustomer.Remove turnkeyMessage("RequestID")
                End If
            End If
            
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_FeedYard strTurnkeyMessage
            End If
            If FormIsLoaded("frmCattleManage") Then
                frmCattleManage.Turnkey_FeedYard strTurnkeyMessage
            End If
        
        Case eGDCattleMessage_Customer
            If UCase(strFirstField) = "BEGIN" Then
                m.VisibleCustomers.Clear
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strTurnkeyMessage
                m.VisibleCustomers.Add turnkeyMessage
            End If
            
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_Customer strTurnkeyMessage
            End If
            If FormIsLoaded("frmCattleManage") Then
                frmCattleManage.Turnkey_Customer strTurnkeyMessage
            End If
            
        Case eGDCattleMessage_FeedyardCustomer
            If UCase(strFirstField) = "BEGIN" Then
                m.FeedyardCustomers.Clear
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strTurnkeyMessage
                m.FeedyardCustomers.Add turnkeyMessage
            End If
            
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_FeedyardCustomer strTurnkeyMessage
            End If
        
        Case eGDCattleMessage_Lot
            If UCase(strFirstField) = "BEGIN" Then
                m.Lots.Clear
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strTurnkeyMessage
                m.Lots.Add turnkeyMessage, turnkeyMessage("FeedYardLotID")
            End If
            
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_Lot strTurnkeyMessage
            End If
            
        Case eGDCattleMessage_Account
            If UCase(strFirstField) = "BEGIN" Then
                m.Accounts.Clear
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set Account = g.CattleKeyValue.AccountFromServerMessage(strTurnkeyMessage)
                
                If Len(Account.FcmNumber) = 0 Then
                    Account.FcmNumber = Account.AccountNumber
                End If
                Account.HasFills = False
                
                m.Accounts.Add Account
            End If
            
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_Account strTurnkeyMessage
            End If
    
        Case eGDCattleMessage_Order
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_Order strTurnkeyMessage
            End If
        
        Case eGDCattleMessage_AssociatedFill
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_AssociatedFill strTurnkeyMessage
            End If
        
        Case eGDCattleMessage_Fill
            If UCase(strFirstField) = "BEGIN" Then
                m.FillQuantities.Clear
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strTurnkeyMessage
                
                m.FillQuantities.AddFromFill turnkeyMessage
                
                If m.Accounts.Exists(turnkeyMessage("BrokerAccountID")) = True Then
                    m.Accounts(turnkeyMessage("BrokerAccountID")).HasFills = True
                End If
            End If
            
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_Fill strTurnkeyMessage
            End If
        
        Case eGDCattleMessage_Position
            If UCase(strFirstField) = "BEGIN" Then
                m.Positions.Clear
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strTurnkeyMessage
                
                m.Positions.Add turnkeyMessage, PositionsKey(turnkeyMessage("FeedYardLotID"), turnkeyMessage("BrokerAccountID"), turnkeyMessage("Symbol"))
            End If
            
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_Position strTurnkeyMessage
            End If
            
        Case eGDCattleMessage_Trade
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_Trade strTurnkeyMessage
            End If
            
        Case eGDCattleMessage_DeleteTrade
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_DeleteTrade strTurnkeyMessage
            End If
    
        Case eGDCattleMessage_DeleteOrder
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_DeleteOrder strTurnkeyMessage
            End If
    
        Case eGDCattleMessage_DeleteAssociatedFill
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_DeleteAssociatedFill strTurnkeyMessage
            End If
    
        Case eGDCattleMessage_LotColumns
            If UCase(strFirstField) = "BEGIN" Then
                m.LotColumns.Clear
                m.LotColumnMap.Clear
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set LotColumn = New cLotColumn
                LotColumn.FromString strTurnkeyMessage
                
                If m.LotColumns.Exists(LotColumn.ID) Then
                    Set m.LotColumns(LotColumn.ID) = LotColumn
                Else
                    m.LotColumns.Add LotColumn, LotColumn.ID
                End If
                
                If m.LotColumnMap.Exists(LotColumn.KeyValueField) Then
                    m.LotColumnMap(LotColumn.KeyValueField) = LotColumn.ID
                Else
                    m.LotColumnMap.Add LotColumn.ID, LotColumn.KeyValueField
                End If
            End If
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_LotColumn strTurnkeyMessage
            End If
        
        Case eGDCattleMessage_VisibleLotColumns
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_VisibleLotColumn strTurnkeyMessage
            End If
            
        Case eGDCattleMessage_VisibleLots
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_VisibleLots strTurnkeyMessage
            End If
    
        Case eGDCattleMessage_LotColumnCategories
            If UCase(strFirstField) = "BEGIN" Then
                m.LotColumnCategories.Clear
            ElseIf UCase(strFirstField) = "END" Then
                GetLotColumnSubCategories
            Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strTurnkeyMessage
                
                m.LotColumnCategories.Add turnkeyMessage, turnkeyMessage("ID")
            End If
        
        Case eGDCattleMessage_LotColumnSubCategories
            If UCase(strFirstField) = "BEGIN" Then
                m.LotColumnSubCategories.Clear
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strTurnkeyMessage
                
                m.LotColumnSubCategories.Add turnkeyMessage, turnkeyMessage("ID")
            End If
        
        Case eGDCattleMessage_LotContentDetails
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_LotContentsDetail strTurnkeyMessage
            End If
        
        Case eGDCattleMessage_GenesisCustomerInfo
            Set turnkeyMessage = New cBrokerMessage
            turnkeyMessage.FromString strTurnkeyMessage
            
            m.bCanEditLots = (turnkeyMessage("CanEditLots") = "1")
            m.nType = CLng(Val(turnkeyMessage("Type")))
            
            If FormIsLoaded("frmLots") Then
                frmLots.cmdAddCustomer.Visible = m.bCanEditLots
                frmLots.cmdAddFeedYard.Visible = m.bCanEditLots
            End If
            
        Case eGDCattleMessage_DetailOptions
            If UCase(strFirstField) = "BEGIN" Then
                m.DetailOptions.Clear
            ElseIf UCase(strFirstField) = "END" Then
                strEndFlag = Parse(strTurnkeyMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetIngredients
                End If
            Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strTurnkeyMessage
                                
                If m.DetailOptions.Exists(turnkeyMessage("LotColumnID")) = True Then
                    m.DetailOptions(turnkeyMessage("LotColumnID")) = turnkeyMessage("DetailOptions")
                Else
                    m.DetailOptions.Add turnkeyMessage("DetailOptions"), turnkeyMessage("LotColumnID")
                End If
            End If
            
        Case eGDCattleMessage_Ration
            If UCase(strFirstField) = "BEGIN" Then
                m.Rations.Clear
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strTurnkeyMessage
                
                If m.Rations.Exists(turnkeyMessage("ID")) = True Then
                    Set m.Rations(turnkeyMessage("ID")) = turnkeyMessage
                Else
                    m.Rations.Add turnkeyMessage, turnkeyMessage("ID")
                End If
            End If
            
            If FormIsLoaded("frmCattleManage") Then
                frmCattleManage.Turnkey_Ration strTurnkeyMessage
            End If
            If FormIsLoaded("frmCattleEditor") Then
                frmCattleEditor.Cattle_Ration strTurnkeyMessage
            End If
        
        Case eGDCattleMessage_Ingredient
            If UCase(strFirstField) = "BEGIN" Then
                m.Ingredients.Clear
            ElseIf UCase(strFirstField) = "END" Then
                strEndFlag = Parse(strTurnkeyMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetRations
                End If
            Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strTurnkeyMessage
                
                If m.Ingredients.Exists(turnkeyMessage("ID")) = True Then
                    Set m.Ingredients(turnkeyMessage("ID")) = turnkeyMessage
                Else
                    m.Ingredients.Add turnkeyMessage, turnkeyMessage("ID")
                End If
            End If
            
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_Ingredient strTurnkeyMessage
            End If
            If FormIsLoaded("frmCattleManage") Then
                frmCattleManage.Turnkey_Ingredient strTurnkeyMessage
            End If
            If FormIsLoaded("frmCattleEditor") Then
                frmCattleEditor.Cattle_Ingredient strTurnkeyMessage
            End If
            
        Case eGDCattleMessage_DeletePosition
            If FormIsLoaded("frmLots") Then
                frmLots.Turnkey_DeletePosition strTurnkeyMessage
            End If
            
            
        Case eGDCattleMessage_GenesisCustomers
            If FormIsLoaded("frmCattleAdmin") Then
                frmCattleAdmin.Cattle_GenesisCustomer strTurnkeyMessage
            End If
        
        Case eGDCattleMessage_AllFeedYards
            If FormIsLoaded("frmCattleAdmin") Then
                frmCattleAdmin.Cattle_FeedYard strTurnkeyMessage
            End If
        
        Case eGDCattleMessage_AllCustomers
            If FormIsLoaded("frmCattleAdmin") Then
                frmCattleAdmin.Cattle_FeedyardCustomer strTurnkeyMessage
            End If
        
        Case eGDCattleMessage_VisibleFeedYards
            If FormIsLoaded("frmCattleAdmin") Then
                frmCattleAdmin.Cattle_VisibleFeedyard strTurnkeyMessage
            End If
        
        Case eGDCattleMessage_VisibleCustomers
            If FormIsLoaded("frmCattleAdmin") Then
                frmCattleAdmin.Cattle_VisibleFeedyardCustomer strTurnkeyMessage
            End If
            
        Case eGDCattleMessage_LotColumnsAdmin
            If FormIsLoaded("frmCattleAdmin") Then
                frmCattleAdmin.Cattle_LotColumn strTurnkeyMessage
            End If
            
        Case eGDCattleMessage_VisibleLotColumnsAdmin
            If FormIsLoaded("frmCattleAdmin") Then
                frmCattleAdmin.Cattle_VisibleLotColumn strTurnkeyMessage
            End If
            
        Case eGDCattleMessage_DefaultVisibleLotColumnsAdmin
            If FormIsLoaded("frmCattleAdmin") Then
                frmCattleAdmin.Cattle_DefaultVisibleLotColumns strTurnkeyMessage
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.TurnkeyMessageReceived"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendTurnkeyMessage
'' Description: Send the Turnkey message through app-mail
'' Inputs:      Message Type, Message, Send Now?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SendTurnkeyMessage(ByVal nTurnkeyMessageType As eGDCattleMessage, ByVal turnkeyMessage As cBrokerMessage, Optional ByVal bSendNow As Boolean = True)
On Error GoTo ErrSection:

    SendTurnkeyMessageString nTurnkeyMessageType, turnkeyMessage.ToString(False), , turnkeyMessage.ToString(True)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.SendTurnkeyMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartStandAlone
'' Description: Start the stand alone application for Turnkey
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub StartStandAlone()
On Error GoTo ErrSection:

    Dim bConnect As Boolean             ' Should we still try to connect?
    Dim strProcess As String            ' Name of the process to start
    Dim lTimeOut As Long                ' Timeout variable
    Dim strArgs As String               ' Arguments to send to the standalone
    
    If (g.AppBridge.Unloading = False) Then
        ConnectionStatus = eGDConnectionStatus_Connecting
        
        DumpDebug "Trade Navigator version: " & g.AppBridge.FormatVersion(True, True)
        DumpDebug "NavCattle version: " & DllVersion(True, True)
        If Len(m.strProcessName) > 0 Then
            If ((frmCattleAM.gdCattle.FindControlID(m.strControlID) = 0) Or (KillProcess(m.strProcessName, True) = 0)) Then
                If KillProcess(m.strProcessName) > 0 Then
                    Sleep 3
                End If
                
                DumpDebug "Loading application '" & m.strProcessPath & "'"
                RunProcess m.strProcessPath, strArgs, , vbHide
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.StartStandAlone"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendConnect
'' Description: Send a connect statement to the stand alone
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SendConnect()
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim astrIpFile As cGdArray          ' Array of information from IP file
    
    Set astrIpFile = New cGdArray
    If m.strMode = "L" Then
        DumpDebug "Connecting to Live Server"
        astrIpFile.FromFile AddSlash(g.strAppPath) & "Provided\Turnkey.IP"
    ElseIf m.strMode = "D" Then
        DumpDebug "Connecting to Demo Server"
        astrIpFile.FromFile AddSlash(g.strAppPath) & "Provided\TurnkeyDemo.IP"
    ElseIf m.strMode = "T" Then
        DumpDebug "Connecting to Test Server"
        astrIpFile.FromFile AddSlash(g.strAppPath) & "Provided\TurnkeyTest.IP"
    End If
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "GenesisID", Str(GenesisID)
    turnkeyMessage.Add "Password", g.AppBridge.GetPassword
    turnkeyMessage.Add "TnVersion", g.AppBridge.FormatVersion(True, False)
    turnkeyMessage.Add "MachineID", g.AppBridge.GetMachineID
    turnkeyMessage.Add "IpInfo", astrIpFile(0)
    
    SendTurnkeyMessage eGDCattleMessage_Connect, turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.SendConnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendTurnkeyMessageString
'' Description: Send the Turnkey message through app-mail
'' Inputs:      Message Type, Message, Send Now?, Display message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SendTurnkeyMessageString(ByVal nCattleMessageType As eGDCattleMessage, ByVal strTurnkeyMessage As String, Optional ByVal bSendNow As Boolean = True, Optional ByVal strDisplayMessage As String = "")
On Error GoTo ErrSection:

    If Len(strDisplayMessage) > 0 Then
        DumpDebug "Sending Message to Turnkey (" & g.CattleEnums.CattleMessageTypeToString(nCattleMessageType) & "): '" & strDisplayMessage & "'"
    Else
        DumpDebug "Sending Message to Turnkey (" & g.CattleEnums.CattleMessageTypeToString(nCattleMessageType) & "): '" & strTurnkeyMessage & "'"
    End If
    
    frmCattleAM.gdCattle.CreateMessage m.strControlID, nCattleMessageType, strTurnkeyMessage, , bSendNow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.SendTurnkeyMessageString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.FeedYards = New cGdTree
    Set m.VisibleCustomers = New cGdTree
    Set m.FeedyardCustomers = New cGdTree
    Set m.Lots = New cGdTree
    Set m.Accounts = New cAccounts
    Set m.Positions = New cGdTree
    Set m.LotColumns = New cGdTree
    Set m.LotColumnMap = New cGdTree
    Set m.LotColumnCategories = New cGdTree
    Set m.LotColumnSubCategories = New cGdTree
    Set m.DetailOptions = New cGdTree
    Set m.Ingredients = New cGdTree
    Set m.Rations = New cGdTree
    Set m.FillQuantities = New cFillQuantities
    
    Set m.WaitAddCustomer = New cGdTree
    
    m.nStatus = eGDConnectionStatus_Disconnected
    m.strProcessName = "GenTurnkey"
    m.strProcessPath = AddSlash(g.strAppPath) & "..\Brokers\Turnkey\GenTurnkey.EXE"
    m.strControlID = "Turnkey"
    m.bReconnect = True
    m.strMode = ""
    m.lRequestCounter = GetIniFileProperty("Counter", 1&, "Turnkey", g.strIniFile)
    
    If Not DirExist(AddSlash(g.strAppPath) & "Turnkey") Then
        MkDir AddSlash(g.strAppPath) & "Turnkey"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.FeedYards = Nothing
    Set m.VisibleCustomers = Nothing
    Set m.FeedyardCustomers = Nothing
    Set m.Lots = Nothing
    Set m.Accounts = Nothing
    Set m.Positions = Nothing
    Set m.LotColumns = Nothing
    Set m.LotColumnMap = Nothing
    Set m.LotColumnCategories = Nothing
    Set m.LotColumnSubCategories = Nothing
    Set m.DetailOptions = Nothing
    Set m.Ingredients = Nothing
    Set m.Rations = Nothing
    Set m.FillQuantities = Nothing
    
    Set m.WaitAddCustomer = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattle.Class_Terminate"
    
End Sub
