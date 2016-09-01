Attribute VB_Name = "mOptNav"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mOptNav.bas
'' Description: Global routines for communicating with Option Navigator
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/11/2009   DAJ         Don't send streaming replay accounts to Opt Nav
'' 04/13/2009   DAJ         Added support for price threshold and secondary min moves
'' 04/20/2009   DAJ         Added support for option chain structure calls
'' 04/30/2009   DAJ         Coded the OptionChainBuilt function
'' 05/01/2009   DAJ         Added new messages to OptNavMessageTypeString function
'' 05/19/2009   DAJ         Added enumeration for Option Navigator status
'' 06/05/2009   DAJ         Added ConnectToAccount call from Option Navigator
'' 06/22/2009   DAJ         Added TicketSubmitted message
'' 09/23/2009   DAJ         Moved StartOptNav in here, added CFG information
'' 01/29/2010   DAJ         Added call for option nav to add symbols to quote board
'' 09/15/2010   DAJ         Added Rithmic to the broker list for Option Navigator
'' 11/01/2010   DAJ         Added Optimus, OpVest, and Vision (Rithmic Brokers)
'' 11/10/2010   DAJ         Added ParkOrder, SubmitOrder, ChangeGroupInfo calls
'' 11/18/2010   DAJ         Added Starting Genesis ID for Option Navigator
'' 11/18/2010   DAJ         Send Order callback to ON upon a Group Name change
'' 12/10/2010   DAJ         Added Zen-Fire, Changed over to the IsBrokerUser function
'' 01/26/2011   DAJ         Fixes for parking/resubmitting orders from Option Nav
'' 03/07/2011   DAJ         Added account filtering, historical fills
'' 03/09/2011   DAJ         Only send positions that have non-zero carried or non-zero current
'' 03/09/2011   DAJ         Added ExpirePosition call
'' 03/10/2011   DAJ         OptNav is passing access ID for calls on previously existing orders
'' 03/16/2011   DAJ         Add Flatten and CancelAll calls
'' 03/31/2011   DAJ         Don't send snapshot fills in the historical fills file (#6233)
'' 04/05/2011   DAJ         Moved broker specific SymbolInformation to cBrokerDispatch
'' 04/13/2011   DAJ         Set private member defaults when ON first loading
'' 05/02/2011   DAJ         Added Interactive Brokers/Ideal as an options broker
'' 05/09/2011   DAJ         Send IB/Ideal accounts over to Option Nav as well
'' 05/09/2011   DAJ         Send option bar properties on symbol info call if underlying asked for
'' 06/21/2011   DAJ         Separate out Simulated trading types, Interactive Brokers Temp ID
'' 07/14/2011   DAJ         Log if order operation not performed because order not found
'' 07/22/2011   DAJ         Reassign Genesis ID if submitting a parked IB order
'' 08/09/2011   DAJ         Don't send fills to option nav with price of 0
'' 08/15/2011   DAJ         Option Nav now needs the fills with a price of 0
'' 08/23/2011   DAJ         Only send account message if it is Option Nav's current account
'' 10/27/2011   DAJ         Don't apply the account filter when sending an account message
'' 10/31/2011   DAJ         Send Option Nav a message when account deleted
'' 11/18/2011   DAJ         Utilize IsBrokerUserOptions function, Send CQG brokers to Option Nav
'' 11/18/2011   DAJ         Don't send account to Opt Nav if not enabled
'' 12/06/2011   DAJ         Added expiration date calls
'' 01/06/2012   DAJ         Allow for no arguments to GetAccounts call to return all accounts
'' 01/09/2012   DAJ         In RiskGraphBuilt, exit for loop after we found our chart
'' 02/14/2012   DAJ         Added multi-leg order support
'' 04/09/2012   DAJ         Fixed account for CancelOptionNavigatorOrder call
'' 04/16/2012   DAJ         Added the RenameOptionNavGroup call
'' 05/15/2012   DAJ         Ability to re-park a parked order, fixes for RenameOptionNavGroup
'' 10/02/2012   DAJ         Made the BuildBrokerList routine more generic
'' 02/22/2013   DAJ         Send the correct number of max order legs in broker list
'' 05/22/2013   DAJ         Don't send message to OptionNav unless it is loaded
'' 10/16/2013   DAJ         Removed PFG and Xpress, Added Oec/FptOec/FptCqg
'' 02/24/2014   DAJ         Pass allowable security types on broker message
'' 04/23/2014   DAJ         Pass Robbins CQG accounts over to Option Nav
'' 05/30/2014   DAJ         Fixed BuildAccountList function to look through all brokers
'' 09/04/2014   DAJ         Pulled Option Navigator conversions out of trade objects
'' 12/09/2015   DAJ         Added the NextGroupID function and pass the value in the OptionNav.CFG file
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Declare Function SalmonAppMail Lib "SalmonClient.dll" (ByVal nMsgType As Long, ByVal hStringArray As Long) As Long
Private Declare Function RemoveEODForOptionsNavigator Lib "SalmonClient.dll" (ByVal hStringArray As Long) As Long

Private mLastMsgReceived As Double

Public Enum eGDOptNavStatus
    eGDOptNavStatus_Unloaded = 0
    eGDOptNavStatus_Loading
    eGDOptNavStatus_Loaded
End Enum

Public Enum eGDOptNavMessageType
    eGDOptNav_UnknownMessage = 0        ' Unknown message error

    eGDOptNav_GetBrokers                ' Option Navigator request for broker list
    eGDOptNav_Broker                    ' Broker returned to Option Navigator
    
    eGDOptNav_GetAccounts               ' Option Navigator request for account list
    eGDOptNav_Account                   ' Account returned to Option Navigator
    
    eGDOptNav_GetOrders                 ' Option Navigator request for order list
    eGDOptNav_Order                     ' Order returned to Option Navigator
    
    eGDOptNav_GetFills                  ' Option Navigator request for fill list
    eGDOptNav_Fill                      ' Fill returned to Option Navigator
    
    eGDOptNav_GetPositions              ' Option Navigator request for position list
    eGDOptNav_Position                  ' Position returned to Option Navigator

    eGDOptNav_SymbolRequest             ' Option Navigator request for symbol selection
    eGDOptNav_Symbols                   ' Selected symbol returned to Option Navigator
    
    eGDOptNav_InfoRequest               ' Option Navigator request for customer information
    eGDOptNav_Information               ' Customer information returned to Option Navigator
    
    eGDOptNav_OptNavLoaded              ' Option Navigator is loaded
    eGDOptNav_Activate                  ' Request Option Navigator to grab the focus
    
    eGDOptNav_OptNavUnloaded            ' Option Navigator is unloaded
    eGDOptNav_Unload                    ' Request Option Navigator to unload
    
    eGDOptNav_AddOrder                  ' Option Navigator is adding a new order
    
    eGDOptNav_AmendOrder = 21           ' Option Navigator is amending an existing order
    
    eGDOptNav_CancelOrder = 23          ' Option Navigator is cancelling an existing order
    
    eGDOptNav_GetSymbolInfo = 25        ' Option Navigator request for symbol information
    eGDOptNav_SymbolInfo                ' Symbol information returned to Option Navigator
    
    eGDOptNav_ChainBuilt = 27           ' Option chain requested has been built
    eGDOptNav_RequestChain              ' Request option chain structure from Option Navigator
    
    eGDOptNav_CreateTicket = 30         ' Tell Option Navigator to create an order ticket
    eGDOptNav_TicketSubmitted           ' Option ticket that was created has been submitted
    
    eGDOptNav_ConnectToAccount = 33     ' Attempt to connect to specified broker/account
    
    eGDOptNav_RiskGraphBuilt = 35       ' Risk graph has been built
    eGDOptNav_RequestRiskGraph          ' Request risk graph from Option Navigator
    
    eGDOptNav_RequestPriceData = 37     ' OptionNav request for price data
    eGDOptNav_PriceData                 ' Price data being sent to OptionNav
    
    eGDOptNav_RequestGroupSymbols = 39  ' OptionNav request for list of symbols in a group
    eGDOptNav_GroupSymbols              ' List of symbols being sent to OptionNav
    
    eGDOptNav_AddToQuoteBoard = 41      ' OptionNav request for adding symbols to the quote board
    
    eGDOptNav_ParkOrder = 43            ' OptionNav request to Park an order
    
    eGDOptNav_SubmitOrder = 45          ' OptionNav request to submit a parked order
    
    eGDOptNav_ChangeGroupInfo = 47      ' Change the group information for the given order
    
    eGDOptNav_CurrentAccount = 49       ' Current account set in Option Navigator
    
    eGDOptNav_GetHistoricalFills = 51   ' OptionNav request for historical fills
    eGDOptNav_HistoricalFills           ' Historical fills returned to OptionNav
    
    eGDOptNav_ClearGroup = 54           ' Tell Option Nav to clear out a particular group
    
    eGDOptNav_OrderIdChanged = 56       ' Tell Option Nav that order ID is changing
    
    eGDOptNav_ExpirePosition            ' Request from Option Navigator to expire a position
    
    eGDOptNav_FlattenPosition = 59      ' Request from Option Navigator to flatten a position
    
    eGDOptNav_CancelAllOrder = 61       ' Request from Option Navigator to cancel all orders for a symbol
    
    eGDOptNav_ReversePosition = 63      ' Request from Option Navigator to reverse a position
    
    eGDOptNav_ExpirationDate = 65       ' Option Navigator statement giving us an expiration date
    eGDOptNav_GetExpirationDate         ' Request from Trade Navigator for an expiration date
    
    eGDOptNav_AddSpreadOrder            ' Add a spread order
    eGDOptNav_SpreadOrder               ' Spread order returned to Option Nav
    
    eGDOptNav_AmendSpreadOrder          ' Amend a spread order
    
    eGDOptNav_ParkSpreadOrder = 71      ' Park a spread order
    
    eGDOptNav_SubmitSpreadOrder = 73    ' Submit a parked spread order
    
    eGDOptNav_RenameGroup = 75          ' Rename a group with the given ID
    
    eGDOptNav_SalmonClientMessages = 1000 ' everything over 1000 goes directly to the salmon client
    eGDOptNav_OptionChainRequest = 1500 ' or Portfolio (1501 reserved for response back)
End Enum

Private Type mPrivate
    strCurrentAccount As String         ' Currently selected account in Option Navigator
    nCurrentBroker As eTT_AccountType   ' Broker for the currently selected account in Option Navigator
End Type
Private m As mPrivate

Public Function OptNavMsg(ByVal nMsgType As eGDOptNavMessageType) As Long
    OptNavMsg = nMsgType
End Function

Public Property Get OptNavIP() As String
    OptNavIP = Trim(FileToString(App.Path & "\Provided\OptionNav.IP", , True))
End Property

Public Property Get OptNavExeFile() As String
    OptNavExeFile = App.Path & "\..\OptionNav\OptionNav.exe"
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendMessageToOptNav
'' Description: Send an App Mail message to Option Navigator
'' Inputs:      Message Type, Message, Send Now?, Number of Tries
'' Returns:     Return value from CreateMessage
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SendMessageToOptNav(ByVal nMsgType As eGDOptNavMessageType, ByVal strMessage As String, Optional ByVal bSendNow As Boolean = True) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim strOptNav As String             ' Option Navigator app mail client name
    Dim lTryNumber As Long              ' Try number
    Dim lNumTries As Long

    lReturn = 0&
    
    If g.nOptNavStatus >= eGDOptNavStatus_Loading Then
        strOptNav = frmOnlineBroker.apmOptNav.Tag
        If Len(strOptNav) = 0 Then
            strOptNav = "OptNavClient"
        End If
        
        ' TLB 10/21/2009: if we are responding to a message just received from OptNav,
        ' then we should allow for more retries (in case OptNav had just activated)
        If gdTickCount(True) < mLastMsgReceived + 1000 Then
            lNumTries = 10
        Else
            lNumTries = 1
        End If
        
        OptNavLog "Sending Message to OptNav(" & OptNavMessageTypeString(nMsgType) & "): " & strMessage
        
        lTryNumber = 0&
        Do
            lReturn = frmOnlineBroker.apmOptNav.CreateMessage(strOptNav, nMsgType, strMessage, , bSendNow)
            lTryNumber = lTryNumber + 1&
            
            If (lReturn = 0) Then
                If (lTryNumber < lNumTries) Then
                    Sleep 0.1
                End If
            End If
        Loop Until (lReturn <> 0) Or (lTryNumber >= lNumTries)
        
        If lTryNumber > 1& Then
            OptNavLog "-->Send Message Tried More than Once: Return=" & Str(lReturn) & "; Tries=" & Str(lTryNumber)
        End If
    End If
    
    SendMessageToOptNav = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.SendMessageToOptNav"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    apmOptNav_MessageReceived
'' Description: Handle a message received from Option Navigator
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OptNavMessageReceived(msg As gdOCX.gdAppMailMsg)
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message
    Dim nBroker As eTT_AccountType      ' Broker from the message
    Dim aStrings As cGdArray

    mLastMsgReceived = gdTickCount(True)

    OptNavLog "Message Received from OptNav(" & OptNavMessageTypeString(msg.MsgType) & "): " & msg.Message
    
    ' Set the flag saying that Option Navigator has sent us messages.  This will be used
    ' to tell that Option Navigator is up before getting the Option Navigator loaded message
    ' so that they can hold off sending that message until they are fully ready for
    ' other things to happen (DAJ 05/19/2009)...
    If g.nOptNavStatus = eGDOptNavStatus_Unloaded Then
        m.nCurrentBroker = -1&
        m.strCurrentAccount = ""
        g.nOptNavStatus = eGDOptNavStatus_Loading
    End If
        
    ' message types between 1000 and 1500 are for the salmon client
    If msg.MsgType >= 1000 And msg.MsgType < 1500 Then
        ' send message as first string in the array
        Set aStrings = New cGdArray
        aStrings.Create eGDARRAY_Strings, 1
        aStrings(0) = msg.Message
        SalmonAppMail msg.MsgType, aStrings.ArrayHandle
        Set aStrings = Nothing
    Else
        Select Case msg.MsgType
        Case OptNavMsg(eGDOptNav_GetBrokers)
            SendBrokerList
        
        Case OptNavMsg(eGDOptNav_GetAccounts)
            If Len(msg.Message) > 0 Then
                nBroker = CLng(Val(Parse(msg.Message, vbTab, 1)))
                SendAccountListForBroker nBroker
            Else
                SendAccountListForBroker -1&
            End If
        
        Case OptNavMsg(eGDOptNav_GetOrders)
            SendOrdersForAccount Parse(msg.Message, vbTab, 1)
        
        Case OptNavMsg(eGDOptNav_GetFills)
            SendFillsForAccount Parse(msg.Message, vbTab, 1)
        
        Case OptNavMsg(eGDOptNav_GetPositions)
            SendPositionsForAccount Parse(msg.Message, vbTab, 1)
        
        Case OptNavMsg(eGDOptNav_SymbolRequest)
            frmMain.tmrSymbol.Tag = msg.Message
            frmMain.tmrSymbol.Enabled = True
        
        Case OptNavMsg(eGDOptNav_InfoRequest)
            ' request for general info:  DataSrvID, Enablements, IP addresses
            strMessage = Trim(FileToString(App.Path & "\Provided\OptionNav.IP", , True))
            SendMessageToOptNav eGDOptNav_Information, Str(RI_GetDataServiceID) & vbTab _
                & Trim(UCase(g.strAuthorizationString)) & vbTab & strMessage & vbTab
        
        Case OptNavMsg(eGDOptNav_OptNavLoaded)
            g.nOptNavStatus = eGDOptNavStatus_Loaded
            
        Case OptNavMsg(eGDOptNav_OptNavUnloaded)
            g.nOptNavStatus = eGDOptNavStatus_Unloaded
        
        Case OptNavMsg(eGDOptNav_AddOrder)
            AddOptionNavigatorOrder msg.Message, False
        
        Case OptNavMsg(eGDOptNav_AmendOrder)
            AmendOptionNavigatorOrder msg.Message, False
        
        Case OptNavMsg(eGDOptNav_CancelOrder)
            CancelOptionNavigatorOrder msg.Message, False
            
        Case OptNavMsg(eGDOptNav_GetSymbolInfo)
            GetSymbolInfoForOptNav msg.Message
            
        Case OptNavMsg(eGDOptNav_ChainBuilt)
            OptionChainBuilt msg.Message
            
        Case OptNavMsg(eGDOptNav_TicketSubmitted)
            TicketSubmitted msg.Message
            
        Case OptNavMsg(eGDOptNav_ConnectToAccount)
            ConnectToAccount msg.Message
            
        Case OptNavMsg(eGDOptNav_RiskGraphBuilt)
            RiskGraphBuilt msg.Message
            
        Case OptNavMsg(eGDOptNav_RequestPriceData)
            GetPriceDataForOptNav msg.Message
            
        Case OptNavMsg(eGDOptNav_RequestGroupSymbols)
            strMessage = g.SymbolPool.GetSymbolsForGroup(msg.Message)
            SendMessageToOptNav eGDOptNav_GroupSymbols, strMessage, True
            
        Case OptNavMsg(eGDOptNav_AddToQuoteBoard)
            AddSymbolsToQuoteBoard msg.Message
                   
        Case OptNavMsg(eGDOptNav_OptionChainRequest)
            'If IsIDE Then
            '    frmTest.AddList "OptionChain Request = " & Left(msg.Message, 100)
            'End If
            SyncOptNavSymbolsWithSalmon msg.Message
            SendMessageToOptNav msg.MsgType + 1, "SymbolAddProcessed", True
               
        Case OptNavMsg(eGDOptNav_ParkOrder)
            ParkOptionNavigatorOrder msg.Message, False
        
        Case OptNavMsg(eGDOptNav_SubmitOrder)
            SubmitOptionNavigatorOrder msg.Message, False
            
        Case OptNavMsg(eGDOptNav_ChangeGroupInfo)
            ChangeGroupInfoForOrder msg.Message
            
        Case OptNavMsg(eGDOptNav_CurrentAccount)
            m.strCurrentAccount = msg.Message
            If Len(m.strCurrentAccount) > 0 Then
                m.nCurrentBroker = g.Broker.AccountTypeForNumber(m.strCurrentAccount)
            Else
                m.nCurrentBroker = -1&
            End If
        
        Case OptNavMsg(eGDOptNav_GetHistoricalFills)
            GetHistoricalFillsForOptionNav
            
        Case OptNavMsg(eGDOptNav_ExpirePosition)
            HandleExpirePosition msg.Message
            
        Case OptNavMsg(eGDOptNav_FlattenPosition)
            HandleFlattenFromOptNav msg.Message
        
        Case OptNavMsg(eGDOptNav_CancelAllOrder)
            HandleCancelAllFromOptNav msg.Message
            
        Case OptNavMsg(eGDOptNav_ReversePosition)
            HandleReverseFromOptNav msg.Message
            
        Case OptNavMsg(eGDOptNav_ExpirationDate)
            HandleExpirationDateFromOptNav msg.Message
        
        Case OptNavMsg(eGDOptNav_AddSpreadOrder)
            AddOptionNavigatorOrder msg.Message, True
        
        Case OptNavMsg(eGDOptNav_AmendSpreadOrder)
            AmendOptionNavigatorOrder msg.Message, True
        
        Case OptNavMsg(eGDOptNav_ParkSpreadOrder)
            ParkOptionNavigatorOrder msg.Message, True
        
        Case OptNavMsg(eGDOptNav_SubmitSpreadOrder)
            SubmitOptionNavigatorOrder msg.Message, True
            
        Case OptNavMsg(eGDOptNav_RenameGroup)
            RenameOptionNavGroup msg.Message
        
        Case Else
            ' just send an unhandled message back
            SendMessageToOptNav eGDOptNav_UnknownMessage, Str(msg.MsgType) & " is an unrecognized message type."
    
        End Select
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.OptNavMessageReceived"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendBrokerList
'' Description: Send the broker list to Option Navigator
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendBrokerList()
On Error GoTo ErrSection:

    Dim astrBrokerList As cGdArray      ' Broker list
    Dim lIndex As Long                  ' Index into a for loop

    Set astrBrokerList = BuildBrokerList
    
    SendMessageToOptNav eGDOptNav_Broker, "BEGIN", False
    
    If Not astrBrokerList Is Nothing Then
        For lIndex = 0 To astrBrokerList.Size - 1
            SendMessageToOptNav eGDOptNav_Broker, astrBrokerList(lIndex), False
        Next lIndex
    End If
    
    SendMessageToOptNav eGDOptNav_Broker, "END", True
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendBrokerList"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendAccountListForBroker
'' Description: Send the accout list for a broker to Option Navigator
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendAccountListForBroker(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim astrAccountList As cGdArray     ' Account list for the given broker
    Dim lIndex As Long                  ' Index into a for loop
        
    SendMessageToOptNav eGDOptNav_Account, "BEGIN", False
        
    If nBroker > -1& Then
        Set astrAccountList = BuildAccountListForBroker(nBroker)
        If Not astrAccountList Is Nothing Then
            For lIndex = 0 To astrAccountList.Size - 1
                SendMessageToOptNav eGDOptNav_Account, astrAccountList(lIndex), False
            Next lIndex
        End If
    Else
        Set astrAccountList = BuildAccountList
        If Not astrAccountList Is Nothing Then
            For lIndex = 0 To astrAccountList.Size - 1
                SendMessageToOptNav eGDOptNav_Account, astrAccountList(lIndex), False
            Next lIndex
        End If
    End If

    SendMessageToOptNav eGDOptNav_Account, "END", True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendAccountListForBroker"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendOrdersForAccount
'' Description: Send the orders for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendOrdersForAccount(ByVal strAccountNumber As String)
On Error GoTo ErrSection:

    Dim Orders As cPtOrders             ' Orders for the given account
    Dim lIndex As Long                  ' Index into a for loop
    Dim lAccountID As Long              ' Account ID for the number passed in
    
    SendMessageToOptNav eGDOptNav_Order, "BEGIN", False
    
    Set Orders = g.Broker.OrdersForAccount(strAccountNumber)
    If Not Orders Is Nothing Then
        lAccountID = g.Broker.AccountIDForNumber(strAccountNumber)
        
        For lIndex = 1 To Orders.Count
            If Orders(lIndex).AccountID = lAccountID Then
                SendOrderToOptionNav Orders(lIndex), True
                'SendMessageToOptNav eGDOptNav_Order, Orders(lIndex).OptionNavigatorString(True), False
            End If
        Next lIndex
    End If
    
    SendMessageToOptNav eGDOptNav_Order, "END", True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendOrdersForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendFillsForAccount
'' Description: Send the fills for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendFillsForAccount(ByVal strAccountNumber As String)
On Error GoTo ErrSection:

    Dim Fills As cPtFills               ' Fills for the given account
    Dim lIndex As Long                  ' Index into a for loop
    Dim lAccountID As Long              ' Account ID for the number passed in
    
    SendMessageToOptNav eGDOptNav_Fill, "BEGIN", False
    
    Set Fills = g.Broker.FillsForAccount(strAccountNumber)
    If Not Fills Is Nothing Then
        lAccountID = g.Broker.AccountIDForNumber(strAccountNumber)
        
        For lIndex = 1 To Fills.Count
            If Fills(lIndex).AccountID = lAccountID Then
                'If Fills(lIndex).Price <> 0# Then
                    SendMessageToOptNav eGDOptNav_Fill, FillToString(Fills(lIndex), True), False
                'End If
            End If
        Next lIndex
    End If
    
    SendMessageToOptNav eGDOptNav_Fill, "END", True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendFillsForAccount"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendPositionsForAccount
'' Description: Send the positions for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendPositionsForAccount(ByVal strAccountNumber As String)
On Error GoTo ErrSection:

    Dim Positions As cAccountPositions  ' Account positions
    Dim lIndex As Long                  ' Index into a for loop
    Dim lAccountID As Long              ' Account ID for the number passed in
    
    SendMessageToOptNav eGDOptNav_Position, "BEGIN", False
    
    Set Positions = g.Broker.FillSummariesForAccount(strAccountNumber)
    If Not Positions Is Nothing Then
        lAccountID = g.Broker.AccountIDForNumber(strAccountNumber)
        
        For lIndex = 1 To Positions.Count
            If (Positions(lIndex).AccountID = lAccountID) And (Positions(lIndex).AutoTradeItemID = -1) Then
                If (Positions(lIndex).CurrentPosition <> 0) Or (Positions(lIndex).CurrentPositionSnapshot <> 0) Then
                    SendMessageToOptNav eGDOptNav_Position, AccountPositionToString(Positions(lIndex), True), False
                End If
            End If
        Next lIndex
    End If

    SendMessageToOptNav eGDOptNav_Position, "END", True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendPositionsForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendNowToOptionNav
'' Description: Perform a SendNow to send all queued Option Nav messages
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendNowToOptionNav()
On Error GoTo ErrSection:

    If g.nOptNavStatus >= eGDOptNavStatus_Loading Then
        frmOnlineBroker.apmOptNav.SendNow
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendNowToOptionNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendBeginToOptionNav
'' Description: Send a "Begin" to Option Navigator if applicable
'' Inputs:      Message Type, Broker
'' Returns:     True if Sent, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SendBeginToOptionNav(ByVal nType As eGDOptNavMessageType, ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If g.nOptNavStatus >= eGDOptNavStatus_Loading Then
        If IncludeBrokerToOptionNav(nBroker) Then
            SendMessageToOptNav eGDOptNav_ClearGroup, Str(nType) & vbTab & Str(nBroker)
            SendMessageToOptNav nType, "BEGIN"
            bReturn = True
        End If
    End If
    
    SendBeginToOptionNav = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.SendBeginToOptionNav"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendEndToOptionNav
'' Description: Send a "End" to Option Navigator if applicable
'' Inputs:      Message Type, Broker, Send Now?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendEndToOptionNav(ByVal nType As eGDOptNavMessageType, ByVal nBroker As eTT_AccountType, Optional ByVal bSendNow As Boolean = False)
On Error GoTo ErrSection:

    If g.nOptNavStatus >= eGDOptNavStatus_Loading Then
        If IncludeBrokerToOptionNav(nBroker) Then
            SendMessageToOptNav nType, "END"
            
            If bSendNow Then
                SendNowToOptionNav
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendEndToOptionNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendAccountToOptionNav
'' Description: Send the given account to Option Navigator if applicable
'' Inputs:      Account, Refresh?, Connection Status, Removed?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendAccountToOptionNav(ByVal Account As cPtAccount, ByVal bRefresh As Boolean, Optional ByVal nConnectionStatus As eGDConnectionStatus = -1&, Optional ByVal bRemoved As Boolean = False)
On Error GoTo ErrSection:

    If g.nOptNavStatus >= eGDOptNavStatus_Loading Then
        If g.Broker.IsBrokerUserOptions(Account.AccountType) Then
            SendMessageToOptNav eGDOptNav_Account, AccountToString(Account, bRefresh, nConnectionStatus, bRemoved), False
            If Not bRefresh Then
                SendNowToOptionNav
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendAccountToOptionNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendOrderToOptionNav
'' Description: Send the given order to Option Navigator if applicable
'' Inputs:      Order, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendOrderToOptionNav(ByVal Order As cPtOrder, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    If g.nOptNavStatus >= eGDOptNavStatus_Loading Then
        If IncludeAccountToOptionNav(g.Broker.AccountNumberForID(Order.AccountID)) Then
            If Order.NumberOfLegs = 1 Then
                SendMessageToOptNav eGDOptNav_Order, OrderToString(Order, bRefresh), False
            Else
                SendMessageToOptNav eGDOptNav_SpreadOrder, OrderToSpreadString(Order, bRefresh), False
            End If
        End If
    
        If Not bRefresh Then
            SendNowToOptionNav
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendOrderToOptionNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendFillToOptionNav
'' Description: Send the given fill to Option Navigator if applicable
'' Inputs:      Fill, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendFillToOptionNav(ByVal Fill As cPtFill, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    If g.nOptNavStatus >= eGDOptNavStatus_Loading Then
        If IncludeAccountToOptionNav(g.Broker.AccountNumberForID(Fill.AccountID)) Then
            'If Fill.Price <> 0# Then
                SendMessageToOptNav eGDOptNav_Fill, FillToString(Fill, bRefresh), False
            'End If
        End If
    
        If Not bRefresh Then
            SendNowToOptionNav
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendFillToOptionNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendPositionToOptionNav
'' Description: Send the given position to Option Navigator if applicable
'' Inputs:      Position, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendPositionToOptionNav(ByVal Position As cAccountPosition, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    If g.nOptNavStatus >= eGDOptNavStatus_Loading Then
        If IncludeAccountToOptionNav(g.Broker.AccountNumberForID(Position.AccountID)) Then
            SendMessageToOptNav eGDOptNav_Position, AccountPositionToString(Position, bRefresh), False
        End If
    
        If Not bRefresh Then
            SendNowToOptionNav
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendPositionToOptionNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendPositionStringToOptionNav
'' Description: Send the given position to Option Navigator if applicable
'' Inputs:      Position, Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendPositionStringToOptionNav(ByVal strPosition As String, ByVal vAccountNumberOrID As Variant, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    If g.nOptNavStatus >= eGDOptNavStatus_Loading Then
        If IncludeAccountToOptionNav(g.Broker.GetAccountNumber(vAccountNumberOrID)) Then
            SendMessageToOptNav eGDOptNav_Position, strPosition, False
        End If
    
        If Not bRefresh Then
            SendNowToOptionNav
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendPositionStringToOptionNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOptionNavigatorOrder
'' Description: Submit the order that was sent to us by Option Navigator
'' Inputs:      Order, Spread Order?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddOptionNavigatorOrder(ByVal strMessage As String, Optional ByVal bSpreadOrder As Boolean = False)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    
    Set Order = New cPtOrder
    If bSpreadOrder Then
        OrderFromSpreadString strMessage, Order
    Else
        OrderFromString strMessage, Order
    End If
    
    ' 06/15/2011 DAJ: Do a save up front here to get an AccessID for the order
    ' to send to Option Navigator via a OrderIdChange event...
    Order.Status = eTT_OrderStatus_Open
    Order.Save
    
    ' 06/15/2011 DAJ: If this is an Interactive Brokers order then we need to
    ' reassign the Genesis ID here...
    If (Order.Broker = eTT_AccountType_IntBrokers) Or (Order.Broker = eTT_AccountType_Ideal) Then
        Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Order.AccountID))
    End If
    
    g.Broker.BrokerDebug g.Broker.AccountTypeForID(Order.AccountID), "Submitting Order from Option Navigator: " & Order.OrderText, True
    SubmitOrder Order
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.AddOptionNavigatorOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AmendOptionNavigatorOrder
'' Description: Amend the order that was sent to us by Option Navigator
'' Inputs:      Order, Spread Order?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AmendOptionNavigatorOrder(ByVal strMessage As String, Optional ByVal bSpreadOrder As Boolean = False)
On Error GoTo ErrSection:

    Dim OldOrder As cPtOrder            ' Order object
    Dim NewOrder As cPtOrder            ' Order object
    Dim SpreadTree As cGdTree           ' Spread tree
    
    Set OldOrder = Nothing
    If bSpreadOrder Then
        Set SpreadTree = New cGdTree
        SpreadTree.FromKeyValueString strMessage
        If SpreadTree.Exists("AccountNumber") And SpreadTree.Exists("AccessID") Then
            Set OldOrder = g.Broker.OrderForAccessID(SpreadTree("AccountNumber"), CLng(Val(SpreadTree("AccessID"))))
            If Not OldOrder Is Nothing Then
                Set NewOrder = New cPtOrder
                Set NewOrder = OldOrder.MakeCopy
                OrderFromSpreadTree SpreadTree, NewOrder
            Else
                OptNavLog "AmendOrder not performed because order '" & SpreadTree("AccessID") & "' not found in account '" & SpreadTree("AccountNumber") & "'"
            End If
        Else
            OptNavLog "AmendOrder not performed because either account number or access ID doesn't exist in string"
        End If
    Else
        Set OldOrder = g.Broker.OrderForAccessID(Parse(strMessage, vbTab, 3), CLng(Val(Parse(strMessage, vbTab, 1))))
        If Not OldOrder Is Nothing Then
            Set NewOrder = New cPtOrder
            Set NewOrder = OldOrder.MakeCopy
            OrderFromString strMessage, NewOrder
        Else
            OptNavLog "AmendOrder not performed because order '" & Parse(strMessage, vbTab, 1) & "' not found in account '" & Parse(strMessage, vbTab, 3) & "'"
        End If
    End If
    
    If Not OldOrder Is Nothing Then
        g.Broker.BrokerDebug OldOrder.Broker, "Modifying Order from Option Navigator: " & OldOrder.OrderText, True
        ModifyOrder OldOrder, , , False, NewOrder
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.AmendOptionNavigatorOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOptionNavigatorOrder
'' Description: Cancel the order that was sent to us by Option Navigator
'' Inputs:      Order, Spread Order?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelOptionNavigatorOrder(ByVal strMessage As String, Optional ByVal bSpreadOrder As Boolean = False)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    Dim SpreadTree As cGdTree           ' Spread tree
    
    Set Order = Nothing
    If bSpreadOrder Then
        Set SpreadTree = New cGdTree
        SpreadTree.FromKeyValueString strMessage
        If SpreadTree.Exists("AccountNumber") And SpreadTree.Exists("AccessID") Then
            Set Order = g.Broker.OrderForAccessID(SpreadTree("AccountNumber"), CLng(Val(SpreadTree("AccessID"))))
            If Order Is Nothing Then
                OptNavLog "CancelOrder not performed because order '" & SpreadTree("AccessID") & "' not found in account '" & SpreadTree("AccountNumber") & "'"
            End If
        Else
            OptNavLog "CancelOrder not performed because either account number or access ID doesn't exist in string"
        End If
    Else
        Set Order = g.Broker.OrderForAccessID(Parse(strMessage, vbTab, 2), CLng(Val(Parse(strMessage, vbTab, 1))))
        If Order Is Nothing Then
            OptNavLog "CancelOrder not performed because order '" & Parse(strMessage, vbTab, 1) & "' not found in account '" & Parse(strMessage, vbTab, 2) & "'"
        End If
    End If
    
    If Not Order Is Nothing Then
        ' If the order is closed or pending, just send the order status back to Option
        ' Navigator for right now...
        If (IsOpenOrder(Order.Status, False) = False) Or (OrderIsPending(Order) = True) Then
            g.Broker.BrokerDebug Order.Broker, "Cancel Order Requested from Option Navigator, but order is closed: " & Order.OrderText, True
            'SendMessageToOptNav eGDOptNav_Order, Order.OptionNavigatorString(False)
            SendOrderToOptionNav Order, False
        Else
            g.Broker.BrokerDebug Order.Broker, "Cancelling Order from Option Navigator: " & Order.OrderText, True
            CancelOrder Order, False
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.CancelOptionNavigatorOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOptionNavigatorOrder
'' Description: Park the order that was sent to us by Option Navigator
'' Inputs:      Order, Spread Order?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ParkOptionNavigatorOrder(ByVal strMessage As String, Optional ByVal bSpreadOrder As Boolean = False)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    Dim SpreadTree As cGdTree           ' Spread tree
    
    If bSpreadOrder Then
        Set SpreadTree = New cGdTree
        SpreadTree.FromKeyValueString strMessage
        If SpreadTree.Exists("AccountNumber") Then
            If SpreadTree.Exists("AccessID") Then
                Set Order = g.Broker.OrderForAccessID(SpreadTree("AccountNumber"), CLng(Val(SpreadTree("AccessID"))))
            ElseIf SpreadTree.Exists("GenesisOrderID") Then
                Set Order = g.Broker.OrderForGenesisID(SpreadTree("AccountNumber"), SpreadTree("GenesisOrderID"))
            End If
        End If
    Else
        If Len(Parse(strMessage, vbTab, 2)) > 0 Then
            Set Order = g.Broker.OrderForAccessID(Parse(strMessage, vbTab, 3), CLng(Val(Parse(strMessage, vbTab, 1))))
        Else
            Set Order = g.Broker.OrderForGenesisID(Parse(strMessage, vbTab, 3), Parse(strMessage, vbTab, 1))
        End If
    End If
    
    If Order Is Nothing Then
        Set Order = New cPtOrder
        If bSpreadOrder Then
            OrderFromSpreadTree SpreadTree, Order
        Else
            OrderFromString strMessage, Order
        End If
        
        ' 06/15/2011 DAJ: Do a save up front here to get an AccessID for the order
        ' to send to Option Navigator via a OrderIdChange event...
        Order.Status = eTT_OrderStatus_Open
        Order.Save
    
        ' 06/15/2011 DAJ: If this is an Interactive Brokers order then we need to
        ' reassign the Genesis ID here...
        If (Order.Broker = eTT_AccountType_IntBrokers) Or (Order.Broker = eTT_AccountType_Ideal) Then
            Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Order.AccountID))
        End If
        
        g.Broker.BrokerDebug g.Broker.AccountTypeForID(Order.AccountID), "Parking new order from Option Navigator: " & Order.OrderText, True
        ParkOrder Order
    Else
        ' If the order is closed or pending, just send the order status back to Option
        ' Navigator for right now...
        If (IsOpenOrder(Order.Status, False) = False) Or (OrderIsPending(Order) = True) Then
            g.Broker.BrokerDebug Order.Broker, "Park Order Requested from Option Navigator, but order is closed: " & Order.OrderText, True
            'SendMessageToOptNav eGDOptNav_Order, Order.OptionNavigatorString(False)
            SendOrderToOptionNav Order, False
        Else
            ' DAJ 04/30/2012 - If the order is already parked, we want a subsequent park
            ' request able to change the order...
            If Order.Status = eTT_OrderStatus_Parked Then
                If bSpreadOrder Then
                    OrderFromSpreadTree SpreadTree, Order
                Else
                    OrderFromString strMessage, Order
                End If
            End If
            
            g.Broker.BrokerDebug Order.Broker, "Parking Order from Option Navigator: " & Order.OrderText, True
            ParkOrder Order
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.ParkOptionNavigatorOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOptionNavigatorOrder
'' Description: Submit the parked order that was sent to us by Option Navigator
'' Inputs:      Order, Spread Order?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SubmitOptionNavigatorOrder(ByVal strMessage As String, Optional ByVal bSpreadOrder As Boolean = False)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    Dim SpreadTree As cGdTree           ' Spread tree
    
    Set Order = Nothing
    If bSpreadOrder Then
        Set SpreadTree = New cGdTree
        SpreadTree.FromKeyValueString strMessage
        If SpreadTree.Exists("AccountNumber") And SpreadTree.Exists("AccessID") Then
            Set Order = g.Broker.OrderForAccessID(SpreadTree("AccountNumber"), CLng(Val(SpreadTree("AccessID"))))
            If Order Is Nothing Then
                OptNavLog "SubmitOrder not performed because order '" & SpreadTree("AccessID") & "' not found in account '" & SpreadTree("AccountNumber") & "'"
            End If
        Else
            OptNavLog "SubmitOrder not performed because either account number or access ID doesn't exist in string"
        End If
    Else
        Set Order = g.Broker.OrderForAccessID(Parse(strMessage, vbTab, 3), CLng(Val(Parse(strMessage, vbTab, 1))))
        If Order Is Nothing Then
            OptNavLog "SubmitOrder not performed because order '" & Parse(strMessage, vbTab, 1) & "' not found in account '" & Parse(strMessage, vbTab, 3) & "'"
        End If
    End If
    
    If Not Order Is Nothing Then
        ' If the order is closed or pending, just send the order status back to Option
        ' Navigator for right now...
        If Order.Status <> eTT_OrderStatus_Parked Then
            g.Broker.BrokerDebug Order.Broker, "Submit Order Requested from Option Navigator, but order is not parked: " & Order.OrderText, True
            'SendMessageToOptNav eGDOptNav_Order, Order.OptionNavigatorString(False)
            SendOrderToOptionNav Order, False
        Else
            ' 01/25/2011 DAJ: Need to reset the properties from the message in case
            ' the order has changed in Option Navigator while it was parked...
            If bSpreadOrder Then
                OrderFromSpreadTree SpreadTree, Order
            Else
                OrderFromString strMessage, Order
            End If
            
            ' 07/22/2011 DAJ: If this is an Interactive Brokers order then we need to
            ' reassign the Genesis ID here...
            If (Order.Broker = eTT_AccountType_IntBrokers) Or (Order.Broker = eTT_AccountType_Ideal) Then
                Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Order.AccountID))
            End If
            
            g.Broker.BrokerDebug Order.Broker, "Submitting Order from Option Navigator: " & Order.OrderText, True
            SubmitOrderFromOrder Order, "Option Navigator"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SubmitOptionNavigatorOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeGroupInfoForOrder
'' Description: Change the group info as appropriate for the given order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeGroupInfoForOrder(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    Dim astrMessage As New cGdArray     ' Message broken out into an array
    
    astrMessage.SplitFields strMessage, vbTab
    
    Set Order = g.Broker.OrderForAccessID(astrMessage(1), CLng(Val(astrMessage(0))))
    If Not Order Is Nothing Then
        g.Broker.BrokerDebug Order.Broker, "Group information changing for order '" & Order.GenesisOrderID & "' from '" & Str(Order.GroupID) & ";" & Order.GroupName & "' to '" & astrMessage(2) & ";" & astrMessage(3)
        Order.GroupID = CLng(Val(astrMessage(2)))
        Order.GroupName = astrMessage(3)
        
        Order.Save
    
        g.Broker.AddOrder Order
        OrderCallback Order
    
        'If g.nOptNavStatus >= eGDOptNavStatus_Loading Then
        '    SendMessageToOptNav eGDOptNav_Order, Order.OptionNavigatorString(False), False
        'End If
        SendOrderToOptionNav Order, False
    Else
        OptNavLog "ChangeGroupInfo not performed because order '" & astrMessage(0) & "' not found in account '" & astrMessage(1) & "'"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.ChangeGroupInfoForOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OptNavLog
'' Description: Send a string to the log file for the day
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OptNavLog(ByVal strMessage As String)
On Error Resume Next

    Dim fh As Integer                   ' File handle to open file with

    fh = FreeFile
    Open AddSlash(App.Path) & "OptNav\TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append Shared As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " - " & strMessage
        Close #fh
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OptNavMessageTypeString
'' Description: Convert an Option Navigator message type to a string
'' Inputs:      Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OptNavMessageTypeString(ByVal nOptNavMsg As eGDOptNavMessageType) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case nOptNavMsg
        Case eGDOptNav_UnknownMessage
            strReturn = "eGDOptNav_UnknownMessage"
        Case eGDOptNav_GetBrokers
            strReturn = "eGDOptNav_GetBrokers"
        Case eGDOptNav_Broker
            strReturn = "eGDOptNav_Broker"
        Case eGDOptNav_GetAccounts
            strReturn = "eGDOptNav_GetAccounts"
        Case eGDOptNav_Account
            strReturn = "eGDOptNav_Account"
        Case eGDOptNav_GetOrders
            strReturn = "eGDOptNav_GetOrders"
        Case eGDOptNav_Order
            strReturn = "eGDOptNav_Order"
        Case eGDOptNav_GetFills
            strReturn = "eGDOptNav_GetFills"
        Case eGDOptNav_Fill
            strReturn = "eGDOptNav_Fill"
        Case eGDOptNav_GetPositions
            strReturn = "eGDOptNav_GetPositions"
        Case eGDOptNav_Position
            strReturn = "eGDOptNav_Position"
        Case eGDOptNav_SymbolRequest
            strReturn = "eGDOptNav_SymbolRequest"
        Case eGDOptNav_Symbols
            strReturn = "eGDOptNav_Symbols"
        Case eGDOptNav_InfoRequest
            strReturn = "eGDOptNav_InfoRequest"
        Case eGDOptNav_Information
            strReturn = "eGDOptNav_Information"
        Case eGDOptNav_OptNavLoaded
            strReturn = "eGDOptNav_OptNavLoaded"
        Case eGDOptNav_Activate
            strReturn = "eGDOptNav_Activate"
        Case eGDOptNav_OptNavUnloaded
            strReturn = "eGDOptNav_OptNavUnloaded"
        Case eGDOptNav_Unload
            strReturn = "eGDOptNav_Unload"
        Case eGDOptNav_AddOrder
            strReturn = "eGDOptNav_AddOrder"
        Case eGDOptNav_AmendOrder
            strReturn = "eGDOptNav_AmendOrder"
        Case eGDOptNav_CancelOrder
            strReturn = "eGDOptNav_CancelOrder"
        Case eGDOptNav_GetSymbolInfo
            strReturn = "eGDOptNav_GetSymbolInfo"
        Case eGDOptNav_SymbolInfo
            strReturn = "eGDOptNav_SymbolInfo"
        Case eGDOptNav_ChainBuilt
            strReturn = "eGDOptNav_ChainBuilt"
        Case eGDOptNav_RequestChain
            strReturn = "eGDOptNav_RequestChain"
        Case eGDOptNav_CreateTicket
            strReturn = "eGDOptNav_CreateTicket"
        Case eGDOptNav_TicketSubmitted
            strReturn = "eGDOptNav_TicketSubmitted"
        Case eGDOptNav_ConnectToAccount
            strReturn = "eGDOptNav_ConnectToAccount"
        Case eGDOptNav_RiskGraphBuilt
            strReturn = "eGDOptNav_RiskGraphBuilt"
        Case eGDOptNav_RequestRiskGraph
            strReturn = "eGDOptNav_RequestRiskGraph"
        Case eGDOptNav_RequestPriceData
            strReturn = "eGDOptNav_RequestPriceData"
        Case eGDOptNav_PriceData
            strReturn = "eGDOptNav_PriceData"
        Case eGDOptNav_RequestGroupSymbols
            strReturn = "eGDOptNav_RequestGroupSymbols"
        Case eGDOptNav_GroupSymbols
            strReturn = "eGDOptNav_GroupSymbols"
        Case eGDOptNav_ParkOrder
            strReturn = "eGDOptNav_ParkOrder"
        Case eGDOptNav_SubmitOrder
            strReturn = "eGDOptNav_SubmitOrder"
        Case eGDOptNav_ChangeGroupInfo
            strReturn = "eGDOptNav_ChangeGroupInfo"
        Case eGDOptNav_CurrentAccount
            strReturn = "eGDOptNav_CurrentAccount"
        Case eGDOptNav_GetHistoricalFills
            strReturn = "eGDOptNav_GetHistoricalFills"
        Case eGDOptNav_HistoricalFills
            strReturn = "eGDOptNav_HistoricalFills"
        Case eGDOptNav_ClearGroup
            strReturn = "eGDOptNav_ClearGroup"
        Case eGDOptNav_OrderIdChanged
            strReturn = "eGDOptNav_OrderIdChanged"
        Case eGDOptNav_ExpirePosition
            strReturn = "eGDOptNav_ExpirePosition"
        Case eGDOptNav_FlattenPosition
            strReturn = "eGDOptNav_FlattenPosition"
        Case eGDOptNav_CancelAllOrder
            strReturn = "eGDOptNav_CancelAllOrder"
        Case eGDOptNav_ReversePosition
            strReturn = "eGDOptNav_ReversePosition"
        Case eGDOptNav_AddSpreadOrder
            strReturn = "eGDOptNav_AddSpreadOrder"
        Case eGDOptNav_SpreadOrder
            strReturn = "eGDOptNav_SpreadOrder"
        Case eGDOptNav_AmendSpreadOrder
            strReturn = "eGDOptNav_AmendSpreadOrder"
        Case eGDOptNav_ParkSpreadOrder
            strReturn = "eGDOptNav_ParkSpreadOrder"
        Case eGDOptNav_SubmitSpreadOrder
            strReturn = "eGDOptNav_SubmitSpreadOrder"
        Case eGDOptNav_RenameGroup
            strReturn = "eGDOptNav_RenameGroup"
        Case Else
            strReturn = Str(nOptNavMsg)
    End Select
    
    OptNavMessageTypeString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.OptNavMessageTypeString"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSymbolInfoForOptNav
'' Description: Get symbol information for Option Navigator
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      Account Number, Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetSymbolInfoForOptNav(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strAccountNumber As String      ' Account number passed in
    Dim strSymbol As String             ' Symbol passed in
    Dim nBroker As eTT_AccountType      ' Broker for the account number
    Dim Bars As New cGdBars             ' Bars object
    Dim strSymbolInfo As String         ' String to send to Option Navigator
    
    strAccountNumber = Parse(strMessage, vbTab, 1)
    strSymbol = Parse(strMessage, vbTab, 2)
    nBroker = g.Broker.AccountTypeForNumber(strAccountNumber)
    
    ' DAJ 03/16/2009: If an individual futures contract is sent over to us from
    ' Option Navigator and we are unable to get market data for it (i.e. the
    ' individual contract doesn't exist in Trade Navigator), then try getting the
    ' market data for the 57 continuous which is more likely to exist...
    If SetBarProperties(Bars, strSymbol) = False Then
        If InStr(strSymbol, " ") = 0 Then
            If InStr(strSymbol, "-") <> 0 Then
                strSymbol = Parse(strSymbol, "-", 1) & "-057"
                SetBarProperties Bars, strSymbol
            End If
        End If
    End If
    
    strSymbolInfo = g.Broker.SymbolInformation(nBroker, strSymbol)
    
    ' DAJ 05/09/2011: Option Nav is going to start using this data for display purposes,
    ' so we want to send the bar properties whether or not the symbol is available
    ' for the given broker.  If the symbol isn't available for trading through the broker,
    ' then we will send blanks now for the bit masks instead of not sending the info at all...
    If Len(strSymbolInfo) = 0 Then
        strSymbolInfo = "" & vbTab & ""
    End If
        
    strSymbolInfo = strSymbolInfo & vbTab & Str(Bars.TickValue)
    strSymbolInfo = strSymbolInfo & vbTab & Str(Bars.TickMove)
    strSymbolInfo = strSymbolInfo & vbTab & Str(Bars.Prop(eBARS_MinMoveInTicks))
    strSymbolInfo = strSymbolInfo & vbTab & Bars.PriceThresholds
    strSymbolInfo = strSymbolInfo & vbTab & Bars.SecondaryMinMoves
    
    ' DAJ 05/09/2011: If we are passed an underlying symbol, send the bars information
    ' for the option as well...
    If InStr(strSymbol, " ") = 0 Then
        SetBarProperties Bars, strSymbol & " C1234"
        
        strSymbolInfo = strSymbolInfo & vbTab & Str(Bars.TickValue)
        strSymbolInfo = strSymbolInfo & vbTab & Str(Bars.TickMove)
        strSymbolInfo = strSymbolInfo & vbTab & Str(Bars.Prop(eBARS_MinMoveInTicks))
        strSymbolInfo = strSymbolInfo & vbTab & Bars.PriceThresholds
        strSymbolInfo = strSymbolInfo & vbTab & Bars.SecondaryMinMoves
    End If
    
    SendMessageToOptNav eGDOptNav_SymbolInfo, strSymbolInfo

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.GetSymbolInfoForOptNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RequestOptionChainStructure
'' Description: Request option chain structure for a symbol Option Navigator
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RequestOptionChainStructure(ByVal strMessage As String)
On Error GoTo ErrSection:

    SendMessageToOptNav eGDOptNav_RequestChain, strMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.RequestOptionChainStructure"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OptionChainBuilt
'' Description: Option chain structure has been built and file is ready
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      Symbol, Error (-1 for error, blank otherwise)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OptionChainBuilt(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim frm As Form                     ' Chart form
    Dim strSymbol As String             ' Symbol from the message
    
    strSymbol = Parse(strMessage, vbTab, 1)
    
    For lIndex = 0 To Forms.Count - 1
        If IsFrmChart(Forms(lIndex)) Then
            Set frm = Forms(lIndex)
            If UCase(Parse(frm.Chart.Symbol, "-", 1)) = UCase(Parse(strSymbol, "-", 1)) Then
                frm.OptionDataAvailable eGDOptNav_ChainBuilt, strMessage
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.OptionChainBuilt"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateTicketInOptNav
'' Description: Tell Option Navigator to create an order ticket with given info
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      BaseSymbol;Qty, OptionSymbol;Qty, OptionSymbol;Qty, ..., ChartID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateTicketInOptNav(ByVal strMessage As String)
On Error GoTo ErrSection:

    SendMessageToOptNav eGDOptNav_CreateTicket, strMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.CreateTicketInOptNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectToAccount
'' Description: Connect to the given account
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      AccountNumber
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ConnectToAccount(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account object

    Set Account = g.Broker.Account(strMessage)
    If Not Account Is Nothing Then
        If Account.ConnectionStatus = eGDConnectionStatus_Disconnected Then
            g.Broker.ToggleConnectionForAccount strMessage, Account.UserName
        Else
            SendMessageToOptNav eGDOptNav_Account, AccountToString(Account, False)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.ConnectToAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RequestRiskGraphStructure
'' Description: Request risk graph structure for a symbol Option Navigator
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      ChartID, GraphType, PriceInfo, Leg1Info, Leg2Info, ...
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RequestRiskGraphStructure(ByVal strMessage As String)
On Error GoTo ErrSection:

    SendMessageToOptNav eGDOptNav_RequestRiskGraph, strMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.RequestRiskGraphStructure"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RiskGraphBuilt
'' Description: Risk graph structure has been built and file is ready
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      ChartID, GraphType, PriceInfo, LineNow, Line1/3, Line2/3, LineExpire
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RiskGraphBuilt(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim frm As Form                     ' Chart form
    Dim strSymbol As String             ' Symbol from the message
    Dim lChartHandle As Long
    
    strSymbol = Parse(strMessage, vbTab, 1)
    lChartHandle = CLng(Val(Parse(strMessage, vbTab, 1)))
    
    For lIndex = 0 To Forms.Count - 1
        If IsFrmChart(Forms(lIndex)) Then
            Set frm = Forms(lIndex)
            If frm.Chart.geChartObj = lChartHandle Then
                frm.OptionDataAvailable eGDOptNav_RiskGraphBuilt, strMessage
                Exit For
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.RiskGraphBuilt"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TicketSubmitted
'' Description: Ticket generated by CreateTicket has been submitted or parked
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      ChartID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TicketSubmitted(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim frm As Form                     ' Chart form
    Dim strSymbol As String             ' Symbol from the message
    Dim lChartHandle As Long
    
    strSymbol = Parse(strMessage, vbTab, 1)
    lChartHandle = CLng(Val(Parse(strMessage, vbTab, 1)))
    
    For lIndex = 0 To Forms.Count - 1
        If IsFrmChart(Forms(lIndex)) Then
            Set frm = Forms(lIndex)
            If frm.Chart.geChartObj = lChartHandle Then
                frm.OptionDataAvailable eGDOptNav_TicketSubmitted, strMessage
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.TicketSubmitted"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildBrokerList
'' Description: Build a broker list for Option Navigator
'' Inputs:      None
'' Returns:     Broker List
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildBrokerList() As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As New cGdArray      ' Array to return from the function
    Dim nBroker As eTT_AccountType      ' Broker type

    astrReturn.Create eGDARRAY_Strings
    
#If 1 Then
    For nBroker = 1 To kNumBrokers - 1
        If g.Broker.IsBrokerUserOptions(nBroker) Then
            astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & Str(g.Broker.MaxOrderLegs(nBroker)) & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker) & vbTab & g.Broker.SecTypesForBroker(nBroker)
        End If
    Next nBroker
#Else
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_AmpCqg) Then
        nBroker = eTT_AccountType_AmpCqg
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_CQG) Then
        nBroker = eTT_AccountType_CQG
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_FptCqg) Then
        nBroker = eTT_AccountType_FptCqg
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_FptOec) Then
        nBroker = eTT_AccountType_FptOec
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_Ideal) Then
        nBroker = eTT_AccountType_Ideal
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_IntBrokers) Then
        nBroker = eTT_AccountType_IntBrokers
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
'    If g.Broker.IsBrokerUserOptions(eTT_AccountType_LindWaldock) Then
'        nBroker = eTT_AccountType_LindWaldock
'        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "3" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
'    End If
    
'    If g.Broker.IsBrokerUserOptions(eTT_AccountType_ManExpress) Then
'        nBroker = eTT_AccountType_ManExpress
'        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "3" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
'    End If

    If g.Broker.IsBrokerUserOptions(eTT_AccountType_Oec) Then
        nBroker = eTT_AccountType_Oec
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_Optimus) Then
        nBroker = eTT_AccountType_Optimus
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_OpVest) Then
        nBroker = eTT_AccountType_OpVest
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
'    If g.Broker.IsBrokerUserOptions(eTT_AccountType_PFG) Then
'        nBroker = eTT_AccountType_PFG
'        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
'    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_Rithmic) Then
        nBroker = eTT_AccountType_Rithmic
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_RjoCqg) Then
        nBroker = eTT_AccountType_RjoCqg
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_Vision) Then
        nBroker = eTT_AccountType_Vision
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_ZenFire) Then
        nBroker = eTT_AccountType_ZenFire
        astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    End If
    
    nBroker = eTT_AccountType_SimStream
    astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
    
    nBroker = eTT_AccountType_SimBroker
    astrReturn.Add Str(nBroker) & vbTab & g.Broker.BrokerName(nBroker) & vbTab & "1" & vbTab & g.Broker.StartingGenesisIdForOptNav(nBroker)
#End If
    
    Set BuildBrokerList = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.BuildBrokerList"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildAccountListForBroker
'' Description: Build an account list for Option Navigator for given broker
'' Inputs:      Broker
'' Returns:     Account List
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildAccountListForBroker(ByVal nBroker As eTT_AccountType) As cGdArray
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Broker info object
    Dim Accounts As cPtAccounts         ' Collection of accounts
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrReturn As New cGdArray      ' Array to return from the function
    
    astrReturn.Create eGDARRAY_Strings
    
    Set BInfo = g.Broker.BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Accounts = BInfo.Accounts
        If Not Accounts Is Nothing Then
            For lIndex = 1 To Accounts.Count
                astrReturn.Add AccountToString(Accounts(lIndex), True)
            Next lIndex
        End If
    End If

    Set BuildAccountListForBroker = astrReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.BuildAccountListForBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildAccountList
'' Description: Build an account list for Option Navigator
'' Inputs:      None
'' Returns:     Account List
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildAccountList() As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array to return from the function
    Dim nBroker As eTT_AccountType      ' Broker type
    Dim astrAccounts As cGdArray        ' Account list for broker
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
#If 1 Then
    For nBroker = 1 To kNumBrokers - 1
        If g.Broker.IsBrokerUserOptions(nBroker) Then
            Set astrAccounts = BuildAccountListForBroker(nBroker)
            If Not astrAccounts Is Nothing Then
                astrReturn.AppendFromArray astrAccounts
            End If
        End If
    Next nBroker
#Else
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_AmpCqg) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_AmpCqg)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_CQG) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_CQG)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_FptCqg) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_FptCqg)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_FptOec) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_FptOec)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_Ideal) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_Ideal)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_IntBrokers) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_IntBrokers)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
'    If g.Broker.IsBrokerUserOptions(eTT_AccountType_LindWaldock) Then
'        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_LindWaldock)
'        If Not astrAccounts Is Nothing Then
'            astrReturn.AppendFromArray astrAccounts
'        End If
'    End If
    
'    If g.Broker.IsBrokerUserOptions(eTT_AccountType_ManExpress) Then
'        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_ManExpress)
'        If Not astrAccounts Is Nothing Then
'            astrReturn.AppendFromArray astrAccounts
'        End If
'    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_Oec) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_Oec)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_Optimus) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_Optimus)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_OpVest) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_OpVest)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
'    If g.Broker.IsBrokerUserOptions(eTT_AccountType_PFG) Then
'        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_PFG)
'        If Not astrAccounts Is Nothing Then
'            astrReturn.AppendFromArray astrAccounts
'        End If
'    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_Rithmic) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_Rithmic)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_RjoCqg) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_RjoCqg)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_RobbinsCqg) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_RobbinsCqg)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
    Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_SimStream)
    If Not astrAccounts Is Nothing Then
        astrReturn.AppendFromArray astrAccounts
    End If
    
    Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_SimBroker)
    If Not astrAccounts Is Nothing Then
        astrReturn.AppendFromArray astrAccounts
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_Vision) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_Vision)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_VisionCqg) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_VisionCqg)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
    
    If g.Broker.IsBrokerUserOptions(eTT_AccountType_ZenFire) Then
        Set astrAccounts = BuildAccountListForBroker(eTT_AccountType_ZenFire)
        If Not astrAccounts Is Nothing Then
            astrReturn.AppendFromArray astrAccounts
        End If
    End If
#End If

    Set BuildAccountList = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.BuildAccountList"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildOptionNavCfg
'' Description: Build the Option Navigator Configuration file
'' Inputs:      Symbol, IP
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildOptionNavCfg(ByVal strSymbol As String, ByVal strIP As String)
On Error GoTo ErrSection:

    Dim astrOutput As New cGdArray      ' Array to output to the configuration file
    Dim astrBrokerList As cGdArray      ' Broker list
    Dim astrAccountList As cGdArray     ' Account list
    Dim lIndex As Long                  ' Index into a for loop
    
    astrOutput.Create eGDARRAY_Strings
    
    ' User Information...
    astrOutput.Add Str(RI_GetDataServiceID) & vbTab & Trim(UCase(g.strAuthorizationString)) & vbTab _
                    & strIP & vbTab & strSymbol & vbTab & Str(NextGroupID) & vbTab
        
    ' Broker List...
    Set astrBrokerList = BuildBrokerList
    astrOutput.Add "[Brokers]"
    astrOutput.Add "BEGIN"
    
    If Not astrBrokerList Is Nothing Then
        For lIndex = 0 To astrBrokerList.Size - 1
            astrOutput.Add astrBrokerList(lIndex)
        Next lIndex
    End If
    
    astrOutput.Add "END"

    ' Account List...
    Set astrAccountList = BuildAccountList
    astrOutput.Add "[Accounts]"
    astrOutput.Add "BEGIN"
    
    If Not astrAccountList Is Nothing Then
        For lIndex = 0 To astrAccountList.Size - 1
            astrOutput.Add astrAccountList(lIndex)
        Next lIndex
    End If
    
    astrOutput.Add "END"

    ' Dump the configuration file...
    FileFromString FilePath(App.Path) & "OptionNav\OptionNav.CFG", astrOutput.JoinFields(vbCrLf) & vbCrLf

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.BuildOptionNavCfg"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartOptNav
'' Description: Start up Option Navigator
'' Inputs:      Symbol or SymbolID
'' Returns:     True if started, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StartOptNav(Optional ByVal vSymbolOrSymbolID As Variant = "", _
    Optional ByVal bRequireRT As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim lChartSymbolID As Long          ' Symbol ID for the active chart
    Dim strIP As String                 ' IP address to send to Option Navigator
    Dim strPath As String               ' Path for Option Navigator
    Dim strSymbol As String             ' Symbol to send along to Option Navigator
    Dim strElec As String               ' Electronic equivalent for the symbol
    Dim lSymbolID As Long               ' Symbol ID for the symbol passed in
    Dim astrSymbols As New cGdArray     ' List of symbols back from symbol selector
    Dim bReturn As Boolean              ' Return value for the function
    
    Dim strMsgRT As String
    
    If bRequireRT Then
        strMsgRT = "Streaming data is required. Would you like to start streaming now?"
    Else
        strMsgRT = "Streaming data is recommended. Would you like to start streaming now?"
    End If

    bReturn = False
    
    ' get IP info for Options Navigator
    strIP = OptNavIP
    If HasModule("OPTNAV") And Len(strIP) > 0 And FileExist(OptNavExeFile) Then
        ' start streaming?  TLB 8/15/2014: now only ask if required
        If bRequireRT And HasModule("RTG,RTE") Then
            If Not g.RealTime.Active Then       'aardvark 5374
                If InfBox(strMsgRT, "?", "+Yes|-No", "RealTime Server") = "Y" Then
                    If ProcessIsBusy Then GoTo ErrExit
                    g.RealTime.Init True
                    'waiting till realtime is done starting gives illusion of OptionNav starting faster for some reason
                    While g.RealTime.ConnectionStatus = eGDConnectionStatus_Connecting
                        DoEvents
                    Wend
                End If
            End If
        End If
        
        If bRequireRT And Not g.RealTime.Active Then
            GoTo ErrExit
        ElseIf SendMessageToOptNav(eGDOptNav_Activate, "") = 0 Then     ' if OptNav already running, just have it give focus to itself
            ' check if DotNet dependencies need to be installed
            If HasDotNet Then
                strSymbol = Trim(GetSymbol(vSymbolOrSymbolID))
                If Len(strSymbol) = 0 Then
                    ' get symbol of active chart
                    lChartSymbolID = 0
                    If Not ActiveChart Is Nothing Then
                        lChartSymbolID = ActiveChart.Chart.SymbolID
                        ' use symbol of active chart as default
                        If lSymbolID = 0 Then lSymbolID = lChartSymbolID
                    End If
                    Set astrSymbols = frmSymbolSelector.ShowMe(g.SymbolPool.SymbolForID(lSymbolID), _
                            False, True, "Display the Options for which underlying symbol ...")
                    strSymbol = astrSymbols(0)
                End If
                If Len(strSymbol) > 0 Then
                    ' if a future, convert to electronic
                    If SecurityType(strSymbol) = "F" Then
                        strElec = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
                        If Len(strElec) > 0 Then
                            strSymbol = strElec
                        End If
                        strSymbol = RollSymbolForDate(Parse(strSymbol, "-", 1) & "-067")
                    End If
                    strPath = FilePath(App.Path) & "OptionNav\"
                    
                    ' Put general info into the CFG file:  DataSrvID, Enablements, IP addresses, Symbol
                    BuildOptionNavCfg strSymbol, strIP
                    
                    ' Start Option Navigator...
                    If g.nOptNavStatus = eGDOptNavStatus_Unloaded Then
                        g.nOptNavStatus = eGDOptNavStatus_Loading
                        RunProcess strPath & "OptionNav.EXE", Chr(34) & strPath & "OptionNav.CFG" & Chr(34)
                        DebugLog "OptionNav process invoked."
                        bReturn = True
                    End If
                End If
            End If
        Else
            bReturn = True
        End If
    End If
    
    StartOptNav = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.StartOptNav"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddSymbolsToQuoteBoard
'' Description: Add the given symbols to the quote board
'' Inputs:      Comma Delimited list of Symbols
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddSymbolsToQuoteBoard(ByVal strSymbols As String)
On Error GoTo ErrSection:

    Dim astrSymbols As cGdArray         ' Array of symbols to add to the quote board
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrSymbols = New cGdArray
    astrSymbols.SplitFields strSymbols, ","
    
    For lIndex = 0 To astrSymbols.Size - 1
        frmQuotes.AddSymbol -1&, "Daily", astrSymbols(lIndex)
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.AddSymbolsToQuoteBoard"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IncludeAccountToOptionNav
'' Description: Include an item with the given account to Option Navigator?
'' Inputs:      Account Number
'' Returns:     True if include, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IncludeAccountToOptionNav(ByVal strAccount As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If Len(m.strCurrentAccount) > 0 Then
        bReturn = (strAccount = m.strCurrentAccount)
    End If
    
    IncludeAccountToOptionNav = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.IncludeAccountToOptionNav"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IncludeBrokerToOptionNav
'' Description: Include an item with the given broker to Option Navigator?
'' Inputs:      Broker
'' Returns:     True if include, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IncludeBrokerToOptionNav(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If m.nCurrentBroker > -1& Then
        bReturn = (nBroker = m.nCurrentBroker)
    End If
    
    IncludeBrokerToOptionNav = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.IncludeBrokerToOptionNav"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetHistoricalFillsForOptionNav
'' Description: Get a list of historical fills for Option Navigator
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetHistoricalFillsForOptionNav()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As cPtFill                 ' Fill object
    Dim astrFills As New cGdArray       ' Array of fills to dump to file
    Dim strFileName As String           ' Filename for the output file
    
    ' 03/31/2011 DAJ: Don't send snapshot fills in the historical fills file (#6233)...
    If Len(m.strCurrentAccount) = 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] WHERE [IsSnapshot]=0;", dbOpenDynaset)
    Else
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] WHERE [AccountID]=" & Str(g.Broker.AccountIDForNumber(m.strCurrentAccount)) & " AND [IsSnapshot]=0;", dbOpenDynaset)
    End If
    
    Do While Not rs.EOF
        Set Fill = New cPtFill
        If Fill.Load(rs!FillID, rs) Then
            astrFills.Add FillToString(Fill, True)
        End If
        
        rs.MoveNext
    Loop
    
    strFileName = AddSlash(App.Path) & "OnFills.TXT"
    astrFills.ToFile strFileName
    SendMessageToOptNav eGDOptNav_HistoricalFills, strFileName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.GetHistoricalFillsForOptionNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendOrderIdChangeToOptNav
'' Description: Send an Order ID change event to Option Navigator
'' Inputs:      Account, Old Order ID, New Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendOrderIdChangeToOptNav(ByVal vAccountNumberOrID As Variant, ByVal strOldOrderID As String, strNewOrderID As String)
On Error GoTo ErrSection:

    If g.nOptNavStatus >= eGDOptNavStatus_Loading Then
        If IncludeAccountToOptionNav(g.Broker.GetAccountNumber(vAccountNumberOrID)) Then
            SendMessageToOptNav eGDOptNav_OrderIdChanged, strOldOrderID & vbTab & strNewOrderID
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SendOrderIdChangeToOptNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleExpirePosition
'' Description: Handle a request from Option Nav to expire a position
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      Account, Symbol, Expiration Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleExpirePosition(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Fields in the message
    
    If Len(strMessage) > 0 Then
        Set astrFields = New cGdArgs
        astrFields.SplitFields strMessage, vbTab
        
        g.Broker.FlattenExpiredPosition astrFields(0), astrFields(1), JulFromLong(CLng(Val(astrFields(2))))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.HandleExpirePosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleFlattenFromOptNav
'' Description: Handle a request from Option Nav to flatten a position
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      Account, Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleFlattenFromOptNav(ByVal strMessage As String)
On Error GoTo ErrSection:
    
    Dim astrFields As cGdArray          ' Fields in the message
    Dim lAccountID As Long              ' Account ID for the account passed in
    
    If Len(strMessage) > 0 Then
        Set astrFields = New cGdArray
        astrFields.SplitFields strMessage, vbTab
        
        lAccountID = g.Broker.AccountIDForNumber(astrFields(0))
        FlattenForSymbol lAccountID, astrFields(1), 0&
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mOptNav.HandleFlattenFromOptNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleCancelAllFromOptNav
'' Description: Handle a request from Option Nav to cancel all orders for a symbol
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      Account, Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleCancelAllFromOptNav(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Fields in the message
    Dim lAccountID As Long              ' Account ID for the account passed in
    
    If Len(strMessage) > 0 Then
        Set astrFields = New cGdArray
        astrFields.SplitFields strMessage, vbTab
        
        lAccountID = g.Broker.AccountIDForNumber(astrFields(0))
        CancelAllForSymbol lAccountID, astrFields(1), 0&
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mOptNav.HandleCancelAllFromOptNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleReverseFromOptNav
'' Description: Handle a request from Option Nav to reverse a position
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      Account, Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleReverseFromOptNav(ByVal strMessage As String)
On Error GoTo ErrSection:
    
    Dim astrFields As cGdArray          ' Fields in the message
    Dim lAccountID As Long              ' Account ID for the account passed in
    
    If Len(strMessage) > 0 Then
        Set astrFields = New cGdArray
        astrFields.SplitFields strMessage, vbTab
        
        lAccountID = g.Broker.AccountIDForNumber(astrFields(0))
        ReverseForSymbol lAccountID, astrFields(1), 0&
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mOptNav.HandleReverseFromOptNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleExpirationDateFromOptNav
'' Description: Handle an expiration date from Option Nav
'' Inputs:      Message
'' Returns:     None
''
'' Fields:      Symbol, Expiration Date (YYYYMMDD)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleExpirationDateFromOptNav(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol from the message
    Dim lExpirationDate As Long         ' Expiration date from the message
    
    strSymbol = Parse(strMessage, vbTab, 1)
    lExpirationDate = JulFromLong(CLng(Val(Parse(strMessage, vbTab, 2))))
    
    g.Broker.SetExpirationDateForSymbol strSymbol, lExpirationDate

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.HandleExpirationDateFromOptNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RequestExpirationDateFromOptNav
'' Description: Request an expiration date from Option Nav
'' Inputs:      Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RequestExpirationDateFromOptNav(ByVal strSymbol As String)
On Error GoTo ErrSection:

    SendMessageToOptNav eGDOptNav_GetExpirationDate, strSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.RequestExpirationDateFromOptNav"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RenameOptionNavGroup
'' Description: Rename the Option Nav group with the given ID
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RenameOptionNavGroup(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim renameMessage As cBrokerMessage ' Message broken out into tokens
    Dim lGroupID As Long                ' Group ID
    Dim strNewName As String            ' New group name
    Dim rs As Recordset                 ' Recordset into the database
    Dim nBroker As eTT_AccountType      ' Broker for the order
    Dim alBrokers As cGdArray           ' Array of brokers
    Dim lPos As Long                    ' Position in the array
    Dim lIndex As Long                  ' Index into a for loop
    Dim BrokerInfo As cBrokerInfo       ' Broker info object
    
    Set renameMessage = New cBrokerMessage
    renameMessage.FromString strMessage
    lGroupID = CLng(Val(renameMessage("GroupID")))
    strNewName = renameMessage("NewGroupName")
    
    Set alBrokers = New cGdArray
    alBrokers.Create eGDARRAY_Longs
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [GroupID]=" & Str(lGroupID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Edit
        rs!GroupName = strNewName
        rs.Update
        
        nBroker = g.Broker.AccountTypeForID(rs!AccountID)
        If alBrokers.BinarySearch(nBroker, lPos) = False Then
            alBrokers.Add nBroker, lPos
        End If
        
        rs.MoveNext
    Loop
    
    For lIndex = 0 To alBrokers.Size - 1
        Set BrokerInfo = g.Broker.BrokerInfo(alBrokers(lIndex))
        If Not BrokerInfo Is Nothing Then
            BrokerInfo.RenameOptionNavGroup lGroupID, strNewName
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.RenameOptionNavGroup"
    
End Sub

Private Sub SyncOptNavSymbolsWithSalmon(ByVal strMessage$)
On Error GoTo ErrSection:

    Dim i&, s$, strChainBase$, bChain As Boolean
    Dim aStrings() As String
    Dim aRemove As New cGdArray, aAdd As New cGdArray, aSymbol As New cGdArray
    Dim SymInf As cSymbolInfo
    
    If Not g.RealTime.SalmonIsRunning Then Exit Sub
    
    ' parse the OptNav message into the separate "commands"
    If Len(strMessage) = 0 Then Exit Sub
    aStrings = Split(UCase(strMessage), vbCrLf)
    
    ' read each command from OptNav
    For i = LBound(aStrings) To UBound(aStrings)
        s = aStrings(i)
        Select Case Left(s, 2)
        Case "B:" ' BASE symbol for the Option Chain
            ' (NOTE: this appears to have become OBSOLETE)
            strChainBase = s
        Case "A:" ' ADD Option Chain symbols
            ' (can pass all the symbols at once to Salmon as an array)
            bChain = True
            aAdd.SplitFields s, vbTab
            aAdd(0) = Trim(Mid(aAdd(0), 3))
        Case "P:" ' ADD Portfolio symbols
            ' (need to pass each symbol to Salmon one at a time)
            bChain = False
            aAdd.SplitFields s, vbTab
            aAdd(0) = Trim(Mid(aAdd(0), 3))
        Case "R:" ' REMOVE symbols
            aRemove.SplitFields s, vbTab
            aRemove(0) = Trim(Mid(aRemove(0), 3))
        End Select
    Next
    
    ' do the remove now (before we add the new symbols)
    If aRemove.Size > 0 Then
        RemoveEODForOptionsNavigator aRemove.ArrayHandle
    End If
    
    For i = 0 To aAdd.Size - 1
        s = Trim(aAdd(i))
        If Len(s) > 0 Then
            Set SymInf = g.RealTime.SymbolInfo(s)
            If Not SymInf Is Nothing Then
                If bChain And InStr(s, " ") > 0 Then
                    ' when adding a chain and we get to the first option symbol,
                    ' just pass all the option symbols at once along with the Base
                    If i > 0 Then
                        aAdd.Remove 0, i ' (remove the symbols prior to this)
                    End If
                    SymInf.AddForOptNav aAdd, strChainBase
                    Exit For
                Else
                    ' for underlying or when passing from the portfolio,
                    ' just pass every symbol one-by-one
                    aSymbol.Size = 1
                    aSymbol(0) = s
                    SymInf.AddForOptNav aSymbol
                End If
            End If
        End If
    Next

ErrExit:
    Set SymInf = Nothing
    Set aSymbol = Nothing
    Set aRemove = Nothing
    Set aAdd = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.SyncOptNavSymbolsWithSalmon"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountToString
'' Description: Compile a delimited string to pass to Option Navigator for an account
'' Inputs:      Account, Refresh?, Connection Status, Removed?
'' Returns:     Delimited String
''
'' Fields:      Account Number, Account Name, Account Type, Closed Balance,
''              Type of Account, Connection Status, Colors, Fill Match Mode
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AccountToString(ByVal Account As cPtAccount, ByVal bRefresh As Boolean, Optional ByVal nConnectionStatus As eGDConnectionStatus = -1, Optional ByVal bRemoved As Boolean = False) As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Account information to pass along
    Dim strButtonFace As String         ' Button face color
                
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings, 8
    
    astrReturn(0) = Account.AccountNumber
    astrReturn(1) = Account.Name
    astrReturn(2) = Str(Account.AccountType)
    astrReturn(3) = Format(Account.CurrentClosedBalance, "0.00")
    
    ' Account type (0=Simulated, 1=Broker Live, 2=Broker Demo)
    astrReturn(4) = Str(Account.TypeOfAccount)
    
    ' Connection status (0=Disconnected, 1=Disconnecting, 2=Connecting, 3=Connected, 11=Remove Account)
    If bRemoved Then
        astrReturn(5) = "11"
    ElseIf nConnectionStatus = -1 Then
        astrReturn(5) = Str(g.Broker.ConnectionStatusForAccount(Account.AccountID))
    Else
        astrReturn(5) = Str(nConnectionStatus)
    End If
    
    ' Colors (Flat, Long, Short if Broker Live, Blank otherwise)
    If Account.TypeOfAccount = eGDTypeOfAccount_BrokerLive Then
        astrReturn(6) = Str(kFrameLive) & "," & Str(kFrameLong) & "," & Str(kFrameShort)
    Else
        strButtonFace = Str(GetSysColor(vbButtonFace And &HFF&))
        astrReturn(6) = strButtonFace & "," & strButtonFace & "," & strButtonFace
    End If
    
    ' Fill Match Mode (0=FIFO, 1=LIFO)
    astrReturn(7) = Str(Account.FillMatchMode)
                
    AccountToString = astrReturn.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.AccountToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountPositionToString
'' Description: Compile a delimited string to pass to Option Navigator for a position
'' Inputs:      Account Position, Refresh?
'' Returns:     Delimited String
''
'' Fields:      Account Number, Symbol, Current Position, Current Average Entry,
''              Carried Position, Carried Average Entry, Refresh, Expiration Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AccountPositionToString(ByVal AcctPos As cAccountPosition, ByVal bRefresh As Boolean) As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Position information to pass along

    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings, 8
        
    astrReturn(0) = g.Broker.AccountNumberForID(AcctPos.AccountID)
    astrReturn(1) = AcctPos.Symbol
    astrReturn(2) = Str(AcctPos.CurrentPositionSnapshot)
    astrReturn(3) = Str(AcctPos.AverageEntrySnapshot)
    astrReturn(4) = Str(AcctPos.CurrentPosition)
    astrReturn(5) = Str(AcctPos.AverageEntry)
    astrReturn(6) = Str(Abs(CLng(bRefresh)))
    astrReturn(7) = Format(AcctPos.ExpirationDate, "YYYYMMDD")
    
    AccountPositionToString = astrReturn.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.AccountPositionToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillToString
'' Description: Compile a delimited string to pass to Option Navigator for a fill
'' Inputs:      Fill, Refresh?
'' Returns:     Delimited String
''
'' Fields:      Account, Fill ID, Genesis Order ID, Broker Order ID, Broker Fill ID,
''              Symbol, Fill Date, Fill Price, Fill Quantity, Group ID, Group Name,
''              Refresh?, Snapshot?, Access Order ID, Reserved, Buy?, Expiration Date,
''              Lot Size
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FillToString(ByVal Fill As cPtFill, ByVal bRefresh As Boolean) As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Fill information to pass along
    Dim rs As Recordset                 ' Recordset into the database
    Dim dExpirationDate As Double       ' Expiration date for the symbol
    Dim dLotSize As Double              ' Lot size
                
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings, 18
                
    astrReturn(0) = g.Broker.AccountNumberForID(Fill.AccountID)
    astrReturn(1) = Str(Fill.FillID)
    astrReturn(3) = Fill.BrokerOrderID
    astrReturn(4) = Fill.BrokerID
    astrReturn(5) = Fill.Symbol
    astrReturn(6) = Str(Fill.FillDate)
    astrReturn(7) = Str(Fill.Price)
    astrReturn(8) = Str(Fill.Quantity)
    astrReturn(11) = Str(Abs(CLng(bRefresh)))
    astrReturn(12) = Str(Abs(CLng(Fill.IsSnapshot)))
    astrReturn(13) = Str(Fill.OrderID)
    astrReturn(14) = ""
    astrReturn(15) = Str(Abs(CLng(Fill.Buy)))
    
    ' If the order does not exist in the database, set smarter defaults (especially set
    ' the Group ID to 0 instead of a blank string because Option Navigator does not
    ' like the blank string).  02/20/2009 DAJ...
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [OrderID]=" & Str(Fill.OrderID) & ";", dbOpenDynaset)
    If (rs.BOF And rs.EOF) Then
        astrReturn(2) = ""
        astrReturn(9) = "0"
        astrReturn(10) = ""
    Else
        astrReturn(2) = rs!GenesisOrderID
        astrReturn(9) = Str(rs!GroupID)
        astrReturn(10) = rs!GroupName
    End If
    
    If g.Broker.LoadSymbolInfo(Fill.SymbolOrSymbolID, dExpirationDate, dLotSize) Then
        astrReturn(16) = Str(dExpirationDate)
        astrReturn(17) = Str(dLotSize)
    Else
        Select Case SecurityType(Fill.SymbolOrSymbolID, True)
            Case "F", "FO"
                astrReturn(16) = Str(JulFromLong(Fill.Bars.LastDayOfContractMonth))
                astrReturn(17) = "1"
            
            Case "S"
                astrReturn(16) = ""
                astrReturn(17) = "1"
            
            Case "SO", "IO"
                astrReturn(16) = Str(JulFromLong(CLng(Val(Parse(Fill.Symbol, " ", 2)))))
                astrReturn(17) = "100"
        
        End Select
    End If
    
    FillToString = astrReturn.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.FillToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderLegToString
'' Description: Compile a delimited string to pass to Option Navigator for an order leg
'' Inputs:      Order Leg
'' Returns:     Delimited String
''
'' Fields:      Buy?, Multiplier, Symbol, Entry?
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderLegToString(ByVal OrderLeg As cOrderLeg) As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Order Leg information to pass along
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings, 4
    
    astrReturn(0) = Str(Abs(CLng(OrderLeg.IsBuy)))
    astrReturn(1) = Str(OrderLeg.Multiplier)
    astrReturn(2) = OrderLeg.Symbol
    astrReturn(3) = Str(Abs(CLng(OrderLeg.IsEntry)))
    
    OrderLegToString = astrReturn.JoinFields(",")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.OrderLegToString"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToString
'' Description: Compile a delimited string to pass to Option Navigator for an order
'' Inputs:      Order, Refresh?
'' Returns:     Delimited String
''
'' Fields:      Genesis ID, Broker ID, Account Number, Trail Amount, TIF,
''              Group ID, Group Name, Order Status, Number Legs, Leg 1,
''              Leg 2, Leg 3, Leg 4, Refresh?, Order Date, Previous ID,
''              Message, Access Order ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderToString(ByVal Order As cPtOrder, ByVal bRefresh As Boolean) As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Order information to pass along
    Dim astrLeg As cGdArray             ' Leg information
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    Set astrLeg = New cGdArray
    astrLeg.Create eGDARRAY_Strings, 7
    
    astrReturn.Add Order.GenesisOrderID
    astrReturn.Add Order.BrokerID
    astrReturn.Add g.Broker.AccountNumberForID(Order.AccountID)
    astrReturn.Add Order.TrailAmount
    If Order.Expiration < 0 Then
        astrReturn.Add "DAY"
    ElseIf Order.Expiration = 0& Then
        astrReturn.Add "GTC"
    Else
        astrReturn.Add Format(Order.Expiration, "YYYYMMDD")
    End If
    astrReturn.Add Str(Order.GroupID)
    astrReturn.Add Order.GroupName
    astrReturn.Add Str(Order.Status)
    
    astrLeg(0) = Str(Abs(CLng(Order.OrderLegs(1).IsBuy)))
    astrLeg(1) = Str(Order.Quantity * Order.OrderLegs(1).Multiplier)
    astrLeg(2) = Order.OrderLegs(1).Symbol
    astrLeg(3) = Str(Order.StopPrice)
    astrLeg(4) = Str(Order.LimitPrice)
    astrLeg(5) = Str(Order.OrderType)
    astrLeg(6) = Str(Abs(CLng(Order.OrderLegs(1).IsEntry)))
    
    astrReturn.Add "1"
    astrReturn.Add astrLeg.JoinFields(",")
    astrReturn.Add ""
    astrReturn.Add ""
    astrReturn.Add ""
    
    astrReturn.Add Str(Abs(CLng(bRefresh)))
    astrReturn.Add Str(Order.OrderDate)
    astrReturn.Add Order.PreviousBrokerID
    astrReturn.Add Order.Message
    astrReturn.Add Str(Order.OrderID)
    
    OrderToString = astrReturn.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.OrderToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToSpreadString
'' Description: Compile a delimited string to pass to Option Navigator for a spread order
'' Inputs:      Order, Refresh?
'' Returns:     Delimited String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderToSpreadString(ByVal Order As cPtOrder, ByVal bRefresh As Boolean) As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of fields to join together to return
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings, 19 + Order.OrderLegs.Count
    
    astrReturn(0) = "GenesisOrderID=" & Order.GenesisOrderID
    If Len(Order.BrokerID) > 0 Then
        astrReturn(1) = "BrokerOrderID=" & Order.BrokerID
    End If
    If Order.OrderID > 0 Then
        astrReturn(2) = "AccessOrderID=" & Str(Order.OrderID)
    End If
    If Len(Order.PreviousBrokerID) > 0 Then
        astrReturn(3) = "PreviousBrokerID=" & Order.PreviousBrokerID
    End If
    astrReturn(4) = "AccountNumber=" & g.Broker.AccountNumberForID(Order.AccountID)
    astrReturn(5) = "DebitCredit=" & Str(Order.DebitCredit)
    astrReturn(6) = "Quantity=" & Str(Order.Quantity)
    astrReturn(7) = "OrderType=" & Str(Order.OrderType)
    
    ' TLB/DAJ: should these 0's be kNullData in the future?
    If Order.StopPrice > 0 Then
        astrReturn(8) = "StopPrice=" & Str(Order.StopPrice)
    End If
    If Order.LimitPrice > 0 Then
        astrReturn(9) = "LimitPrice=" & Str(Order.LimitPrice)
    End If
    
    If Order.Expiration < 0 Then
        astrReturn(10) = "TimeInForce=DAY"
    ElseIf Order.Expiration = 0& Then
        astrReturn(10) = "TimeInForce=GTC"
    Else
        astrReturn(10) = "TimeInForce=" & Format(Order.Expiration, "YYYYMMDD")
    End If
    astrReturn(11) = "UnderlyingSymbol=" & Order.UnderlyingSymbol
    If Order.GroupID > 0 Then
        astrReturn(12) = "GroupID=" & Str(Order.GroupID)
    End If
    If Len(Order.GroupName) > 0 Then
        astrReturn(13) = "GroupName=" & Order.GroupName
    End If
    astrReturn(14) = "Status=" & Str(Order.Status)
    astrReturn(15) = "OrderDate=" & Str(Order.OrderDate)
    If Len(Order.Message) > 0 Then
        astrReturn(16) = "Message=" & Order.Message
    End If
    If bRefresh Then astrReturn(17) = "Refresh=1" Else astrReturn(17) = "Refresh=0"
    astrReturn(18) = "NumberOfLegs=" & Str(Order.OrderLegs.Count)
    
    For lIndex = 1 To Order.OrderLegs.Count
        astrReturn(lIndex + 18) = "Leg" & Str(lIndex) & "=" & OrderLegToString(Order.OrderLegs(lIndex))
    Next lIndex
    
    OrderToSpreadString = astrReturn.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.OrderToSpreadString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderLegFromString
'' Description: Fill the order leg object from the Option Navigator string
'' Inputs:      Delimited String
'' Returns:     Order Leg
''
'' Fields:      Buy/Sell, Quantity, Symbol, Stop Price, Limit Price, Order Type,
''              Entry/Exit, Open, High, Low, Last, Bid, Ask, Strike Price,
''              Month, Year, Put/Call, Underlying Symbol, Expiration Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderLegFromString(ByVal strOrderLeg As String) As cOrderLeg
On Error GoTo ErrSection:

    Dim OrderLeg As cOrderLeg           ' Order Leg to return from the function
    Dim astrOrderLeg As cGdArray        ' Order Leg information to pass along
    Dim Bars As cGdBars                 ' Bars for the data
    Dim dStrike As Double               ' Strike price for the option
    Dim lMonth As Long                  ' Expiration Month for the option
    Dim lYear As Long                   ' Expiration Year for the option
    Dim lIsCall As Long                 ' Call or Put?
    Dim lUnderSymbolID As Long          ' Underlying symbol ID
    Dim lExpirationDate As Long         ' Expiration date for the symbol
    
    Set OrderLeg = New cOrderLeg
    Set astrOrderLeg = New cGdArray
    astrOrderLeg.SplitFields strOrderLeg, ","
    
    OrderLeg.IsBuy = (CLng(Val(astrOrderLeg(0))) <> 0)
    OrderLeg.Multiplier = CLng(Val(astrOrderLeg(1)))
    OrderLeg.SymbolOrSymbolID = astrOrderLeg(2)
    'OrderLeg.StopPrice = Val(astrOrderLeg(3))
    'OrderLeg.LimitPrice = Val(astrOrderLeg(4))
    'OrderLeg.OrderType = CLng(Val(astrOrderLeg(5)))
    OrderLeg.IsEntry = (CLng(Val(astrOrderLeg(6))) <> 0)
    
    ' DAJ 03/06/2009: Option Navigator is going to start passing data for
    ' the symbol along with the leg information, so we need to grab it and
    ' send it along to the data manager here...
    If astrOrderLeg.Size > 7 Then
        Set Bars = New cGdBars
        
        Bars.Size = 1
        SetBarProperties Bars, OrderLeg.Symbol
        
        Bars.ArrayMask = eBARS_EodBidAsk
        
        Bars(eBARS_DateTime, 0) = Bars.SessionDateForTime(CurrentTime(Bars.Prop(eBARS_ExchangeTimeZoneInf), Bars.Prop(eBARS_Symbol)), False)
        Bars(eBARS_Open, 0) = Val(astrOrderLeg(7))
        Bars(eBARS_High, 0) = Val(astrOrderLeg(8))
        Bars(eBARS_Low, 0) = Val(astrOrderLeg(9))
        Bars(eBARS_Close, 0) = Val(astrOrderLeg(10))
        Bars(eBARS_Bid, 0) = Val(astrOrderLeg(11))
        Bars(eBARS_Ask, 0) = Val(astrOrderLeg(12))
        
        dStrike = Val(astrOrderLeg(13))
        lMonth = CLng(Val(astrOrderLeg(14)))
        lYear = CLng(Val(astrOrderLeg(15)))
        lIsCall = CLng(Val(astrOrderLeg(16)))
        lUnderSymbolID = GetSymbolID(astrOrderLeg(17))
        
        DM_PutOptionSnap lUnderSymbolID, "", lMonth, lYear, dStrike, lIsCall, Bars
        g.RealTime.AddOptNavHistSymbol OrderLeg.Symbol
        
        lExpirationDate = CLng(Val(astrOrderLeg(18)))
        If lExpirationDate > 0 Then
            g.Broker.SetExpirationDateForSymbol OrderLeg.SymbolOrSymbolID, lExpirationDate
        End If
    End If
    
    Set OrderLegFromString = OrderLeg

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.OrderLegFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderLegFromSpreadString
'' Description: Fill the order leg object from the Option Navigator string
'' Inputs:      Delimited String
'' Returns:     Order Leg
''
'' Fields:      Buy/Sell, Quantity, Symbol, Entry/Exit, Expiration Date, Lot Size
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderLegFromSpreadString(ByVal strOrderLeg As String) As cOrderLeg
On Error GoTo ErrSection:

    Dim OrderLeg As cOrderLeg           ' Order Leg to return from the function
    Dim astrOrderLeg As cGdArray        ' Order Leg information to pass along
    Dim lExpirationDate As Long         ' Expiration date for the symbol
    Dim dLotSize As Double              ' Lot size
    
    Set OrderLeg = New cOrderLeg
    Set astrOrderLeg = New cGdArray
    astrOrderLeg.SplitFields strOrderLeg, ","
    
    OrderLeg.IsBuy = (CLng(Val(astrOrderLeg(0))) <> 0)
    OrderLeg.Multiplier = CLng(Val(astrOrderLeg(1)))
    OrderLeg.SymbolOrSymbolID = astrOrderLeg(2)
    OrderLeg.IsEntry = (CLng(Val(astrOrderLeg(3))) <> 0)
        
    lExpirationDate = CLng(Val(astrOrderLeg(4)))
    If lExpirationDate > 0 Then
        g.Broker.SetExpirationDateForSymbol OrderLeg.SymbolOrSymbolID, lExpirationDate
    End If
    
    dLotSize = Val(astrOrderLeg(5))
    
    ' Don't save the information unless we got both values from Option Navigator...
    If (lExpirationDate > 0) And (dLotSize > 0) Then
        g.Broker.SaveSymbolInfo OrderLeg.SymbolOrSymbolID, lExpirationDate, dLotSize
    End If
    
    Set OrderLegFromSpreadString = OrderLeg

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.OrderLegFromSpreadString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromString
'' Description: Fill the order object from the Option Navigator string
'' Inputs:      Delimited String, Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrderFromString(ByVal strOrder As String, Order As cPtOrder)
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Order information to pass along
    Dim lNumLegs As Long                ' Number of legs
    Dim lIndex As Long                  ' Index into a for loop
    Dim OrderLeg As cOrderLeg           ' Order leg object
    Dim astrLeg As cGdArray             ' Leg information broken out
    Dim lExpirationDate As Long         ' Expiration date for the symbol
    
    Set astrOrder = New cGdArray
    astrOrder.SplitFields strOrder, vbTab
    
    If (Len(astrOrder(0)) > 0) And (astrOrder(0) <> Str(Order.OrderID)) Then
        Order.GenesisOrderID = astrOrder(0)
    End If
    If Len(astrOrder(1)) > 0 Then Order.BrokerID = astrOrder(1)
    Order.AccountID = g.Broker.AccountIDForNumber(astrOrder(2))
    Order.TrailAmount = Val(astrOrder(3))
    If UCase(astrOrder(4)) = "DAY" Then
        Order.Expiration = -1&
    ElseIf UCase(astrOrder(4)) = "GTC" Then
        Order.Expiration = 0&
    Else
        Order.Expiration = JulFromLong(CLng(Val(astrOrder(4))))
    End If
    Order.GroupID = CLng(Val(astrOrder(5)))
    Order.GroupName = astrOrder(6)

    Order.OrderLegs.Clear
    lNumLegs = CLng(Val(astrOrder(7)))
    
    Set OrderLeg = New cOrderLeg
    Set astrLeg = New cGdArray
    astrLeg.SplitFields astrOrder(8), ","
    
    OrderLeg.LegNumber = 1
    OrderLeg.IsBuy = (astrLeg(0) = "1")
    OrderLeg.Multiplier = 1
    Order.Quantity = Str(astrLeg(1))
    OrderLeg.SymbolOrSymbolID = astrLeg(2)
    Order.StopPrice = Val(astrLeg(3))
    Order.LimitPrice = Val(astrLeg(4))
    Order.OrderType = CLng(Val(astrLeg(5)))
    OrderLeg.IsEntry = Str(astrLeg(6))
    Order.UnderlyingSymbolOrSymbolID = CLng(Val(astrLeg(17)))
    
    lExpirationDate = CLng(Val(astrLeg(18)))
    If lExpirationDate > 0 Then
        g.Broker.SetExpirationDateForSymbol OrderLeg.SymbolOrSymbolID, lExpirationDate
    End If
    
    Order.OrderLegs.Add OrderLeg
    
    Order.OptionNavImageFile = astrOrder(7 + lNumLegs + 1)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.OrderFromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromSpreadString
'' Description: Fill the order object from the Option Navigator string
'' Inputs:      Delimited String, Order
'' Returns:     None
''
'' Fields:      Genesis ID, Broker ID, Account, Debit/Credit, Quantity, Stop Price,
''              Limit Price, Order Type, TIF, Group ID, Group Name, Underlying Symbol,
''              Number Legs, Leg1, Leg2, Leg3, Leg4
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrderFromSpreadString(ByVal strOrder As String, Order As cPtOrder)
On Error GoTo ErrSection:

    Dim Fields As cGdTree               ' Dictionary of fields
    
    Set Fields = New cGdTree
    Fields.FromKeyValueString strOrder
    
    OrderFromSpreadTree Fields, Order
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.OrderFromSpreadString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromSpreadTree
'' Description: Fill the order object from the Option Navigator tree
'' Inputs:      Delimited String, Order
'' Returns:     None
''
'' Fields:      Genesis ID, Broker ID, Account, Debit/Credit, Quantity, Stop Price,
''              Limit Price, Order Type, TIF, Group ID, Group Name, Underlying Symbol,
''              Number Legs, Leg1, Leg2, Leg3, Leg4
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrderFromSpreadTree(SpreadTree As cGdTree, Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lNumLegs As Long                ' Number of legs
    Dim lIndex As Long                  ' Index into a for loop
    Dim OrderLeg As cOrderLeg           ' Order leg object
    Dim strKey As String                ' Key in the fields collection
    
    If SpreadTree.Exists("GenesisOrderID") Then
        Order.GenesisOrderID = SpreadTree("GenesisOrderID")
    End If
    If SpreadTree.Exists("BrokerOrderID") Then
        Order.BrokerID = SpreadTree("BrokerOrderID")
    End If
    If SpreadTree.Exists("AccountNumber") Then
        Order.AccountID = g.Broker.AccountIDForNumber(SpreadTree("AccountNumber"))
    End If
    If SpreadTree.Exists("DebitCredit") Then
        Order.DebitCredit = CLng(Val(SpreadTree("DebitCredit")))
    End If
    If SpreadTree.Exists("Quantity") Then
        Order.Quantity = CLng(Val(SpreadTree("Quantity")))
    End If
    If SpreadTree.Exists("StopPrice") Then
        Order.StopPrice = Val(SpreadTree("StopPrice"))
    End If
    If SpreadTree.Exists("LimitPrice") Then
        Order.LimitPrice = Val(SpreadTree("LimitPrice"))
    End If
    If SpreadTree.Exists("OrderType") Then
        Order.OrderType = CLng(Val(SpreadTree("OrderType")))
    End If
    If SpreadTree.Exists("TimeInForce") Then
        If UCase(SpreadTree("TimeInForce")) = "DAY" Then
            Order.Expiration = -1&
        ElseIf UCase(SpreadTree("TimeInForce")) = "GTC" Then
            Order.Expiration = 0&
        Else
            Order.Expiration = JulFromLong(CLng(Val(SpreadTree("TimeInForce"))))
        End If
    End If
    If SpreadTree.Exists("GroupID") Then
        Order.GroupID = CLng(Val(SpreadTree("GroupID")))
    End If
    If SpreadTree.Exists("GroupName") Then
        Order.GroupName = SpreadTree("GroupName")
    End If
    If SpreadTree.Exists("UnderlyingSymbol") Then
        Order.UnderlyingSymbolOrSymbolID = SpreadTree("UnderlyingSymbol")
    End If
    If SpreadTree.Exists("ImageFilename") Then
        Order.OptionNavImageFile = SpreadTree("ImageFilename")
    End If
    If SpreadTree.Exists("NumberOfLegs") Then
        lNumLegs = CLng(Val(SpreadTree("NumberOfLegs")))
        
        Order.OrderLegs.Clear
        For lIndex = 1 To lNumLegs
            strKey = "Leg" & Str(lIndex)
            If SpreadTree.Exists(strKey) Then
                Set OrderLeg = OrderLegFromSpreadString(SpreadTree(strKey))
                OrderLeg.LegNumber = lIndex
                
                Order.OrderLegs.Add OrderLeg
            End If
        Next lIndex
    End If
    
    Order.BuildSpreadSymbol
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mOptNav.OrderFromSpreadTree"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextGroupID
'' Description: Determine the next available group ID
'' Inputs:      None
'' Returns:     Next Group ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NextGroupID() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database

    lReturn = 1&
    Set rs = g.dbPaper.OpenRecordset("SELECT MAX([GroupID]) AS MaxGroupID FROM [tblOrders];", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        If IsNull(rs!MaxGroupID) = False Then
            lReturn = rs!MaxGroupID + 1
        End If
    End If
    
    NextGroupID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mOptNav.NextGroupID"
    
End Function
