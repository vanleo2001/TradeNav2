VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSimTradeBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cSimTradeBase.cls
'' Description: Base class object for simulated trading
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/24/2011   DAJ         Fix for loading snapshot orders
'' 06/29/2011   DAJ         Include today's closed orders for SimBroker snapshot refresh
'' 06/30/2011   DAJ         Fix for calculating session date for fills
'' 08/15/2011   DAJ         Added GetPositionForSymbol stuff
'' 08/15/2011   DAJ         On an order refresh, set the order type correctly
'' 10/24/2011   DAJ         If receive order/fill/position for unrecognized account, create it
'' 10/27/2011   DAJ         Use the status date to calc session date for orders instead of date sent
'' 12/09/2011   DAJ         Expiration dates for positions
'' 02/14/2012   DAJ         Added multi-leg order support
'' 06/07/2012   DAJ         Set the leg number in OrderFromInfo
'' 08/28/2012   DAJ         CarriedFillFromInfo, First Field Begin/End
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 09/04/2014   DAJ         Pulled SimTrade conversions out of trade objects
'' 12/10/2014   DAJ         Utilize new DateIsSnapshot routines
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccountsFromDb
'' Description: Get simulated trading accounts from the database
'' Inputs:      Broker
'' Returns:     Array of accounts
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAccountsFromDb(ByVal nBroker As eTT_AccountType) As cGdArray
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim astrReturn As cGdArray          ' Return array

    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountType]=" & Str(nBroker) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        astrReturn.Add Str(rs!AccountNumber)
        rs.MoveNext
    Loop
    
    Set GetAccountsFromDb = astrReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.GetAccountsFromDb"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Request accounts from the server
'' Inputs:      Broker Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAccounts(ByVal Broker As cBroker)
On Error GoTo ErrSection:

    Dim astrAccounts As cGdArray        ' Array of accounts from the database
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrAccounts = GetAccountsFromDb(Broker.Broker)
    
    HandleMessage Broker, eGDSimTradeMessageType_Account, "BEGIN"
    For lIndex = 0 To astrAccounts.Size - 1
        HandleMessage Broker, eGDSimTradeMessageType_Account, astrAccounts(lIndex)
    Next lIndex
    HandleMessage Broker, eGDSimTradeMessageType_Account, "END"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeBase.GetAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSnapshotOrdersFromDb
'' Description: Get simulated trading snapshot orders from the database
'' Inputs:      Broker, Account
'' Returns:     Array of orders
''
'' Order:       Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date, At ID, Session Date, Exchange ID, Status Date,
''              Date Sent, Date Exch Recd, Message
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSnapshotOrdersFromDb(ByVal nBroker As eTT_AccountType, Optional ByVal strAccount As String = "") As cGdArray
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim astrReturn As cGdArray          ' Return array
    Dim astrOrder As cGdArray           ' Array of information to join together
    Dim Order As cPtOrder               ' Order object
    Dim lAccountID As Long              ' Account ID
    Dim lLDD As Long                    ' Last daily download
    Dim bInclude As Boolean             ' Include the order in the refresh?

    lLDD = LastDailyDownload
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    Set astrOrder = New cGdArray
    astrOrder.Create eGDARRAY_Strings, 21
    
    If Len(strAccount) = 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT tblOrders.* " & _
                "FROM tblOrders INNER JOIN tblAccounts ON tblOrders.AccountID=tblAccounts.AccountID " & _
                "WHERE tblAccounts.AccountType=" & Str(nBroker) & " AND tblOrders.IsSnapshot=-1;", dbOpenDynaset)
    Else
        lAccountID = g.Broker.AccountIDForNumber(strAccount)
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM tblOrders " & _
                "WHERE tblOrders.AccountID=" & Str(lAccountID) & " AND tblOrders.IsSnapshot=-1;", dbOpenDynaset)
    End If
        
    Do While Not rs.EOF
        If nBroker = eTT_AccountType_SimBroker Then
            bInclude = (IsOpenOrder(rs!Status) = False) And (rs!SessionDate > lLDD)
        Else
            bInclude = (IsOpenOrder(rs!Status) = True) Or (rs!SessionDate > lLDD)
        End If
        
        If bInclude Then
            Set Order = New cPtOrder
            If Order.Load(rs!OrderID, rs) Then
                If nBroker = eTT_AccountType_SimStream Then
                    ' DAJ 09/23/2009: If Trade Navigator is starting up and we have working
                    ' quick fill orders, Trade Navigator must have gone down ungracefully
                    ' the last time -- mark Park Pending orders as Parked and all other working
                    ' orders that had been "sent" as Cancelled...
                    If g.bStarting Then
                        If Order.Status = eTT_OrderStatus_ParkPending Then
                            g.Broker.BrokerDebug nBroker, "'" & Order.OrderText(True, True, True) & "' set to Parked because it was left Park Pending"
                            Order.Status = eTT_OrderStatus_Parked
                            Order.StatusDate = CurrentTime(TimeZone(Order.Symbol), Order.Symbol, True)
                        ElseIf IsOpenOrder(Order.Status) And HasBeenSent(Order.Status) Then
                            g.Broker.BrokerDebug nBroker, "'" & Order.OrderText(True, True, True) & "' set to Cancelled because it was left working"
                            Order.Status = eTT_OrderStatus_Cancelled
                            Order.StatusDate = CurrentTime(TimeZone(Order.Symbol), Order.Symbol, True)
                        End If
                    
                    ' DAJ 12/03/2010: Otherwise, if we are doing a refresh and an order is in
                    ' a pending status, perform the appropriate action to it...
                    Else
                        Select Case Order.Status
                            Case eTT_OrderStatus_CancelPending
                                g.Broker.BrokerDebug nBroker, "'" & Order.OrderText(True, True, True) & "' set to Cancelled because it was left Cancel Pending"
                                Order.Status = eTT_OrderStatus_Cancelled
                                Order.StatusDate = CurrentTime(TimeZone(Order.Symbol), Order.Symbol, True)
                                
                            Case eTT_OrderStatus_ParkPending
                                g.Broker.BrokerDebug nBroker, "'" & Order.OrderText(True, True, True) & "' set to Parked because it was left Park Pending"
                                Order.Status = eTT_OrderStatus_Parked
                                Order.StatusDate = CurrentTime(TimeZone(Order.Symbol), Order.Symbol, True)
                        End Select
                    End If
                End If
                
#If 1 Then
                astrReturn.Add OrderToSimTradeString(Order, True)
#Else
                astrOrder(0) = Order.BrokerID
                astrOrder(1) = Order.GenesisOrderID
                astrOrder(2) = Order.PreviousBrokerID
                astrOrder(3) = g.Broker.AccountNumberForID(Order.AccountID)
                astrOrder(4) = Order.Symbol
                astrOrder(5) = Str(Order.OrderType)
                astrOrder(6) = Str(CLng(Order.Buy))
                astrOrder(7) = Str(Order.Quantity)
                astrOrder(8) = Str(Order.LimitPrice)
                astrOrder(9) = Str(Order.StopPrice)
                astrOrder(10) = "" ' Expiration
                astrOrder(11) = Str(Order.Status)
                astrOrder(12) = Str(Order.RemainingQuantity)
                astrOrder(13) = Str(Order.OrderDate)
                astrOrder(14) = ""
                astrOrder(15) = ""
                astrOrder(16) = Order.ExchangeID
                astrOrder(17) = Str(Order.StatusDate)
                astrOrder(18) = Str(Order.DateSent)
                astrOrder(19) = Str(Order.DateHostRecd)
                astrOrder(20) = Order.Message
            
                astrReturn.Add astrOrder.JoinFields(vbTab)
#End If
            End If
        End If
        
        rs.MoveNext
    Loop
    
    Set GetSnapshotOrdersFromDb = astrReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.GetSnapshotOrdersFromDb"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSnapshotFillsFromDb
'' Description: Get simulated trading snapshot fills from the database
'' Inputs:      Broker, Account
'' Returns:     Array of fills
''
'' Fill:        Broker ID, Genesis ID, Fill ID, Account, Symbol, Fill Date,
''              B/S, Fill Quantity, Fill Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSnapshotFillsFromDb(ByVal nBroker As eTT_AccountType, Optional ByVal strAccount = "") As cGdArray
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim astrReturn As cGdArray          ' Return array
    Dim Fill As cPtFill                 ' Fill object from the database
    Dim astrFill As cGdArray            ' Array of fill information
    Dim lAccountID As Long              ' Account ID
    Dim lLDD As Long                    ' Last daily download

    lLDD = LastDailyDownload
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    Set astrFill = New cGdArray
    astrFill.Create eGDARRAY_Strings, 9
    
    If Len(strAccount) = 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT tblFills.* " & _
                "FROM tblFills INNER JOIN tblAccounts ON tblFills.AccountID=tblAccounts.AccountID " & _
                "WHERE tblAccounts.AccountType=" & Str(nBroker) & " AND tblFills.SessionDate>" & Str(lLDD) & " AND tblFills.IsSpread=0;", dbOpenDynaset)
    Else
        lAccountID = g.Broker.AccountIDForNumber(strAccount)
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM tblFills " & _
                "WHERE tblFills.AccountID=" & Str(lAccountID) & " AND tblFills.SessionDate>" & Str(lLDD) & " AND tblFills.IsSpread=0;", dbOpenDynaset)
    End If
    
    Do While Not rs.EOF
        Set Fill = New cPtFill
        If Fill.Load(rs!FillID, rs) Then
            astrFill(0) = Fill.BrokerOrderID
            astrFill(1) = g.Broker.GenesisIdForOrderId(Fill.OrderID)
            astrFill(2) = Fill.BrokerID
            astrFill(3) = g.Broker.AccountNumberForID(Fill.AccountID)
            astrFill(4) = Fill.Symbol
            astrFill(5) = Str(Fill.FillDate)
            astrFill(6) = Str(CLng(Fill.Buy))
            astrFill(7) = Str(Fill.Quantity)
            astrFill(8) = Str(Fill.Price)
        
            astrReturn.Add astrFill.JoinFields(vbTab)
        End If
        
        rs.MoveNext
    Loop
    
    Set GetSnapshotFillsFromDb = astrReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.GetSnapshotFillsFromDb"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFills
'' Description: Request fills from the server
'' Inputs:      Broker Object, Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetFills(ByVal Broker As cBroker, Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim astrFills As cGdArray           ' Array of fills from the database
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrFills = GetSnapshotFillsFromDb(Broker.Broker, strAccount)
    
    HandleMessage Broker, eGDSimTradeMessageType_RefreshFill, "BEGIN"
    For lIndex = 0 To astrFills.Size - 1
        HandleMessage Broker, eGDSimTradeMessageType_RefreshFill, astrFills(lIndex)
    Next lIndex
    HandleMessage Broker, eGDSimTradeMessageType_RefreshFill, "END"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeBase.GetFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositionsFromDb
'' Description: Get simulated trading positions from the database
'' Inputs:      Broker, Account
'' Returns:     Array of fills
''
'' Position:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetPositionsFromDb(ByVal nBroker As eTT_AccountType, Optional ByVal strAccount = "") As cGdArray
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim astrReturn As cGdArray          ' Return array
    Dim astrPosition As cGdArray        ' Array of position information
    Dim lAccountID As Long              ' Account ID
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    Set astrPosition = New cGdArray
    astrPosition.Create eGDARRAY_Strings, 5
    
    If Len(strAccount) = 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT tblAccountPositions.* " & _
                "FROM tblAccountPositions INNER JOIN tblAccounts ON tblAccountPositions.AccountID=tblAccounts.AccountID " & _
                "WHERE tblAccounts.AccountType=" & Str(nBroker) & " AND tblAccountPositions.AutoTradeItemID=-1;", dbOpenDynaset)
    Else
        lAccountID = g.Broker.AccountIDForNumber(strAccount)
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM tblAccountPositions " & _
                "WHERE tblAccountPositions.AccountID=" & Str(lAccountID) & " AND tblAccountPositions.AutoTradeItemID=-1;", dbOpenDynaset)
    End If
    
    Do While Not rs.EOF
        astrPosition(0) = g.Broker.AccountNumberForID(rs!AccountID)
        astrPosition(1) = rs!Symbol
        astrPosition(2) = Str(rs!CurrentPositionSnapshot)
        astrPosition(3) = Str(rs!AverageEntrySnapshot)
        astrPosition(4) = Str(rs!CurrentPosition)
    
        astrReturn.Add astrPosition.JoinFields(vbTab)
        rs.MoveNext
    Loop
    
    Set GetPositionsFromDb = astrReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.GetPositionsFromDb"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositions
'' Description: Request positions from the server
'' Inputs:      Account
'' Returns:     None
''
'' Fields:      Account, Symbol, Current Position, Average Entry, Carried Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPositions(ByVal Broker As cBroker, Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim astrPositions As cGdArray       ' Array of positions from the database
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrPositions = GetPositionsFromDb(Broker.Broker, strAccount)
    
    HandleMessage Broker, eGDSimTradeMessageType_RefreshPosition, "BEGIN"
    For lIndex = 0 To astrPositions.Size - 1
        HandleMessage Broker, eGDSimTradeMessageType_RefreshPosition, astrPositions(lIndex)
    Next lIndex
    HandleMessage Broker, eGDSimTradeMessageType_RefreshPosition, "END"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeBase.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositionForSymbolFromDb
'' Description: Get the position for the given symbol/account from the database
'' Inputs:      Account, Symbol
'' Returns:     Position Information
''
'' Fields:      Account, Symbol, Current Position, Average Entry, Carried Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetPositionForSymbolFromDb(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant) As String
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID
    Dim lSymbolID As Long               ' Symbol ID
    Dim strSymbol As String             ' Symbol
    Dim rs As Recordset                 ' Recordset into the database
    Dim astrPosition As cGdArray        ' Array of position information
    Dim strReturn As String             ' Return value for the function
    
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    
    If lSymbolID = 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                "WHERE AccountID=" & Str(lAccountID) & " AND Symbol='" & strSymbol & "' AND AutoTradeItemID=-1;", dbOpenDynaset)
    Else
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                "WHERE AccountID=" & Str(lAccountID) & " AND SymbolID=" & Str(lSymbolID) & " AND AutoTradeItemID=-1;", dbOpenDynaset)
    End If
    
    If Not (rs.BOF And rs.EOF) Then
        Set astrPosition = New cGdArray
        astrPosition.Create eGDARRAY_Strings, 5
        
        astrPosition(0) = g.Broker.AccountNumberForID(rs!AccountID)
        astrPosition(1) = rs!Symbol
        astrPosition(2) = Str(rs!CurrentPositionSnapshot)
        astrPosition(3) = Str(rs!AverageEntrySnapshot)
        astrPosition(4) = Str(rs!CurrentPosition)
        
        strReturn = astrPosition.JoinFields(vbTab)
    End If
    
    GetPositionForSymbolFromDb = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.GetPositionForSymbolFromDb"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositionForSymbol
'' Description: Get the position for the given symbol/account and send
'' Inputs:      Broker, Account, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPositionForSymbol(ByVal Broker As cBroker, ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    Dim strPosition As String           ' Position information
    
    strPosition = GetPositionForSymbolFromDb(vAccountNumberOrID, vSymbolOrSymbolID)
    If Len(strPosition) > 0 Then
        HandleMessage Broker, eGDSimTradeMessageType_Position, strPosition
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeBase.GetPositionForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextGenesisID
'' Description: Determine the next unique Genesis ID for the given account
'' Inputs:      Account Number
'' Returns:     Next Unique Genesis ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextGenesisID(ByVal strIniFile As String) As String
On Error GoTo ErrSection:

    Dim lDate As Long                   ' Date from the ini file
    Dim lCounter As Long                ' Counter from the ini file
    
    lDate = GetIniFileProperty("Date", 0&, "ID", strIniFile)
    lCounter = GetIniFileProperty("Counter", 0&, "ID", strIniFile)
    
    If lDate <> Date Then
        lDate = Date
        lCounter = 0&
    Else
        lCounter = lCounter + 1
    End If

    SetIniFileProperty "Date", lDate, "ID", strIniFile
    SetIniFileProperty "Counter", lCounter, "ID", strIniFile
    
    NextGenesisID = Format(lDate, "00000") & Format(lCounter, "0000")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.NextGenesisID"
    
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

    StartingGenesisIdForOptNav = Format(Date, "00000") & "7000"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.StartingGenesisIdForOptNav"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming message
'' Inputs:      Broker Object, Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal Broker As cBroker, ByVal nType As eGDSimTradeMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strPosition As String           ' Position information

    Broker.DumpDebug vbTab & Broker.BrokerName & " Message Received(" & MessageTypeToString(nType) & "): " & strMessage
    
    Select Case nType
        Case eGDSimTradeMessageType_Account
            Broker.AccountReceived strMessage
            
        Case eGDSimTradeMessageType_Order
            Broker.OrderReceived strMessage, False
            
        Case eGDSimTradeMessageType_RefreshOrder
            Broker.OrderReceived strMessage, True
            
        Case eGDSimTradeMessageType_Fill
            Broker.FillReceived strMessage, False
            strPosition = GetPositionForFill(Broker.Broker, strMessage)
            If Len(strPosition) > 0 Then
                HandleMessage Broker, eGDSimTradeMessageType_Position, strPosition
            End If
            
        Case eGDSimTradeMessageType_RefreshFill
            Broker.FillReceived strMessage, True
            
        Case eGDSimTradeMessageType_Position
            Broker.PositionReceived strMessage, False
    
        Case eGDSimTradeMessageType_RefreshPosition
            Broker.PositionReceived strMessage, True
            
        Case eGDSimTradeMessageType_SpreadFill
            Broker.HandleMessage eGDBrokerMessageType_SpreadFill, strMessage
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeBase.HandleMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountFromInfo
'' Description: Fill an account object from the given message
'' Inputs:      Broker, Message, Account
'' Returns:     True if OK, False otherwise
''
'' Fields:      Account Number
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountFromInfo(ByVal nBroker As eTT_AccountType, ByVal strMessage As String, Acct As cPtAccount) As Boolean
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
                .AccountType = nBroker
                .Broker = ""
                .ClosedProfit = 0#
                .Comms = 0#
                .CurrentBalance = 0#
                .FillRT = False
                .HoldOcoAtBroker = False
                .Name = .AccountNumber
                .Password = ""
                .SecTypeMask = 31
                .StartingBalance = 0#
                .StartingDate = Date
            End With
            
            bReturn = True
        End If
    End If
    
    AccountFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.AccountFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromInfo
'' Description: Fill an order object from the given message
'' Inputs:      Broker, Message, Order
'' Returns:     True if OK, False otherwise
''
'' Fields:      Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date, At ID, Session Date, Exchange ID, Status Date,
''              Date Sent, Date Exch Recd, Message
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderFromInfo(ByVal nBroker As eTT_AccountType, ByVal strMessage As String, Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrOrder As cGdArray           ' Order information split out from the message
    Dim strGenesisSymbol As String      ' Genesis symbol for the order
    Dim Bars As cGdBars                 ' Bars object
    Dim lExpiration As Long             ' Expiration
    Dim lAccountID As Long              ' Account ID
    Dim astrLegs As cGdArray            ' Array of leg information
    Dim OrderLeg As cOrderLeg           ' Order leg information
    Dim lQuantity As Long               ' Leg quantity
    Dim lIndex As Long                  ' Index into a for loop
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set astrOrder = New cGdArray
            astrOrder.SplitFields strMessage, vbTab
            
            If (g.Broker.AccountExists(astrOrder(3)) = False) And (Len(astrOrder(3)) > 0) Then
                lAccountID = mTradeTracker.CreateAccountFromNumber(astrOrder(3), nBroker)
            Else
                lAccountID = g.Broker.AccountIDForNumber(astrOrder(3))
            End If
            
            Set Order = New cPtOrder
            With Order
                .AccountID = lAccountID
                .BrokerID = astrOrder(0)
                .Expiration = CLng(Val(astrOrder(10)))
                .GenesisOrderID = astrOrder(1)
                .LimitPrice = Val(astrOrder(8))
                .Message = astrOrder(20)
                .OrderType = CLng(Val(astrOrder(5)))
                .Quantity = CLng(Val(astrOrder(7)))
                .Status = CLng(Val(astrOrder(11)))
                .StatusDate = Val(astrOrder(17))
                .StopPrice = Val(astrOrder(9))
                
                If Len(astrOrder(21)) = 0 Then
                    strGenesisSymbol = astrOrder(4)
                    .Buy = (astrOrder(6) <> 0)
                    .SymbolOrSymbolID = strGenesisSymbol
                Else
                    Set astrLegs = New cGdArray
                    astrLegs.SplitFields astrOrder(21), "|"
                    For lIndex = 0 To astrLegs.Size - 1
                        lQuantity = CLng(Val(Parse(astrLegs(lIndex), ";", 1)))
                        strGenesisSymbol = Parse(astrLegs(lIndex), ";", 2)
                        
                        Set OrderLeg = New cOrderLeg
                        OrderLeg.LegNumber = lIndex + 1
                        OrderLeg.IsBuy = (lQuantity > 0)
                        OrderLeg.Multiplier = Abs(lQuantity)
                        OrderLeg.SymbolOrSymbolID = strGenesisSymbol
                        
                        .OrderLegs.Add OrderLeg
                    Next lIndex
                End If
                
                If .Status <> eTT_OrderStatus_Error Then
                    Set Bars = New cGdBars
                    SetBarProperties Bars, strGenesisSymbol
                    
                    .ExchangeID = astrOrder(16)
                    .DateSent = Val(astrOrder(18))
                    .DateExchRecd = Val(astrOrder(19))
                    .SessionDate = Bars.SessionDateForTradeTime(ConvertTimeZone(.StatusDate, TimeZone(.Symbol), Bars.Prop(eBARS_ExchangeTimeZoneInf)))
                End If
            End With
            
            bReturn = True
        End If
    End If
    
    OrderFromInfo = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.OrderFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillFromInfo
'' Description: Fill a fill object from the given message
'' Inputs:      Broker, Message, Fill, Genesis Order ID, Carried?
'' Returns:     True if OK, False otherwise
''
'' Fields:      Broker ID, Genesis ID, Fill ID, Account, Symbol, Fill Date,
''              B/S, Fill Quantity, Fill Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillFromInfo(ByVal nBroker As eTT_AccountType, ByVal strMessage As String, Fill As cPtFill, Optional strGenesisOrderID As String = "", Optional ByVal bCarried As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrFill As cGdArray            ' Fill information
    Dim Bars As New cGdBars             ' Bars object
    Dim lAccountID As Long              ' Account ID
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set astrFill = New cGdArray
            astrFill.SplitFields strMessage, vbTab
            
            If (g.Broker.AccountExists(astrFill(3)) = False) And (Len(astrFill(3)) > 0) Then
                lAccountID = mTradeTracker.CreateAccountFromNumber(astrFill(3), nBroker)
            Else
                lAccountID = g.Broker.AccountIDForNumber(astrFill(3))
            End If
            
            Set Fill = New cPtFill
            With Fill
                .AccountID = g.Broker.AccountIDForNumber(astrFill(3))
                .SymbolOrSymbolID = astrFill(4)
                .BrokerID = astrFill(2)
                .BrokerOrderID = astrFill(0)
                .Buy = (astrFill(6) <> 0)
                .FillDate = Val(astrFill(5))
                .Price = Val(astrFill(8))
                .Quantity = CLng(Val(astrFill(7)))
                
                SetBarProperties Bars, .SymbolOrSymbolID
                .SessionDate = Bars.SessionDateForTradeTime(ConvertTimeZone(.FillDate, TimeZone(.Symbol), Bars.Prop(eBARS_ExchangeTimeZoneInf)))
            End With
            
            strGenesisOrderID = astrFill(1)
            
            bReturn = True
        End If
    End If
    
    FillFromInfo = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.FillFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CarriedFillFromInfo
'' Description: Fill a fill object from the given message
'' Inputs:      Broker, Message, Fill, Genesis Order ID
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CarriedFillFromInfo(ByVal nBroker As eTT_AccountType, ByVal strMessage As String, Fill As cPtFill, Optional strGenesisOrderID As String = "") As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            bReturn = FillFromInfo(nBroker, strMessage, Fill, strGenesisOrderID, True)
        End If
    End If
    
    CarriedFillFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.CarriedFillFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionFromInfo
'' Description: Fill a position object from the given message
'' Inputs:      Broker, Message, Position
'' Returns:     True if OK, False otherwise
''
'' Fields:      Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionFromInfo(ByVal nBroker As eTT_AccountType, ByVal strMessage As String, Position As cPtPosition) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrPosition As cGdArray        ' Position information split out from the message
    Dim lAccountID As Long              ' Account ID
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set astrPosition = New cGdArray
            astrPosition.SplitFields strMessage, vbTab
            
            If (g.Broker.AccountExists(astrPosition(0)) = False) And (Len(astrPosition(0)) > 0) Then
                lAccountID = mTradeTracker.CreateAccountFromNumber(astrPosition(0), nBroker)
            Else
                lAccountID = g.Broker.AccountIDForNumber(astrPosition(0))
            End If
            
            Set Position = New cPtPosition
            With Position
                .AccountID = g.Broker.AccountIDForNumber(astrPosition(0))
                .SymbolOrSymbolID = astrPosition(1)
                
                .CurrentPosition = CLng(Val(astrPosition(2)))
                .CarriedPosition = CLng(Val(astrPosition(4)))
            End With
            
            bReturn = True
        End If
    End If
    
    PositionFromInfo = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.PositionFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositionForFill
'' Description: Get position information for a fill message
'' Inputs:      Broker, Fill Message
'' Returns:     Position string
''
'' Fields:      Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetPositionForFill(ByVal nBroker As eTT_AccountType, ByVal strFillMessage As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim Fill As cPtFill                 ' Fill object
    
    If FillFromInfo(nBroker, strFillMessage, Fill) Then
        strReturn = GetPositionForSymbolFromDb(Fill.AccountID, Fill.SymbolOrSymbolID)
    End If
    
    GetPositionForFill = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.GetPositionForFill"
    
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

    TimeZone = "NY"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.TimeZone"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenExpiredPositions
'' Description: Flattent any expired positions
'' Inputs:      Broker Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FlattenExpiredPositions(Broker As cBroker)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim AcctPos As cAccountPosition     ' Account position object
    Dim FillSumms As cAccountPositions  ' Account positions
    
    Set FillSumms = Broker.BrokerInfo.FillSummaries
    If Not FillSumms Is Nothing Then
        For lIndex = 1 To FillSumms.Count
            Set AcctPos = FillSumms.Item(lIndex)
            If AcctPos.AutoTradeItemID > -1& Then
                FlattenExpiredPosition Broker, AcctPos.AccountID, AcctPos.Symbol, AcctPos.AutoTradeItemID, AcctPos.ExpirationDate
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSimTradeBase.FlattenExpiredPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenExpiredPosition
'' Description: Flattent the expired position for the given account and symbol
'' Inputs:      Broker Object, Account, Symbol, Auto Trade Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FlattenExpiredPosition(Broker As cBroker, ByVal vAccountNumberOrID As Variant, ByVal strSymbol As String, ByVal lAtID As Long, Optional ByVal lExpirationDate As Long = kNullData)
On Error GoTo ErrSection:

    Dim strSecType As String            ' Security type for the symbol
    Dim Bars As cGdBars                 ' Bars object
    Dim Order As cPtOrder               ' Order object
    Dim Fill As cPtFill                 ' Fill object
    Dim lPosition As Long               ' Current position
    Dim dFillDate As Double             ' Date for the fill
    Dim dFillPrice As Double            ' Price for the fill
    Dim strAccount As String            ' Account number

    If IsExpiredContract(strSymbol, lExpirationDate) Then
        strAccount = g.Broker.GetAccountNumber(vAccountNumberOrID)
        
        lPosition = Broker.BrokerInfo.CurrentPosition(strAccount, strSymbol, lAtID)
        If lPosition <> 0& Then
            Set Bars = New cGdBars
            strSecType = SecurityType(strSymbol, True)
            Select Case strSecType
                Case "F"
                    If DM_GetBars(Bars, strSymbol, "Daily") Then
                        If lExpirationDate = kNullData Then
                            lExpirationDate = Bars(eBARS_DateTime, Bars.Size - 1)
                        End If
                        dFillDate = CDbl(lExpirationDate) + (Bars.Prop(eBARS_DefaultEndTime) / 1440)
                        lExpirationDate = Bars.SessionDateForTradeTime(dFillDate)
                        dFillPrice = Bars(eBARS_Close, Bars.Size - 1)
                    Else
                        SetBarProperties Bars, strSymbol
                        If lExpirationDate = kNullData Then
                            lExpirationDate = JulFromLong(Bars.LastDayOfContractMonth)
                            Do While IsWeekday(lExpirationDate) = False
                                lExpirationDate = lExpirationDate - 1&
                            Loop
                        End If
                        dFillDate = CDbl(lExpirationDate) + (Bars.Prop(eBARS_DefaultEndTime) / 1440)
                        dFillPrice = Broker.BrokerInfo.FillSummaryForSymbol(strAccount, strSymbol, lAtID).AverageEntrySnapshot
                    End If
                    
                Case "SO"
                    SetBarProperties Bars, strSymbol
                    If lExpirationDate = kNullData Then
                        lExpirationDate = JulFromLong(CLng(Val(Parse(strSymbol, " ", 2))))
                    End If
                    dFillDate = CDbl(lExpirationDate) + (Bars.Prop(eBARS_DefaultEndTime) / 1440)
                    dFillPrice = 0#
                
                Case "FO"
                    SetBarProperties Bars, strSymbol
                    If lExpirationDate = kNullData Then
                        lExpirationDate = JulFromLong(Bars.LastDayOfContractMonth)
                        Do While IsWeekday(lExpirationDate) = False
                            lExpirationDate = lExpirationDate - 1&
                        Loop
                    End If
                    dFillDate = CDbl(lExpirationDate) + (Bars.Prop(eBARS_DefaultEndTime) / 1440)
                    dFillPrice = 0#
                
            End Select
            
            Broker.DumpDebug "Flattening Expired Position(" & strAccount & ", " & strSymbol & ", " & Str(lAtID) & "): " & Str(dFillPrice) & " at " & DateFormat(dFillDate, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
            
            Set Order = New cPtOrder
            Order.AccountID = g.Broker.GetAccountID(vAccountNumberOrID)
            Order.AutoTradeItemID = lAtID
            Order.GenesisOrderID = Broker.NextGenesisID(strAccount)
            Order.BrokerID = "DBO" & Order.GenesisOrderID
            Order.ExchangeID = "EBO" & Order.GenesisOrderID
            Order.Buy = (lPosition < 0)
            Order.Enter = False
            Order.OrderDate = dFillDate
            Order.OrderType = eTT_OrderType_Market
            Order.Quantity = Abs(lPosition)
            Order.SessionDate = lExpirationDate
            Order.Status = eTT_OrderStatus_Filled
            Order.StatusDate = dFillDate
            Order.SymbolOrSymbolID = strSymbol
            
            Set Fill = New cPtFill
            Fill.AccountID = g.Broker.GetAccountID(vAccountNumberOrID)
            Fill.AutoTradingItemID = lAtID
            Fill.BrokerID = "DFO" & Order.GenesisOrderID
            Fill.BrokerOrderID = Order.BrokerID
            Fill.Buy = (lPosition < 0)
            Fill.FillDate = dFillDate
            Fill.IsManual = False
            Fill.OrderID = 0&
            Fill.Price = dFillPrice
            Fill.Quantity = Abs(lPosition)
            Fill.SessionDate = lExpirationDate
            Fill.SymbolOrSymbolID = strSymbol
            Fill.IsSnapshot = g.Broker.DateIsSnapshotForFill(Fill)
        
            If Fill.IsSnapshot Then
                HandleMessage Broker, eGDSimTradeMessageType_Order, OrderToSimTradeString(Order)
                HandleMessage Broker, eGDSimTradeMessageType_Fill, FillToSimTradeString(Fill)
            Else
                g.Broker.RefreshManualFill Fill, , False
                GetPositionForSymbol Broker, vAccountNumberOrID, Fill.SymbolOrSymbolID
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSimTradeBase.FlattenExpiredPosition"
    
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
    
    ' Only limit orders are allowed for future options...
    If (InStr(strGenesisSymbol, " ") <> 0) And (InStr(strGenesisSymbol, "-") <> 0) Then
        strReturn = "001000000" & vbTab & "110"
        
    ' Allow all order types for all other security types...
    Else
        strReturn = "111111111" & vbTab & "111"
    End If
    
    SymbolInformation = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.SymbolInformation"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MessageTypeToString
'' Description: Convert a message type enumeration to a string
'' Inputs:      Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MessageTypeToString(ByVal nType As eGDSimTradeMessageTypes) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from the function

    Select Case nType
        Case eGDSimTradeMessageType_Account
            strReturn = "eGDSimTradeMessageType_Account"
        Case eGDSimTradeMessageType_Order
            strReturn = "eGDSimTradeMessageType_Order"
        Case eGDSimTradeMessageType_Fill
            strReturn = "eGDSimTradeMessageType_Fill"
        Case eGDSimTradeMessageType_Position
            strReturn = "eGDSimTradeMessageType_Position"
        Case eGDSimTradeMessageType_RefreshOrder
            strReturn = "eGDSimTradeMessageType_RefreshOrder"
        Case eGDSimTradeMessageType_RefreshFill
            strReturn = "eGDSimTradeMessageType_RefreshFill"
        Case eGDSimTradeMessageType_RefreshPosition
            strReturn = "eGDSimTradeMessageType_RefreshPosition"
        
        Case Else
            strReturn = Str(nType)
    End Select
    
    MessageTypeToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.MessageTypeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillToSimTradeString
'' Description: Convert the fill to a tab delimited string
'' Inputs:      Fill
'' Returns:     Tab Delimited String
''
'' Fill:        Broker ID, Genesis ID, Fill ID, Account, Symbol, Fill Date,
''              B/S, Fill Quantity, Fill Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillToSimTradeString(ByVal Fill As cPtFill) As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of order information to join together
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    astrReturn(0) = Fill.BrokerOrderID
    astrReturn(1) = ""
    astrReturn(2) = Fill.BrokerID
    astrReturn(3) = g.Broker.AccountNumberForID(Fill.AccountID)
    astrReturn(4) = Fill.Symbol
    astrReturn(5) = Str(Fill.FillDate)
    astrReturn(6) = Str(CLng(Fill.Buy))
    astrReturn(7) = Str(Fill.Quantity)
    astrReturn(8) = Str(Fill.Price)
    
    FillToSimTradeString = astrReturn.JoinFields(vbTab)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.FillToSimTradeString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToSimTradeString
'' Description: Convert the order into a broker info string
'' Inputs:      Order
'' Returns:     Broker Info String
''
'' Order:       Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date, At ID, Session Date, Exchange ID, Status Date,
''              Date Sent, Date Exch Recd, Message
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderToSimTradeString(ByVal Order As cPtOrder, Optional ByVal bForRefresh As Boolean = False) As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of order information to join together
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings, 22
    
    astrReturn(0) = Order.BrokerID
    astrReturn(1) = Order.GenesisOrderID
    astrReturn(3) = g.Broker.AccountNumberForID(Order.AccountID)
    astrReturn(5) = Str(Order.OrderType)
    astrReturn(7) = Str(Order.Quantity)
    astrReturn(8) = Str(Order.LimitPrice)
    astrReturn(9) = Str(Order.StopPrice)
    astrReturn(11) = Str(Order.Status)
    astrReturn(12) = Str(Order.RemainingQuantity)
    astrReturn(13) = Str(Order.OrderDate)
    astrReturn(16) = Order.ExchangeID
    astrReturn(17) = Str(Order.StatusDate)
    astrReturn(18) = Str(Order.DateSent)
    astrReturn(19) = Str(Order.DateExchRecd)
    astrReturn(20) = Order.Message
    
    ' 10/23/2012 DAJ: This had been down in the following block and was being set to a blank string
    ' if it was for a refresh.  Unfortunately, that was causing orders to turn from Day orders to
    ' GTC orders after a refresh was done for a simulated account.  By always setting this here, the
    ' TIF will stay in tact for the order...
    astrReturn(10) = Str(Order.Expiration)
    
    If bForRefresh Then
        astrReturn(2) = Order.PreviousBrokerID
        astrReturn(14) = ""
        astrReturn(15) = ""
    Else
        astrReturn(2) = ""
        astrReturn(14) = Str(Order.AutoTradeItemID)
        astrReturn(15) = Str(Order.SessionDate)
    End If
    
    If Order.NumberOfLegs = 1 Then
        astrReturn(4) = Order.Symbol
        astrReturn(6) = Str(CLng(Order.Buy))
        
        If Len(Order.SpreadSymbol) = 0 Then
            Order.BuildSpreadSymbol
        End If
        astrReturn(21) = Order.SpreadSymbol
    ElseIf Order.NumberOfLegs > 1 Then
        astrReturn(4) = ""
        astrReturn(6) = ""
        astrReturn(21) = Order.SpreadSymbol
    End If
    
    OrderToSimTradeString = astrReturn.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSimTradeBase.OrderToSimTradeString"
    
End Function
