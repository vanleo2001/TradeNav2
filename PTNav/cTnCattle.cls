VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTnCattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTnCattle.cls
'' Description: Application side functions for the Cattle DLL
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/07/2014   DAJ         Created
'' 03/13/2014   DAJ         Allow brokers to associate carried fill information
'' 03/19/2014   DAJ         Renamed Turnkey to Cattle
'' 03/20/2014   DAJ         Accept the VS Printer object as an object to fix Type Mismatch
'' 04/28/2014   DAJ         Added ability to get broker orders
'' 06/06/2014   DAJ         Added ability to create/edit a fill from Cattle stuff
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Property Get AccountKey(ByVal strBroker As String, ByVal strAccountNumber As String) As String
    AccountKey = strBroker & "|" & strAccountNumber
End Property

Private Property Get FillKey(ByVal strBroker As String, ByVal strBrokerFillID As String) As String
    FillKey = strBroker & "|" & strBrokerFillID
End Property

Private Property Get OrderKey(ByVal strBroker As String, ByVal strBrokerOrderID As String) As String
    OrderKey = strBroker & "|" & strBrokerOrderID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Picture16
'' Description: Grab the appropriate image from the appropriate image list
'' Inputs:      Picture name, Image List
'' Returns:     Image
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Picture16(ByVal strPicture$, Optional ByVal iImageList As Integer = 0) As Object
    Set Picture16 = mMain.Picture16(strPicture, iImageList)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PriceDisplay
'' Description: Get the formatted version of the given price
'' Inputs:      Price, Symbol
'' Returns:     Formatted Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PriceDisplay(ByVal dPrice As Double, ByVal strSymbol As String) As String
    PriceDisplay = mDataNav.PriceDisplay(dPrice, strSymbol)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasModule
'' Description: Determine if the user has the given module(s)
'' Inputs:      Module(s), Include Source Code?
'' Returns:     True if have module(s), False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasModule(ByVal strModules As String, Optional ByVal bIncludeSourceCode As Boolean = False) As Boolean
    HasModule = mMain.HasModule(strModules, bIncludeSourceCode)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddDialogMessage
'' Description: Add a dialog message to be shown with a timer
'' Inputs:      Message, Caption, Icon, Buttons
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddDialogMessage(ByVal strMessage As String, Optional ByVal strIcon As String = "", Optional ByVal strButtons As String = "", Optional ByVal strCaption As String = "")
    frmOnlineBroker.AddDialogMessage strMessage, strIcon, strButtons, strCaption
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Unloading
'' Description: Is Trade Navigator currently unloading?
'' Inputs:      None
'' Returns:     True if unloading, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Unloading() As Boolean
    Unloading = g.bUnloading
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FormatVersion
'' Description: Formatted version of the Trade Navigator version
'' Inputs:      Include Revision?, Include Date?
'' Returns:     Formatted version
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FormatVersion(Optional ByVal bIncludeRevision As Boolean = False, Optional ByVal bIncludeFileDate As Boolean = False) As String
    FormatVersion = mMain.FormatVersion(bIncludeRevision, bIncludeFileDate)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoPrintHeader
'' Description: Do the print header for Trade Navigator
'' Inputs:      Font Size, Printer Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoPrintHeader(Optional ByVal nFontSize& = 12, Optional ByVal vp As Object = Nothing)
    mMain.DoPrintHeader nFontSize, vp
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToolbarIcon
'' Description: Return the toolbar icon name for the given ID
'' Inputs:      ID
'' Returns:     Toolbar Icon name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToolbarIcon(ByVal strID As String) As String
    ToolbarIcon = mMain.ToolbarIcon(strID)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SecurityType
'' Description: Return the security type for the given symbol or bars
'' Inputs:      Symbol or Bars, Append O if option?
'' Returns:     Security Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SecurityType(BarsOrSymbol As Variant, Optional bIfOptionAppendO As Boolean = False) As String
    SecurityType = mDataNav.SecurityType(BarsOrSymbol, bIfOptionAppendO)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Profit
'' Description: Determine the profit (or loss) for the given information
'' Inputs:      Symbol, Price Difference, Quantity, Exit Date/Time (in exchange time),
''              Bars for Symbol, Bars for Base Forex, Account number
'' Returns:     None
''
'' Note:        If passing in base forex bars, pass in full history of daily
''              bars for the correct base forex symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Profit(ByVal strSymbol As String, ByVal dPriceDifference As Double, Optional ByVal lQuantity As Long = 1&, Optional ByVal dExchangeExitTime As Double = -99999#, Optional ByVal hSymbolBars As Long = 0, Optional ByVal hBaseForexBars As Long = 0, Optional ByVal vAccountNumberOrID As Variant = "") As Double
On Error GoTo ErrSection:

    Dim SymbolBars As cGdBars           ' Bars object for the symbol
    Dim BaseForexBars As cGdBars        ' Bars object for the base forex
    
    If hSymbolBars = 0 Then
        Set SymbolBars = Nothing
    Else
        Set SymbolBars = New cGdBars
        SymbolBars.CopyFromHandle hSymbolBars
    End If

    If hBaseForexBars = 0 Then
        Set BaseForexBars = Nothing
    Else
        Set BaseForexBars = New cGdBars
        BaseForexBars.CopyFromHandle hBaseForexBars
    End If

    Profit = g.Profit.Profit(strSymbol, dPriceDifference, lQuantity, dExchangeExitTime, SymbolBars, BaseForexBars, vAccountNumberOrID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.Profit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBrokerAccounts
'' Description: Get a list of broker accounts for associations
'' Inputs:      None
'' Returns:     Broker Accounts
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetBrokerAccounts() As cGdTree
On Error GoTo ErrSection:

    Dim ReturnAccounts As cGdTree       ' Collection of accounts to return
    Dim Accounts As cPtAccounts         ' Collection of accounts
    Dim BrokerAccounts As cGdTree       ' Collection of accounts for a broker
    Dim lIndex As Long                  ' Index into a for loop
    Dim strKey As String                ' Key into the collection
    Dim nBroker As eTT_AccountType      ' Broker for the broker view form
    Dim Account As cBrokerMessage       ' Account from the broker collection
    
    Set ReturnAccounts = New cGdTree
    Set Accounts = g.Broker.AllAccounts
    
    For lIndex = 1 To Accounts.Count
        strKey = AccountKey(Str(Accounts(lIndex).AccountType), Accounts(lIndex).AccountNumber)
        If ReturnAccounts.Exists(strKey) = False Then
            ReturnAccounts.Add GenesisAccountToTurnkey(Accounts(lIndex)), strKey
        End If
    Next lIndex
    
    If FormIsLoaded("frmBrokerView") Then
        nBroker = frmBrokerView.Broker
        'Set BrokerAccounts = frmBrokerView.Accounts
        Set BrokerAccounts = frmBrokerView.AllAccounts
        For lIndex = 1 To BrokerAccounts.Count
            Set Account = BrokerAccounts(lIndex)
            strKey = AccountKey(Str(nBroker), Account("Account"))
            If ReturnAccounts.Exists(strKey) = False Then
                ReturnAccounts.Add KeyValueAccountToTurnkey(Account, nBroker), strKey
            End If
        Next lIndex
    End If
    
    Set GetBrokerAccounts = ReturnAccounts

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.GetBrokerAccounts"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountForBrokerNumber
'' Description: Get a Turnkey account for the given broker account number
'' Inputs:      Account Number, Broker, Key Value?
'' Returns:     Turnkey Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountForBrokerNumber(ByVal strAccountNumber As String, ByVal nBroker As eTT_AccountType, ByVal bKeyValue As Boolean) As cBrokerMessage
On Error GoTo ErrSection:

    Dim ReturnAccount As cBrokerMessage ' Account to return from the function
    Dim Account As cPtAccount           ' Account object
    Dim BrokerAccounts As cGdTree       ' Collection of broker accounts
    Dim BrokerAccount As cBrokerMessage ' Account object

    Set ReturnAccount = Nothing
    If bKeyValue Then
        If FormIsLoaded("frmBrokerView") Then
            Set BrokerAccounts = frmBrokerView.Accounts
            If BrokerAccounts.Exists(strAccountNumber) Then
                Set ReturnAccount = KeyValueAccountToTurnkey(BrokerAccounts(strAccountNumber), nBroker)
            End If
        End If
    Else
        Set Account = g.Broker.Account(strAccountNumber)
        If Not Account Is Nothing Then
            Set ReturnAccount = GenesisAccountToTurnkey(Account)
        End If
    End If
    
    Set AccountForBrokerNumber = ReturnAccount

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.AccountForBrokerNumber"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetWorkingOrders
'' Description: Get working orders to display in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetWorkingOrders() As cGdTree
On Error GoTo ErrSection:

    Dim ReturnOrders As cGdTree         ' Collection of orders to return from the function
    Dim Orders As cPtOrders             ' Collection of working orders
    Dim BrokerOrders As cGdTree         ' Collection of orders for a broker
    Dim lIndex As Long                  ' Index into a for loop
    Dim nBroker As eTT_AccountType      ' Broker for the broker view form
    Dim brokerOrder As cBrokerMessage   ' Broker order
    
    Set ReturnOrders = New cGdTree
    Set Orders = g.Broker.AllOrders(True)
    For lIndex = 1 To Orders.Count
        ReturnOrders.Add GenesisOrderToTurnkey(Orders(lIndex))
    Next lIndex
    
    If FormIsLoaded("frmBrokerView") Then
        nBroker = frmBrokerView.Broker
        Set BrokerOrders = frmBrokerView.Orders
        For lIndex = 1 To BrokerOrders.Count
            Set brokerOrder = KeyValueOrderToTurnkey(BrokerOrders(lIndex), nBroker)
            ReturnOrders.Add brokerOrder, brokerOrder("Broker") & "|" & brokerOrder("BrokerOrderID")
        Next lIndex
    End If
    
    Set GetWorkingOrders = ReturnOrders

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cTnCattle.GetWorkingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFills
'' Description: Get fills to display in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetFills() As cGdTree
On Error GoTo ErrSection:

    Dim ReturnFills As cGdTree          ' Collection of Fills to return from the function
    Dim Fills As cPtFills               ' Collection of working Fills
    Dim BrokerFills As cGdTree          ' Collection of Fills for a broker
    Dim lIndex As Long                  ' Index into a for loop
    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim brokerFill As cBrokerMessage    ' Broker fill
    
    Set ReturnFills = New cGdTree
    Set Fills = g.Broker.AllFills
    For lIndex = 1 To Fills.Count
        ReturnFills.Add GenesisFillToTurnkey(Fills(lIndex))
    Next lIndex
    
    If FormIsLoaded("frmBrokerView") Then
        nBroker = frmBrokerView.Broker
        Set BrokerFills = frmBrokerView.Fills
        For lIndex = 1 To BrokerFills.Count
            Set brokerFill = KeyValueFillToTurnkey(BrokerFills(lIndex), nBroker)
            ReturnFills.Add brokerFill, FillKey(brokerFill("Broker"), brokerFill("BrokerFillID"))
        Next lIndex
    End If
    
    Set GetFills = ReturnFills

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cTnCattle.GetFills"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerName
'' Description: Return the name of the given brokerage
'' Inputs:      Broker
'' Returns:     Name of Brokerage
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerName(ByVal nBroker As eTT_AccountType) As String
    BrokerName = g.Broker.BrokerName(nBroker)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectionStatusForAccount
'' Description: Determine the connection status for the given account
'' Inputs:      Account Number, Check Broker View?
'' Returns:     Connection Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConnectionStatusForAccount(ByVal vAccountNumberOrID As Variant, Optional ByVal bCheckBrokerView As Boolean = False) As eGDConnectionStatus
    ConnectionStatusForAccount = g.Broker.ConnectionStatusForAccount(vAccountNumberOrID, bCheckBrokerView)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LastKnownPrice
'' Description: Get the last known price for the given symbol
'' Inputs:      Symbol, Local Time?, Date/Time of last price(return)
''      iBidTradeOrAsk:  -1 = LastBid,  0 = LastTrade,  1 = LastAsk
'' Returns:     Last Known Price (kNullData if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastKnownPrice(ByVal vSymbolOrSymbolID As Variant, Optional ByVal iBidTradeOrAsk As Integer = 0, Optional ByVal bLocalTime As Boolean = False, Optional dDateTime As Double) As Double
    LastKnownPrice = g.RealTime.LastKnownPrice(vSymbolOrSymbolID, iBidTradeOrAsk, bLocalTime, dDateTime)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimerStart
'' Description: Notify that a timer is starting
'' Inputs:      Name of the timer
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TimerStart(ByVal strName As String)
    mMain.TimerStart strName
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimerEnd
'' Description: Notify that a timer is ending
'' Inputs:      Name of the timer, Interval of the timer
'' Returns:     True if logged the time, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TimerEnd(ByVal strName As String, ByVal lInterval As Long) As Boolean
    TimerEnd = mMain.TimerEnd(strName, lInterval)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateOrder
'' Description: Allow the user to create a new order
'' Inputs:      Feed Yard Lot ID, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateOrder(ByVal strFeedYardLotID As String, Optional ByVal strSymbol As String = "")
On Error GoTo ErrSection:
    
    If Len(strSymbol) = 0 Then
        strSymbol = ConvertToTradeSymbol("LE-067", Date)
    End If
    
    If FormIsLoaded("frmBrokerView") Then
        frmBrokerView.CreateOrderForLot strFeedYardLotID, strSymbol
    Else
        mTradeTracker.CreateOrder strSymbol, , , , , , strFeedYardLotID
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTnCattle.CreateOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBrokerOrdersForAccount
'' Description: Get the broker orders for the given account
'' Inputs:      Account
'' Returns:     Orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetBrokerOrdersForAccount(ByVal Account As cBrokerMessage) As cGdTree
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim Orders As cGdTree               ' Collection of orders
    
    Set Orders = New cGdTree
    
    nBroker = CLng(Val(Account("Broker")))
    If g.Broker.ConnectionStatusForAccount(Account("Number"), True) = eGDConnectionStatus_Connected Then
        If g.Broker.IsBroker(nBroker) Then
            Set Orders = GetBrokerOrdersForAccountBroker(Account)
        Else
            Set Orders = GetBrokerOrdersForAccountUser(Account)
        End If
    End If
    
    Set GetBrokerOrdersForAccount = Orders
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.GetAllBrokerOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBrokerFillsForAccount
'' Description: Get the broker fills for the given account
'' Inputs:      Account
'' Returns:     Fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetBrokerFillsForAccount(ByVal Account As cBrokerMessage) As cGdTree
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim Fills As cGdTree                ' Collection of fills
    
    Set Fills = New cGdTree
    
    nBroker = CLng(Val(Account("Broker")))
    If g.Broker.ConnectionStatusForAccount(Account("Number"), True) = eGDConnectionStatus_Connected Then
        If g.Broker.IsBroker(nBroker) Then
            Set Fills = GetBrokerFillsForAccountBroker(Account)
        Else
            Set Fills = GetBrokerFillsForAccountUser(Account)
        End If
    End If
    
    Set GetBrokerFillsForAccount = Fills
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.GetAllBrokerFills"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowLadder
'' Description: Show the price ladder appropriately for the given symbol
'' Inputs:      Genesis Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowLadder(ByVal strGenesisSymbol As String)
On Error GoTo ErrSection:

    Dim frm As frmTickDistribution      ' Price ladder form

    If InStr(strGenesisSymbol, " ") = 0 Then
        If FormIsLoaded("frmBrokerView") Then
            frmBrokerView.ShowLadder strGenesisSymbol
        Else
            Set frm = New frmTickDistribution
            frm.ShowMe GetSymbolID(strGenesisSymbol), 0
        End If
    Else
        InfBox "Price ladder cannot be shown for an option", "!", , "Error"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTnCattle.ShowLadder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetHistoricalFillsForAccount
'' Description: Retrieve historical fills for the given account
'' Inputs:      Account Number
'' Returns:     Collection of historical fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetHistoricalFillsForAccount(ByVal strAccountNumber As String) As cGdTree
On Error GoTo ErrSection:

    Dim Fills As cGdTree                ' Collection of fills to return
    Dim AccountFills As cPtFills        ' Collection of fills for an account
    Dim lIndex As Long                  ' Index into a for loop
    Dim brokerFill As cBrokerMessage    ' Broker fill

    Set Fills = New cGdTree
    Set AccountFills = New cPtFills
    
    AccountFills.LoadHistoricalFillsForAccount strAccountNumber
    For lIndex = 1 To AccountFills.Count
        Set brokerFill = GenesisFillToTurnkey(AccountFills(lIndex))
        Fills.Add brokerFill, FillKey(brokerFill("Broker"), brokerFill("BrokerFillID"))
    Next lIndex
    
    Set GetHistoricalFillsForAccount = Fills

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.GetHistoricalFillsForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridScrollCheck
'' Description: Help fix the inadvertant scrolling issues with the FlexGrids
''              while streaming is on
'' Inputs:      Grid, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GridScrollCheck(fg As VSFlexGrid, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    mMain.GridScrollCheck fg, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditLotColumns
'' Description: Allow the user to edit lot columns
'' Inputs:      Columns
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EditLotColumns(hastrColumns As Long) As Boolean
On Error GoTo ErrSection:

    Dim astrColumns As cGdArray         ' Array for the handle passed in
    
    Set astrColumns = New cGdArray
    astrColumns.SetArrayHandle hastrColumns, False

    EditLotColumns = frmQuoteBoardFields.ShowMe(astrColumns, eQbfMode_Cattle)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.EditLotColumns"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetMachineID
'' Description: Get the machine ID
'' Inputs:      None
'' Returns:     Machine ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetMachineID() As String
    GetMachineID = RI_GetMachineID
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPassword
'' Description: Get the pasword
'' Inputs:      None
'' Returns:     Password
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetPassword() As String
    GetPassword = Trim(RI_GetUserPassword)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountLookup
'' Description: Allow the user to lookup an account
'' Inputs:      Accounts, Default account number, Default customer, Cattle?
'' Returns:     Account Number (or blank if Cancelled)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountLookup(ByVal hastrAccounts As Long, Optional ByVal strAccountNumber As String = "", Optional ByVal strCustomer As String = "", Optional ByVal bForCattle As Boolean = False) As String
On Error GoTo ErrSection:

    Dim astrAccounts As cGdArray        ' Array for the handle passed in
    
    Set astrAccounts = New cGdArray
    astrAccounts.CopyFromHandle hastrAccounts

    AccountLookup = frmAccountLookup.ShowMe(astrAccounts, strAccountNumber, strCustomer, bForCattle)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.EditLotColumns"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cattle_Customer
'' Description: Handle a new cattle customer being added
'' Inputs:      Customer information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Cattle_Customer(cattleMessage As cBrokerMessage)
On Error GoTo ErrSection:

    If FormIsLoaded("frmAccountLookup") Then
        If frmAccountLookup.ForCattle Then
            frmAccountLookup.Cattle_Customer cattleMessage
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTnCattle.Cattle_Customer", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerViewLoaded
'' Description: Return if the broker view form is loaded
'' Inputs:      None
'' Returns:     True if Broker View form loaded, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerViewLoaded() As Boolean
    BrokerViewLoaded = FormIsLoaded("frmBrokerView")
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetCarriedFillsForBroker
'' Description: Get a collection of carried fills from the broker view
'' Inputs:      None
'' Returns:     Collection of carried fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCarriedFillsForBroker() As cGdTree
On Error GoTo ErrSection:

    Dim ReturnFills As cGdTree          ' Collection of fills to return
    Dim CarriedFills As cGdTree         ' Carried fills from broker view form
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cBrokerMessage          ' Converted fill
    Dim nBroker As eTT_AccountType      ' Broker for the account
    
    Set ReturnFills = New cGdTree
    If BrokerViewLoaded Then
        nBroker = frmBrokerView.Broker
        'Set CarriedFills = g.Broker.CarriedFillsForBroker(nBroker)
        Set CarriedFills = frmBrokerView.CarriedFills
        
        For lIndex = 1 To CarriedFills.Count
            Set Fill = KeyValueFillToTurnkey(CarriedFills(lIndex), nBroker)
            AddBrokerIdToCarriedFill Fill, Fill("BrokerAccountNumber"), ReturnFills
            ReturnFills.Add Fill, FillKey(Str(nBroker), Fill("BrokerFillID"))
        Next lIndex
    End If

    Set GetCarriedFillsForBroker = ReturnFills

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.GetCarriedFillsForBroker"
    
End Function

'*************************************

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    KeyValueOrderToTurnkey
'' Description: Convert a key-value order to Turnkey
'' Inputs:      Key-Value Order
'' Returns:     Turnkey Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function KeyValueOrderToTurnkey(KeyValueOrder As cBrokerMessage, ByVal nBroker As eTT_AccountType) As cBrokerMessage
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Turnkey order
    Dim BrokerKv As cBrokerKeyValue     ' Key Value routines
    Dim strIsBuy As String              ' Is the leg to buy?
    
    Set turnkeyMessage = New cBrokerMessage
    Set BrokerKv = New cBrokerKeyValue
    
    turnkeyMessage.Add "Broker", Str(nBroker)
    'turnkeyMessage.Add "BrokerAccountID", AccountIdForBrokerNumber(KeyValueOrder("Account"), nBroker)
    turnkeyMessage.Add "BrokerAccountNumber", KeyValueOrder("Account")
    turnkeyMessage.Add "FcmAccountNumber", KeyValueOrder("FcmAccount")
    turnkeyMessage.Add "BrokerOrderID", KeyValueOrder("BrokerID")
    turnkeyMessage.Add "GenesisOrderID", KeyValueOrder("GenesisID")
    turnkeyMessage.Add "PreviousBrokerOrderID", KeyValueOrder("PreviousBrokerID")
    turnkeyMessage.Add "Quantity", KeyValueOrder("Quantity")
    turnkeyMessage.Add "Type", KeyValueOrder("GenesisType")
    turnkeyMessage.Add "LimitPrice", KeyValueOrder("GenesisLimitPrice")
    turnkeyMessage.Add "StopPrice", KeyValueOrder("GenesisStopPrice")
    turnkeyMessage.Add "TimeInForce", KeyValueOrder("GenesisTif")
    turnkeyMessage.Add "ExpirationDate", KeyValueOrder("Expiration")
    turnkeyMessage.Add "Status", KeyValueOrder("GenesisStatus")
    turnkeyMessage.Add "StatusDate", Str(BrokerKv.DateToDouble(KeyValueOrder("StatusDate")))
    If IsOpenOrder(CLng(Val(KeyValueOrder("GenesisStatus")))) Then
        turnkeyMessage.Add "IsWorking", "1"
    Else
        turnkeyMessage.Add "IsWorking", "0"
    End If
    turnkeyMessage.Add "NumberOfLegs", "1"
    
    If UCase(KeyValueOrder("Side")) = "BUY" Then
        strIsBuy = "1"
    Else
        strIsBuy = "0"
    End If
    turnkeyMessage.Add "Leg1", "1," & strIsBuy & "," & KeyValueOrder("GenesisSymbol") & ",1,1"
    turnkeyMessage.Add "Symbol", KeyValueOrder("GenesisSymbol")
    
    Set KeyValueOrderToTurnkey = turnkeyMessage

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.KeyValueOrderToTurnkey"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderToTurnkey
'' Description: Convert a Genesis order to Turnkey
'' Inputs:      Genesis Order
'' Returns:     Turnkey Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisOrderToTurnkey(ByVal Order As cPtOrder) As cBrokerMessage
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Turnkey order
    Dim strAccountNumber As String      ' Account number
    Dim lIndex As Long                  ' Index into a for loop
    Dim strIsBuy As String              ' Is the leg to buy?
    Dim strIsEntry As String            ' Is the leg to enter?
    
    Set turnkeyMessage = New cBrokerMessage
    
    strAccountNumber = g.Broker.AccountNumberForID(Order.AccountID)
    turnkeyMessage.Add "Broker", Str(Order.Broker)
    'turnkeyMessage.Add "BrokerAccountID", AccountIdForBrokerNumber(strAccountNumber, Order.Broker)
    turnkeyMessage.Add "BrokerAccountNumber", strAccountNumber
    turnkeyMessage.Add "FcmAccountNumber", ""
    turnkeyMessage.Add "BrokerOrderID", Order.BrokerID
    turnkeyMessage.Add "GenesisOrderID", Order.GenesisOrderID
    turnkeyMessage.Add "PreviousBrokerOrderID", Order.PreviousBrokerID
    turnkeyMessage.Add "Quantity", Str(Order.Quantity)
    turnkeyMessage.Add "Type", Str(Order.OrderType)
    turnkeyMessage.Add "LimitPrice", Str(Order.LimitPrice)
    turnkeyMessage.Add "StopPrice", Str(Order.StopPrice)
    turnkeyMessage.Add "TimeInForce", Str(Order.TimeInForce)
    turnkeyMessage.Add "ExpirationDate", Str(Order.Expiration)
    turnkeyMessage.Add "Status", Str(Order.Status)
    turnkeyMessage.Add "StatusDate", Str(Order.StatusDate)
    If IsOpenOrder(Order.Status) Then
        turnkeyMessage.Add "IsWorking", "1"
    Else
        turnkeyMessage.Add "IsWorking", "0"
    End If
    turnkeyMessage.Add "NumberOfLegs", Str(Order.NumberOfLegs)
    
    For lIndex = 1 To Order.NumberOfLegs
        With Order.OrderLegs(lIndex)
            If .IsBuy Then
                strIsBuy = "1"
            Else
                strIsBuy = "0"
            End If
            If .IsEntry Then
                strIsEntry = "1"
            Else
                strIsEntry = "0"
            End If
            turnkeyMessage.Add "Leg" & Str(lIndex), Str(.LegNumber) & "," & strIsBuy & "," & .Symbol & "," & Str(.Multiplier) & "," & strIsEntry
        End With
    Next lIndex
    
    If Order.NumberOfLegs = 1 Then
        turnkeyMessage.Add "Symbol", Order.Symbol
    Else
        turnkeyMessage.Add "Symbol", Order.SpreadSymbol
    End If
        
    Set GenesisOrderToTurnkey = turnkeyMessage

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.GenesisOrderToTurnkey"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    KeyValueFillToTurnkey
'' Description: Convert a key-value fill to Turnkey
'' Inputs:      Key-Value Fill
'' Returns:     Turnkey Fill
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function KeyValueFillToTurnkey(KeyValueFill As cBrokerMessage, ByVal nBroker As eTT_AccountType) As cBrokerMessage
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Turnkey fill
    Dim BrokerKv As cBrokerKeyValue     ' Key Value routines

    Set turnkeyMessage = New cBrokerMessage
    Set BrokerKv = New cBrokerKeyValue
    
    turnkeyMessage.Add "Broker", Str(nBroker)
    'turnkeyMessage.Add "BrokerAccountID", AccountIdForBrokerNumber(KeyValueFill("Account"), nBroker)
    turnkeyMessage.Add "BrokerAccountNumber", KeyValueFill("Account")
    turnkeyMessage.Add "FcmAccount", KeyValueFill("FcmAccount")
    turnkeyMessage.Add "BrokerOrderID", KeyValueFill("BrokerID")
    turnkeyMessage.Add "BrokerFillID", KeyValueFill("FillID")
    turnkeyMessage.Add "FillTime", Str(BrokerKv.DateToDouble(KeyValueFill("FillDate")))
    turnkeyMessage.Add "Symbol", KeyValueFill("GenesisSymbol")
    If UCase(KeyValueFill("Side")) = "BUY" Then
        turnkeyMessage.Add "IsBuy", "1"
    Else
        turnkeyMessage.Add "IsBuy", "0"
    End If
    turnkeyMessage.Add "Quantity", KeyValueFill("Quantity")
    turnkeyMessage.Add "Price", KeyValueFill("GenesisFillPrice")
    
    Set KeyValueFillToTurnkey = turnkeyMessage

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.KeyValueFillToTurnkey"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisFillToTurnkey
'' Description: Convert a Genesis fill to Turnkey
'' Inputs:      Genesis Fill
'' Returns:     Turnkey Fill
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisFillToTurnkey(ByVal Fill As cPtFill) As cBrokerMessage
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Turnkey fill
    Dim strAccountNumber As String      ' Account number

    Set turnkeyMessage = New cBrokerMessage
    
    strAccountNumber = g.Broker.AccountNumberForID(Fill.AccountID)
    turnkeyMessage.Add "Broker", Str(Fill.Broker)
    'turnkeyMessage.Add "BrokerAccountID", AccountIdForBrokerNumber(strAccountNumber, Fill.Broker)
    turnkeyMessage.Add "BrokerAccountNumber", strAccountNumber
    turnkeyMessage.Add "FcmAccount", ""
    turnkeyMessage.Add "BrokerOrderID", Fill.BrokerOrderID
    turnkeyMessage.Add "BrokerFillID", Fill.BrokerID
    turnkeyMessage.Add "FillTime", Str(Fill.FillDate)
    turnkeyMessage.Add "Symbol", Fill.Symbol
    If Fill.Buy Then
        turnkeyMessage.Add "IsBuy", "1"
    Else
        turnkeyMessage.Add "IsBuy", "0"
    End If
    turnkeyMessage.Add "Quantity", Str(Fill.Quantity)
    turnkeyMessage.Add "Price", Str(Fill.Price)
    turnkeyMessage.Add "Commission", Str(Fill.Fees)
    
    Set GenesisFillToTurnkey = turnkeyMessage

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.GenesisFillToTurnkey"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    KeyValueAccountToTurnkey
'' Description: Convert a key-value account to Turnkey
'' Inputs:      Key-Value Account, Broker
'' Returns:     Turnkey Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function KeyValueAccountToTurnkey(KeyValueAccount As cBrokerMessage, ByVal nBroker As eTT_AccountType) As cBrokerMessage
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Turnkey fill

    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "Number", KeyValueAccount("Account")
    turnkeyMessage.Add "Name", KeyValueAccount("AccountName")
    turnkeyMessage.Add "Broker", Str(nBroker)
    turnkeyMessage.Add "FcmNumber", KeyValueAccount("FcmAccount")

    Set KeyValueAccountToTurnkey = turnkeyMessage

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.KeyValueAccountToTurnkey"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisAccountToTurnkey
'' Description: Convert a Genesis account to Turnkey
'' Inputs:      Genesis Account
'' Returns:     Turnkey Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisAccountToTurnkey(ByVal Account As cPtAccount) As cBrokerMessage
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Turnkey fill

    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.Add "Number", Account.AccountNumber
    turnkeyMessage.Add "Name", Account.Name
    turnkeyMessage.Add "Broker", Str(Account.AccountType)
    turnkeyMessage.Add "FcmNumber", Account.FcmAccountNumber

    Set GenesisAccountToTurnkey = turnkeyMessage

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.GenesisAccountToTurnkey"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowCattleFill
'' Description: Show the edit fill form for the given cattle fill
'' Inputs:      Cattle Fill
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowCattleFill(CattleFill As cBrokerMessage) As Boolean
On Error GoTo ErrSection:

    ShowCattleFill = frmTTEditFill.ShowMeCattle(CattleFill)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.ShowCattleFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBrokerFillsForAccountUser
'' Description: Get the broker fills for the given account
'' Inputs:      Account
'' Returns:     Collection of broker fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetBrokerFillsForAccountUser(ByVal Account As cBrokerMessage) As cGdTree
On Error GoTo ErrSection:

    Dim ReturnFills As cGdTree          ' Collection of fills to return
    Dim BrokerFills As cPtFills         ' Collection of broker fills
    Dim CarriedFills As cGdTree         ' Collection of carried fills
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cBrokerMessage          ' Turnkey fill
    Dim strAccountNumber As String      ' Account number
    Dim strCarriedFillId As String      ' Broker Fill ID for a carried fill
    Dim lCounter As Long                ' Counter variable
    Dim strBase As String               ' Base of the broker fill ID
    Dim strBroker As String             ' Broker

    Set ReturnFills = New cGdTree
    strAccountNumber = Account("Number")
    strBroker = Account("Broker")

    Set BrokerFills = g.Broker.FillsForAccount(strAccountNumber)
    If Not BrokerFills Is Nothing Then
        ' If no fills have been sent for this account, we need to send the
        ' carried fill information...
        If Account("HasFills") = "0" Then
            Set CarriedFills = g.Broker.CarriedFillsForAccount(strAccountNumber).MakeCopy
            For lIndex = 1 To CarriedFills.Count
                Set Fill = GenesisFillToTurnkey(CarriedFills(lIndex))
                AddBrokerIdToCarriedFill Fill, strAccountNumber, ReturnFills
                ReturnFills.Add Fill, FillKey(strBroker, Fill("BrokerFillID"))
            Next lIndex
        End If
        
        For lIndex = 1 To BrokerFills.Count
            Set Fill = GenesisFillToTurnkey(BrokerFills(lIndex))
            ReturnFills.Add Fill, FillKey(strBroker, Fill("BrokerFillID"))
        Next lIndex
    End If
    
    Set GetBrokerFillsForAccountUser = ReturnFills

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.GetBrokerFillsForAccountUser"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBrokerFillsForAccountBroker
'' Description: Get the broker fills for the given account
'' Inputs:      Account
'' Returns:     Collection of broker fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetBrokerFillsForAccountBroker(ByVal Account As cBrokerMessage) As cGdTree
On Error GoTo ErrSection:

    Dim ReturnFills As cGdTree          ' Collection of fills to return
    Dim BrokerFills As cGdTree          ' Collection of broker fills
    Dim CarriedFills As cGdTree         ' Collection of carried fills
    Dim Fill As cBrokerMessage          ' Turnkey fill
    Dim strAccountNumber As String      ' Account number
    Dim lIndex As Long                  ' Index into a for loop
    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim strBroker As String             ' Broker
    Dim lTimeOut As Long                ' Time out variable
    
    Set ReturnFills = New cGdTree
    strAccountNumber = Account("Number")
    strBroker = Account("Broker")
    nBroker = CLng(Val(strBroker))
    
    If FormIsLoaded("frmBrokerView") Then
        If frmBrokerView.Broker = nBroker Then
            If frmBrokerView.AccountNumber <> strAccountNumber Then
                frmBrokerView.RefreshAccount strAccountNumber
            End If
            
            lTimeOut = 0&
            Do While (frmBrokerView.RefreshDone = False) And (lTimeOut < 30&)
                Sleep 1#
                lTimeOut = lTimeOut + 1&
            Loop
            
            If (frmBrokerView.RefreshDone = True) And (frmBrokerView.RefreshTimedOut = False) Then
                Set BrokerFills = frmBrokerView.RefreshFills
                If Account("HasFills") = "0" Then
                    Set CarriedFills = frmBrokerView.RefreshCarriedFills
                    For lIndex = 1 To CarriedFills.Count
                        Set Fill = KeyValueFillToTurnkey(CarriedFills(lIndex), nBroker)
                        AddBrokerIdToCarriedFill Fill, strAccountNumber, ReturnFills
                        ReturnFills.Add Fill, FillKey(strBroker, Fill("BrokerFillID"))
                    Next lIndex
                End If
        
                For lIndex = 1 To BrokerFills.Count
                    Set Fill = KeyValueFillToTurnkey(BrokerFills(lIndex), nBroker)
                    ReturnFills.Add Fill, FillKey(strBroker, Fill("BrokerFillID"))
                Next lIndex
            End If
        End If
    End If
    
    Set GetBrokerFillsForAccountBroker = ReturnFills

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.GetBrokerFillsForAccountBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddBrokerIdToCarriedFill
'' Description: Add a generated broker fill ID to the carried fill
'' Inputs:      Fill, Account Number, Fills Collection
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddBrokerIdToCarriedFill(Fill As cBrokerMessage, ByVal strAccountNumber As String, ByVal Fills As cGdTree)
On Error GoTo ErrSection:

    Dim strBase As String               ' Base of the broker fill ID
    Dim strBrokerFillID As String       ' Generated broker fill ID
    Dim lCounter As Long                ' Counter variable
    Dim Bars As cGdBars                 ' Bars object
    
    Set Bars = New cGdBars
    SetBarProperties Bars, Fill("Symbol")

    strBase = "CF-" & strAccountNumber & "-" & Fill("Symbol") & "-" & Format(Val(Fill("FillTime")), "YYYYMMDD") & "-" & Bars.PriceDisplay(Val(Fill("Price")), True)
    strBrokerFillID = strBase
    
    lCounter = 0&
    Do While Fills.Exists(FillKey(Fill("Broker"), strBrokerFillID))
        lCounter = lCounter + 1&
        strBrokerFillID = strBase & "-" & Str(lCounter)
    Loop
                
    Fill.Add "BrokerFillID", strBrokerFillID
    Fill.Add "BrokerOrderID", strBrokerFillID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTnCattle.AddBrokerIdToCarriedFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBrokerOrdersForAccountUser
'' Description: Get the broker orders for the given account
'' Inputs:      Account
'' Returns:     Collection of broker orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetBrokerOrdersForAccountUser(ByVal Account As cBrokerMessage) As cGdTree
On Error GoTo ErrSection:

    Dim ReturnOrders As cGdTree         ' Collection of orders to return
    Dim BrokerOrders As cPtOrders       ' Collection of broker orders
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cBrokerMessage         ' Turnkey order
    Dim strAccountNumber As String      ' Account number
    Dim lCounter As Long                ' Counter variable
    Dim strBase As String               ' Base of the broker fill ID
    Dim strBroker As String             ' Broker

    Set ReturnOrders = New cGdTree
    strAccountNumber = Account("Number")
    strBroker = Account("Broker")

    Set BrokerOrders = g.Broker.OrdersForAccount(strAccountNumber)
    If Not BrokerOrders Is Nothing Then
        For lIndex = 1 To BrokerOrders.Count
            Set Order = GenesisOrderToTurnkey(BrokerOrders(lIndex))
            ReturnOrders.Add Order, OrderKey(strBroker, Order("BrokerOrderID"))
        Next lIndex
    End If
    
    Set GetBrokerOrdersForAccountUser = ReturnOrders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.GetBrokerOrdersForAccountUser"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBrokerOrdersForAccountBroker
'' Description: Get the broker orders for the given account
'' Inputs:      Account
'' Returns:     Collection of broker orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetBrokerOrdersForAccountBroker(ByVal Account As cBrokerMessage) As cGdTree
On Error GoTo ErrSection:

    Dim ReturnOrders As cGdTree         ' Collection of orders to return
    Dim BrokerOrders As cGdTree         ' Collection of broker fills
    Dim Order As cBrokerMessage         ' Turnkey order
    Dim strAccountNumber As String      ' Account number
    Dim lIndex As Long                  ' Index into a for loop
    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim strBroker As String             ' Broker
    Dim lTimeOut As Long                ' Time out variable
    
    Set ReturnOrders = New cGdTree
    strAccountNumber = Account("Number")
    strBroker = Account("Broker")
    nBroker = CLng(Val(strBroker))
    
    If FormIsLoaded("frmBrokerView") Then
        If frmBrokerView.Broker = nBroker Then
            If frmBrokerView.AccountNumber <> strAccountNumber Then
                frmBrokerView.RefreshAccount strAccountNumber
            End If
            
            lTimeOut = 0&
            Do While (frmBrokerView.RefreshDone = False) And (lTimeOut < 30&)
                Sleep 1#
                lTimeOut = lTimeOut + 1&
            Loop
            
            If (frmBrokerView.RefreshDone = True) And (frmBrokerView.RefreshTimedOut = False) Then
                Set BrokerOrders = frmBrokerView.RefreshOrders
        
                For lIndex = 1 To BrokerOrders.Count
                    Set Order = KeyValueOrderToTurnkey(BrokerOrders(lIndex), nBroker)
                    ReturnOrders.Add Order, OrderKey(strBroker, Order("BrokerOrderID"))
                Next lIndex
            End If
        End If
    End If
    
    Set GetBrokerOrdersForAccountBroker = ReturnOrders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCattle.GetBrokerOrdersForAccountBroker"
    
End Function
