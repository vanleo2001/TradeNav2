VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCattleTn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cCattleTn.cls
'' Description: Application bridge for the Cattle DLL to talk to Trade Navigator
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
'' 04/28/2014   DAJ         Added the GetBrokerOrdersForAccount function
'' 05/22/2014   DAJ         Renamed g.Turnkey to g.Cattle
'' 05/30/2014   DAJ         Added ShowCattleFill; Utilized new accounts object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    AppBridge As Object                 ' Application bridge
End Type
Private m As mPrivate

Public Property Get AppBridge() As Object
    Set AppBridge = m.AppBridge
End Property
Public Property Let AppBridge(TheAppBridge As Object)
    Set m.AppBridge = TheAppBridge
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
    Set Picture16 = AppBridge.Picture16(strPicture, iImageList)
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
    PriceDisplay = AppBridge.PriceDisplay(dPrice, strSymbol)
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
    HasModule = AppBridge.HasModule(strModules, bIncludeSourceCode)
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
    AppBridge.AddDialogMessage strMessage, strIcon, strButtons, strCaption
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
    Unloading = AppBridge.Unloading
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FormatVersion
'' Description: Formatted version of the Trade Navigator version
'' Inputs:      Include Revision?, Include File Date?
'' Returns:     Formatted version
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FormatVersion(Optional ByVal bIncludeRevision As Boolean = False, Optional ByVal bIncludeFileDate As Boolean = False) As String
    FormatVersion = AppBridge.FormatVersion(bIncludeRevision, bIncludeFileDate)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoPrintHeader
'' Description: Do the print header for Trade Navigator
'' Inputs:      Font Size, Printer Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoPrintHeader(Optional ByVal nFontSize& = 12)
    AppBridge.DoPrintHeader nFontSize, frmPrintPreview.vp
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
    ToolbarIcon = AppBridge.ToolbarIcon(strID)
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
    SecurityType = AppBridge.SecurityType(BarsOrSymbol, bIfOptionAppendO)
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
Public Function Profit(ByVal strSymbol As String, ByVal dPriceDifference As Double, Optional ByVal lQuantity As Long = 1&, Optional ByVal dExchangeExitTime As Double = -99999#, Optional ByVal SymbolBars As cGdBars = Nothing, Optional ByVal BaseForexBars As cGdBars = Nothing, Optional ByVal vAccountNumberOrID As Variant = "") As Double
On Error GoTo ErrSection:

    Dim hSymbolBars As Long             ' Bars handle for the symbol bars
    Dim hBaseForexBars As Long          ' Bars handle for the base forex bars
    
    If SymbolBars Is Nothing Then
        hSymbolBars = 0
    Else
        hSymbolBars = SymbolBars.BarsHandle
    End If
    
    If BaseForexBars Is Nothing Then
        hBaseForexBars = 0
    Else
        hBaseForexBars = BaseForexBars.BarsHandle
    End If

    Profit = AppBridge.Profit(strSymbol, dPriceDifference, lQuantity, dExchangeExitTime, hSymbolBars, hBaseForexBars, vAccountNumberOrID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattleTn.Profit"
    
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
    Dim Orders As Object                ' Collection of working orders
    Dim lIndex As Long                  ' Index into a for loop
    Dim brokerOrder As cBrokerMessage   ' Broker order
    
    Set ReturnOrders = New cGdTree
    Set Orders = AppBridge.GetWorkingOrders
    
    For lIndex = 1 To Orders.Count
        Set brokerOrder = Orders(lIndex)
        
        brokerOrder.Add "BrokerAccountID", g.Cattle.Accounts.AccountIdForBrokerNumber(brokerOrder("BrokerAccountNumber"), CLng(Val(brokerOrder("Broker"))))
        ReturnOrders.Add brokerOrder
    Next lIndex
    
    Set GetWorkingOrders = ReturnOrders

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cCattleTn.GetWorkingOrders"
    
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

    Dim ReturnFills As cGdTree          ' Collection of fills to return
    Dim Fills As Object                 ' Collection of fills from Trade Navigator
    
    Set Fills = AppBridge.GetFills
    Set ReturnFills = CopyTreeFromObject(Fills)
    
    AddBrokerAccountNumberToFills ReturnFills
    
    Set GetFills = ReturnFills

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cCattleTn.GetFills"
    
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
    BrokerName = AppBridge.BrokerName(nBroker)
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
    ConnectionStatusForAccount = AppBridge.ConnectionStatusForAccount(vAccountNumberOrID, bCheckBrokerView)
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
    LastKnownPrice = AppBridge.LastKnownPrice(vSymbolOrSymbolID, iBidTradeOrAsk, bLocalTime, dDateTime)
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
    AppBridge.TimerStart strName
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
    TimerEnd = AppBridge.TimerEnd(strName, lInterval)
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
    AppBridge.CreateOrder strFeedYardLotID, strSymbol
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBrokerAccounts
'' Description: Get a list of broker accounts for associations
'' Inputs:      None
'' Returns:     Broker Accounts
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetBrokerAccounts() As cGdTree
    Set GetBrokerAccounts = CopyTreeFromObject(AppBridge.GetBrokerAccounts)
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
    Set AccountForBrokerNumber = AppBridge.AccountForBrokerNumber(strAccountNumber, nBroker, bKeyValue)
End Function

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

    Dim BrokerOrders As Object          ' Collection of orders back from Trade Navigator
    Dim Orders As cGdTree               ' Collection of orders
    
    Set BrokerOrders = AppBridge.GetBrokerOrdersForAccount(Account)
    Set Orders = CopyTreeFromObject(BrokerOrders)
    
    AddBrokerAccountNumberToOrders Orders
    
    Set GetBrokerOrdersForAccount = Orders

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cCattleTn.GetBrokerOrdersForAccount"
    
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

    Dim BrokerFills As Object           ' Collection of fills back from Trade Navigator
    Dim Fills As cGdTree                ' Collection of working Fills
    
    Set BrokerFills = AppBridge.GetBrokerFillsForAccount(Account)
    Set Fills = CopyTreeFromObject(BrokerFills)
    
    AddBrokerAccountNumberToFills Fills
    
    Set GetBrokerFillsForAccount = Fills

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cCattleTn.GetBrokerFillsForAccount"
    
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
    AppBridge.ShowLadder strGenesisSymbol
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

    Dim HistoricalFills As Object       ' Collection of historical fills from Trade Navigator
    Dim Fills As cGdTree                ' Collection of historical fills
    
    Set HistoricalFills = AppBridge.GetHistoricalFillsForAccount(strAccountNumber)
    Set Fills = CopyTreeFromObject(HistoricalFills)
    
    AddBrokerAccountNumberToFills Fills
    
    Set GetHistoricalFillsForAccount = Fills

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cCattleTn.GetHistoricalFillsForAccount"
    
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
    AppBridge.GridScrollCheck fg, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditLotColumns
'' Description: Allow the user to edit lot columns
'' Inputs:      Columns
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EditLotColumns(astrColumns As cGdArray) As Boolean
    EditLotColumns = AppBridge.EditLotColumns(astrColumns.ArrayHandle)
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
    GetMachineID = AppBridge.GetMachineID
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
    GetPassword = AppBridge.GetPassword
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountLookup
'' Description: Allow the user to lookup an account
'' Inputs:      Accounts, Default account number, Default customer, Turnkey?
'' Returns:     Account Number (or blank if Cancelled)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountLookup(ByVal astrAccounts As cGdArray, Optional ByVal strAccountNumber As String = "", Optional ByVal strCustomer As String = "", Optional ByVal bForTurnkey As Boolean = False) As String
    AccountLookup = AppBridge.AccountLookup(astrAccounts.ArrayHandle, strAccountNumber, strCustomer, bForTurnkey)
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
    AppBridge.Cattle_Customer cattleMessage
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
    BrokerViewLoaded = AppBridge.BrokerViewLoaded
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

    Dim CarriedFills As Object          ' Collection of carried fills from Trade Navigator
    Dim Fills As cGdTree                ' Collection of carried fills
    
    Set CarriedFills = AppBridge.GetCarriedFillsForBroker
    Set Fills = CopyTreeFromObject(CarriedFills)
    
    AddBrokerAccountNumberToFills Fills
    
    Set GetCarriedFillsForBroker = Fills

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cCattleTn.GetCarriedFillsForBroker"
    
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

    ShowCattleFill = AppBridge.ShowCattleFill(CattleFill)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattleTn.ShowCattleFill"
    
End Function
