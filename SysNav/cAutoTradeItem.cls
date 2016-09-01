VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAutoTradeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAutoTradeItem.cls
'' Description: Object to manage an Automated Trading Item
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 05/06/2009   DAJ         Fail on enable if no data in the bars
'' 05/20/2009   DAJ         Round order prices (stop, limit) to nearest tick
'' 05/20/2009   DAJ         Turn pyramiding off during automtated trading
'' 09/01/2009   DAJ         Use new Parked order status
'' 01/06/2010   DAJ         Fixed position just changed for intraday bar periods
'' 03/09/2010   DAJ         Added app path to the Orders.DAJ file open
'' 03/11/2010   DAJ         Added RefreshUI function
'' 03/15/2010   DAJ         Attempt to load bars in reports if no data
'' 10/04/2010   DAJ         Only allow enable if user has at least platinum (#5957)
'' 05/16/2011   DAJ         Set OrderDate correctly in case delayed streaming
'' 05/27/2011   DAJ         Set the IsAutomated flag for orders that get created
'' 06/21/2011   DAJ         Separate out Simulated trading types
'' 07/15/2011   DAJ         Enhancements for allowing auto trading on continuous contracts
'' 07/21/2011   DAJ         Further tweaks for rolling a continuous contract
'' 07/22/2011   DAJ         Dump tick count in debug messages
'' 07/22/2011   DAJ         Get current position from fill summary instead of fill
'' 07/25/2011   DAJ         Added delete function and deleted flag
'' 07/26/2011   DAJ         Set m.bBetween before calling NewBarCheck
'' 08/02/2011   DAJ         Added the RefreshPosition function
'' 08/17/2011   DAJ         When calculating Trade Symbol, use session date
'' 08/19/2011   DAJ         Added ReloadData function
'' 08/22/2011   DAJ         Changed ReloadData function (don't use NextBarCheck)
'' 08/29/2011   DAJ         Change Position symbol after a roll even if flat
'' 08/31/2011   DAJ         Cancel any working orders if change position symbol when flat
'' 09/06/2011   DAJ         Fix for calculating on-close time
'' 09/08/2011   DAJ         Don't check trade server orders, fix for on-close check
'' 09/16/2011   DAJ         Fix for cancelling orders in ReloadData when symbol changes
'' 09/28/2011   DAJ         Enhanced the automated journaling
'' 01/17/2012   DAJ         Refresh the UI on an order status change
'' 01/18/2012   DAJ         Enhanced logging for automated trading
'' 01/26/2012   DAJ         Fixed issue with changing the quantity of the next entry
'' 02/14/2012   DAJ         New status alerts for position mismatch / auto trade disabled
'' 02/24/2012   DAJ         Fixed exit position when disabling trade item (#6608)
'' 02/28/2012   DAJ         Fixed MOC order confirmation on each bar (#6607)
'' 02/29/2012   DAJ         Don't allow enable if ForceLimitThrough flag is false
'' 03/19/2012   DAJ         Fixed default return value for confirmation question in ClosePosition
'' 06/28/2012   DAJ         Run GenerateOrders when position goes flat
'' 07/31/2012   DAJ         Because of previous fix, check if rolling when position goes flat
'' 08/14/2012   DAJ         Wait for all data to be available before enabling auto trade item
'' 08/15/2012   DAJ         Once data is available, reload bars before enabling auto trade item
'' 08/29/2012   DAJ         If user changes entry quantity while between, don't sync orders
'' 09/06/2012   DAJ         Made and use CancelOrder routine, handle RunEngine issues
'' 11/20/2012   DAJ         Do "OnCloseCheck" in enable if "between"
'' 11/20/2012   DAJ         Allow "NewBarCheck" if symbol different, but flat and not rolling
'' 12/11/2012   DAJ         Handle broker status in automated trading item
'' 12/11/2012   DAJ         Fix for calculating if "in between"
'' 12/17/2012   DAJ         Store off and use OnClose time in exchange time
'' 01/03/2013   DAJ         Fixed calcuation for "in between" again
'' 01/04/2013   DAJ         Fixed calcuation for "in between" again
'' 01/08/2013   DAJ         Only do automatic OnCloseCheck when in between On Close time and close time
'' 01/22/2013   DAJ         Handle consolidated orders coming back as "N/A"
'' 02/11/2013   DAJ         Change check for position just chnging in OnCloseCheck
'' 02/12/2013   DAJ         Set unused order prices to Null instead of Zero
'' 04/03/2013   DAJ         Automated Strategy Baskets
'' 05/01/2013   DAJ         Shadow Trading
'' 05/06/2013   DAJ         Don't try to park order in SubmitOneOrder if the order is still pending
'' 05/14/2013   DAJ         When checking if guru, load the basket even if not the owner
'' 05/21/2013   DAJ         If closing a position, wait until position is closed before doing anything
'' 05/28/2013   DAJ         Don't try to do things when not connected to the broker
'' 05/29/2013   DAJ         Queue up NewBarCheck and OnCloseCheck so that DoEvents doesn't block data update
'' 06/03/2013   DAJ         Only do a 'DoEvents' after a CancelAll if we are actually cancelling orders
'' 06/03/2013   DAJ         Set system Market1 bars before asking for secondary markets ( #6853 )
'' 06/06/2013   DAJ         Fix cound of active automated trading item in toolbar
'' 06/07/2013   DAJ         When Disable, check to see if parent should be disabled as well
'' 08/13/2013   DAJ         Reload bar properties if symbol changes and no data has been loaded
'' 08/15/2013   DAJ         Don't attempt a reload of data for a parent
'' 08/21/2013   DAJ         When reloading data, check for ClosingPosition instead of ClosePosition
'' 08/27/2013   DAJ         Submit GTC orders for @PFG, @IB, and @CNX forex symbols
'' 09/26/2013   DAJ         Tweak to the fill check when the positions don't match
'' 09/26/2013   DAJ         Further tweak to the fill check when the positions don't match
'' 01/14/2014   DAJ         Added 'Order Rejected' alert
'' 02/05/2014   DAJ         Log last and previous bar time and last price for secondary markets
'' 03/12/2014   DAJ         Moved stuff into Cattle and Broker DLL's
'' 03/19/2014   DAJ         Allow pyramiding with an automated trading item if flag file exists
'' 04/04/2014   DAJ         Allow automated trading for pyramiding systems
'' 04/10/2014   DAJ         Fix for sync orders when receive a partial fill ( Michael Cook )
'' 04/23/2014   DAJ         Allow FractZen bars for automated trading
'' 05/01/2014   DAJ         Fix for bug where the auto trade item won't get you into a position
'' 05/05/2014   DAJ         Switched buttons on Exit Pos/Keep Pos dialog
'' 05/06/2014   DAJ         Tweaked logging for FractZen period; Believe position on NewBarCheck if
''                          FractZen period just changed this bar; Don't process fill until feed time
''                          is later than fill time
'' 05/07/2014   DAJ         Sync position regardless in NewBarCheck and ReloadData if FractZen
'' 05/08/2014   DAJ         Handle 'Over Filled' orders as closed orders ( don't keep in SubmittedOrders collection )
'' 05/09/2014   DAJ         Do a NewBarCheck if the bars size changes OR the bar period just changed
'' 05/29/2014   DAJ         Fix for m.strActualBarPeriod not being set if not FractZen
'' 06/26/2014   DAJ         When position symbol changes and running on IB account, get contract info from IB
'' 06/26/2014   DAJ         If strategy doesn't have on-close orders, wait until after close and then
''                          just cancel working orders -- don't generate new orders or anything
'' 06/26/2014   DAJ         If running on a synthetic, add the non-synthetic equivalent to the secondary symbols
'' 08/19/2014   DAJ         Expose Strategy Basket Item Inputs
'' 08/25/2014   DAJ         Use IB3X enablement to control automated trading with a pyramiding
''                          strategy and only allow with non-live accounts for now
'' 08/27/2014   DAJ         Use the quantity in the next bar order report to fix pyramiding issues
'' 08/29/2014   DAJ         Don't include header when searching for last entry/exit date
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 09/08/2014   DAJ         Don't do anything in On-Close check if no On-Close orders in strategy
''                          Don't do New-Bar check if closest order is a Market order
'' 09/08/2014   DAJ         Delay On-Close check until market order fills
'' 09/09/2014   DAJ         Cancel all submitted orders after the market closes
'' 09/12/2014   DAJ         Fix for FillCheck when receiving partial fills while closing position
'' 09/24/2014   DAJ         Fix for issues with reversing position on new bar check and on-close check
'' 10/23/2014   DAJ         Set active flag to false when deleting to make sure num active display is correct
'' 10/24/2014   DAJ         Do fill check for first fill for new order when pyramiding
'' 10/27/2014   DAJ         Pass in fill to FillCheck instead of individual fields
'' 10/28/2014   DAJ         Pass order into the automated trading item for a FillCheck
'' 10/29/2014   DAJ         Compare price values instead of strings in SameOrder; Use quantity corrected
''                          engine string when comparing against working orders in SyncOrders
'' 10/29/2014   DAJ         Determine if we had a limit order fill at the high or low of the last bar
''                          and pass that over to the engine
'' 11/18/2014   DAJ         When comparing engine strings for orders, round the prices to nearest tick
'' 11/20/2014   DAJ         Walk through all submitted orders to look for market order in ClosestOrderIsMarketOrder
'' 11/21/2014   DAJ         Allow pyramiding for broker demo accounts as well
'' 11/24/2014   DAJ         Fixed logic error in allowing pyramiding
'' 12/02/2014   DAJ         Fix for display of strategy position when short
'' 12/04/2014   DAJ         Ignore engine generated exit orders if not in the appropriate position
'' 12/17/2014   DAJ         Don't call SubmitOneOrder if NewBarCheck coming; Don't generate orders
''                          if no data for current session yet; Include symbols in SameOrder check
'' 12/18/2014   DAJ         Return True from GenerateOrders if we decided not to generate orders
'' 02/18/2015   DAJ         Don't do SubmitOneOrder if Fill check coming; Do SubmitOneOrder if only one order
'' 02/19/2015   DAJ         Added and utilized the ChangePositionSymbol function
'' 02/26/2015   DAJ         Added and utilized the GetCurrentPositionFromBrokerInfo function
'' 03/03/2015   DAJ         Fix the strategy basket item ID if necessary
'' 04/01/2015   DAJ         Queue up submit one order call in NewBarCheck if one order or market order
'' 04/14/2015   DAJ         Clear the closing position flag at the close of the session
'' 04/28/2015   DAJ         Changed GenerateOrders to return an integer instead of a boolean; If GenerateOrders
''                          returns a 1 ( didn't generate orders ), do nothing
'' 05/13/2015   DAJ         Explicitly set new order status to parked in OnCloseCheck and RollContracts
'' 05/26/2015   DAJ         Fix for miscalculating on-close time for New Zealand
'' 07/10/2015   DAJ         Change a Stop-with-Limit order into a Limit order if the stop is on wrong side of the market
'' 07/15/2015   DAJ         Change a market order to a limit order for VX during ETH
'' 07/16/2015   DAJ         Change market orders if necessary in RollContracts and NewBarCheck
'' 07/31/2015   DAJ         Tweak for "JustChangedThisBar" fills for minute bars
'' 08/04/2015   DAJ         Fix On-Close time in ReloadData if it is now outside of market hours
'' 08/28/2015   DAJ         In ReloadData, if symbol changed but position is flat, cancel orders even no data for new session yet
'' 09/16/2015   DAJ         Change stop into market or stop-with-limit into limit if on the wrong side of the market when submit
'' 10/04/2015   DAJ         Added AllowPyramiding function; set up pyramid for live accounts with new enablement
'' 10/06/2015   DAJ         Always exit at end of day flag for automated trading
'' 10/08/2015   DAJ         Fix for OnClose check getting customer into a missed entry
'' 10/14/2015   DAJ         Set strategy position to flat when new exit at end of day code triggers to fix display
'' 10/22/2015   DAJ         Fix the exchange on-close time if negative
'' 11/13/2015   DAJ         Fix bug in SetQuantityOnEngineString where quantity on an exit is wrong if user changes quantity next entry
'' 11/16/2015   DAJ         Fix bug in SetQuantityOnEngineString where quantity is set to negative number on exit if position is short
'' 01/08/2016   DAJ         Fixed issues for running automated trading item on greater time periods than Daily
'' 01/12/2016   DAJ         Fix for On-Close checks still happening every day even on bar periods greater than daily
'' 01/21/2016   DAJ         Fix for enable doing an on-close running of orders incorrectly for bar periods greater than daily
'' 01/26/2016   DAJ         Reload data instead of just splicing bars on a new bar if bar period is greater than daily
'' 01/27/2016   DAJ         Fix for session date wrong on enable if after crossover
'' 01/29/2016   DAJ         Enter a position on a current bar check if bar period greater than daily and strategy position just changed
'' 02/03/2016   DAJ         Only update current session if the market is open
'' 02/03/2016   DAJ         Cancel all submitted orders if closing position because "ExitAtEndOfMarket"
'' 02/03/2016   DAJ         Store the reason for current bar check in member variable
'' 02/03/2016   DAJ         If an order gets rejected and they are in a position, run a current bar check to make sure the engine hasn't exited
'' 02/10/2016   DAJ         Update current position in fill check even if closing a position
'' 02/12/2016   DAJ         Store and utilize previous position as a class level variable
'' 03/01/2016   DAJ         Re-run engine on new bar check if mismatch because of limit order;
''                          Sanity checks for non-intraday bar periods; Don't allow order to cause position to exceed quantity next entry;
''                          Changed TradeSymbolOrId to use last open session; Tweaked logging for position symbol changing;
''                          Save the automated trading item if we fix the on-close time
'' 03/07/2016   DAJ         Fixes for sanity checks happening too often and happening even if item is not active
'' 03/10/2016   DAJ         Tweak to logging in GenerateOrders; Only DumpOrders if a flag file exists; Fix for entering a short position when pyramiding
'' 03/11/2016   DAJ         If a stop or stop-with-limit order gets rejected and this item is not in a position, run a current bar check;
''                          Change the position symbol in CurrentBarCheck if new daily session
'' 04/05/2016   DAJ         Set the previous position in the refresh position function
'' 04/07/2016   DAJ         Moved the checks around in CurrentBarCheck
'' 06/17/2016   DAJ         Don't try to do a current session bar check for a rejected order if that rejected order was a market order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDOrderFields
    eGDOrderField_Position = 1
    eGDOrderField_RuleID
    eGDOrderField_OrderType
    eGDOrderField_Price1
    eGDOrderField_Offset1
    eGDOrderField_Price2
    eGDOrderField_Offset2
    eGDOrderField_NumContracts
End Enum

Private Const kDiffPct As Double = 0.75

Private Type mPrivate
    ' Trading Item properties...
    lAutoTradeItemID As Long                ' Automated Trading Item ID
    lStrategyID As Long                     ' ID of the Strategy to run
    strStrategyName As String               ' Name of the Strategy to run
    lSymbolID As Long                       ' ID of the Symbol to run
    strSymbol As String                     ' Symbol to run the strategy on
    strBarPeriod As String                  ' Bar Period to load symbol
    strActualBarPeriod As String            ' Actual bar period to use for loading data
    lAccountID As Long                      ' ID of the Account to use
    lCurrentPosition As Long                ' Current Real Position
    lQtyNextEntry As Long                   ' Quantity of the next entry
    'dOnCloseTime As Double                  ' Local Time to check/submit On Close orders
    dOnCloseTimeExch As Double              ' Exchange Time to check/submit On Close orders
    bConfirm As Boolean                     ' Confirm all orders?
    lMinutesBefore As Long                  ' Minutes before close to check On Close orders
    bActive As Boolean                      ' Is this item active?
    strName As String                       ' Name of the auto trading item
    lConfirmTimeout As Long                 ' Number of seconds of timeout for confirmation box
    lAccountPositionID As Long              ' Account position ID
    bDeleted As Boolean                     ' Has this item been deleted?
    strPositionSymbol As String             ' Position symbol
    lPositionSymbolID As Long               ' ID of the position symbol
    nBroker As eTT_AccountType              ' Broker for the given account
    nBrokerStatus As eGDConnectionStatus    ' Connection status for the broker
    bExitAtEndOfDay As Boolean              ' Always exit at the end of the day?
    lSessionDate As Long                    ' Current session date
    
    lParentID As Long                       ' Parent ID ( -1 if parent, 0 if none, >0 if part of basket )
    lStrategyBasketID As Long               ' Strategy Basket ID
    lStrategyBasketItemID As Long           ' Strategy Basket Item ID
    lStrategyBasketItemMult As Long         ' Strategy Basket Item Multiplier
    dStrategyBasketLastModified As Double   ' Strategy Basket Last Modified
    strOverrides As String                  ' Parameter overrides
    strStrategyBasketItemKey As String      ' Strategy Basket Item Key
    
    ' Working variables...
    Strategy As cSystem                     ' Strategy object
    Bars As cGdBars                         ' Bars object for the symbol/period
    adLowestLimitBuy As cGdArray            ' Lowest real buy fill per bar
    adHighestLimitSell As cGdArray          ' Highest real sell fill per bar
    anLimitBuyAtLow As cGdArray             ' Is the lowest real buy limit fill at the low for the bar?
    anLimitSellAtHigh As cGdArray           ' Is the highest real sell limit fill at the high for the bar?
    lStrategyPos As Long                    ' Current Strategy Position
    lPrevStrategyPos As Long                ' Previous Strategy Position
    bIgnoreExits As Boolean                 ' Are we ignoring exits?
    astrOrders As cGdArray                  ' Array of orders
    bBetweenOpen As Boolean                 ' Are we in between On Close Time and Next Session Open?
    bBetweenClose As Boolean                ' Are we in between On Close Time and Session Close?
    bBetweenSessions As Boolean             ' Are we in between Session Close and Next Session Open?
    SubmittedOrders As cGdTree              ' Collection of currently submitted orders
    astrSecondaryMarkets As New cGdArray    ' Array of secondary markets
    SecondaryMarkets As New cGdTree         ' Collection of secondary market bars
    strConfirmedSubmit As String            ' Orders confirmed for submission
    lLDDwhenLoadBars As Long                ' Last Daily Download date when bars were loaded
    iRolling As Integer                     ' Are we currently rolling the position?
    bWaitingForData As Boolean              ' Are we waiting for data?
    bDataBehind As Boolean                  ' Is the data behind?
    bDoSubmitOneOrder As Boolean            ' Do we want to force a SubmitOneOrder call?
    strCurrentBarCheckReason As String      ' Reason for the current bar check
    lPreviousPosition As Long               ' Position after the last fill check that was done
    dNextSanityCheck As Double              ' Next time for a sanity check
    
    bClosingPosition As Boolean             ' Are we closing a position?
    lClosingQuantity As Long                ' Quantity submitted to close the position
    
    bBelievePosition As Boolean             ' Do we want to believe the position?
    AcctPos As cAccountPosition             ' Account position passed into the Enable function
    
    ' Boolean variables for alert checks...
    bOrderSubmitted As Boolean              ' Has a new order been submitted since last check?
    bOrderPriceHit As Boolean               ' Has an order price been hit since last check?
    bOrderFilled As Boolean                 ' Has an order been filled since last check?
    bOrderCancelled As Boolean              ' Has an order been cancelled?
    bOrderRejected As Boolean               ' Has an order been rejected?
End Type
Private m As mPrivate

Private Function OrderField(ByVal nField As eGDOrderFields) As Long
    OrderField = nField
End Function

Public Property Get AutoTradeItemID() As Long
    AutoTradeItemID = m.lAutoTradeItemID
End Property
Public Property Let AutoTradeItemID(ByVal lAutoTradeItemID As Long)
    m.lAutoTradeItemID = lAutoTradeItemID
End Property

Public Property Get StrategyID() As Long
    StrategyID = m.lStrategyID
End Property
Public Property Let StrategyID(ByVal lStrategyID As Long)
    m.lStrategyID = lStrategyID
End Property

Public Property Get StrategyName() As String
    StrategyName = m.strStrategyName
End Property
Public Property Let StrategyName(ByVal strStrategyName As String)
    m.strStrategyName = strStrategyName
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property
Public Property Let SymbolID(ByVal lSymbolID As Long)
    m.lSymbolID = lSymbolID
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Let Symbol(ByVal strSymbol As String)
    m.strSymbol = strSymbol
End Property

Public Property Get BarPeriod() As String
    BarPeriod = m.strBarPeriod
End Property
Public Property Let BarPeriod(ByVal strBarPeriod As String)
    m.strBarPeriod = strBarPeriod
    If IsFractZen = False Then
        m.strActualBarPeriod = m.strBarPeriod
    End If
End Property

Public Property Get ActualBarPeriod() As String
    If Not m.Bars Is Nothing Then
        ActualBarPeriod = m.Bars.Prop(eBARS_PeriodicityStr)
    Else
        ActualBarPeriod = ""
    End If
End Property

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal lAccountID As Long)
    m.lAccountID = lAccountID
    m.nBroker = g.Broker.AccountTypeForID(lAccountID)
End Property

Public Property Get CurrentPosition() As Long
    CurrentPosition = m.lCurrentPosition
End Property
Public Property Let CurrentPosition(ByVal lCurrentPosition As Long)
    m.lCurrentPosition = lCurrentPosition
    DumpDebug "Current Position set to " & Str(m.lCurrentPosition)
End Property

Public Property Get QtyNextEntry() As Long
    QtyNextEntry = m.lQtyNextEntry
End Property
Public Property Let QtyNextEntry(ByVal lQtyNextEntry As Long)
    m.lQtyNextEntry = lQtyNextEntry
End Property

Public Property Get OnCloseTime() As Double
On Error GoTo ErrSection:
    
    Dim dReturn As Double               ' Return value for the function
    
    dReturn = ConvertTimeZone(Date + m.dOnCloseTimeExch, m.Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
    dReturn = dReturn - Int(dReturn)
    
    OnCloseTime = dReturn
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAutoTradeItem.OnCloseTime.Get"
    
End Property
'Public Property Let OnCloseTime(ByVal dOnCloseTime As Double)
'    If dOnCloseTime <> m.dOnCloseTime Then
'        DumpDebug vbTab & "OnCloseTime changed from " & DateFormat(m.dOnCloseTime, NO_DATE, HH_MM_SS, AMPM_UPPER) & " to " & DateFormat(dOnCloseTime, NO_DATE, HH_MM_SS, AMPM_UPPER)
'        m.dOnCloseTime = dOnCloseTime
'    End If
'End Property

Public Property Get OnCloseTimeExch() As Double
    OnCloseTimeExch = m.dOnCloseTimeExch
End Property
Public Property Let OnCloseTimeExch(ByVal dOnCloseTimeExch As Double)
    ' 10/22/2015 DAJ: In some cases where the local machine time is in a foreign time zone, we are ending up with
    ' a negative exchange on-close time.  In these cases, we need to add one to get a positive number...
    If dOnCloseTimeExch < 0 Then
        dOnCloseTimeExch = dOnCloseTimeExch + 1#
    End If
    
    If dOnCloseTimeExch <> m.dOnCloseTimeExch Then
        DumpDebug vbTab & "OnCloseTimeExch changed from " & DateFormat(m.dOnCloseTimeExch, NO_DATE, HH_MM_SS, AMPM_UPPER) & " to " & DateFormat(dOnCloseTimeExch, NO_DATE, HH_MM_SS, AMPM_UPPER)
        m.dOnCloseTimeExch = dOnCloseTimeExch
    End If
End Property

Public Property Get MinutesBefore() As Long
    MinutesBefore = m.lMinutesBefore
End Property
Public Property Let MinutesBefore(ByVal lMinutesBefore As Long)
    m.lMinutesBefore = lMinutesBefore
End Property

Public Property Get ConfirmOrders() As Boolean
    ConfirmOrders = m.bConfirm
End Property
Public Property Let ConfirmOrders(ByVal bConfirmOrders As Boolean)
    m.bConfirm = bConfirmOrders
End Property

Public Property Get OrderSubmitted() As Boolean
    OrderSubmitted = m.bOrderSubmitted
End Property
Public Property Let OrderSubmitted(ByVal bOrderSubmitted As Boolean)
    m.bOrderSubmitted = bOrderSubmitted
End Property

Public Property Get OrderPriceHit() As Boolean
    OrderPriceHit = m.bOrderPriceHit
End Property
Public Property Let OrderPriceHit(ByVal bOrderPriceHit As Boolean)
    m.bOrderPriceHit = bOrderPriceHit
End Property

Public Property Get OrderFilled() As Boolean
    OrderFilled = m.bOrderFilled
End Property
Public Property Let OrderFilled(ByVal bOrderFilled As Boolean)
    m.bOrderFilled = bOrderFilled
End Property

Public Property Get OrderCancelled() As Boolean
    OrderCancelled = m.bOrderCancelled
End Property
Public Property Let OrderCancelled(ByVal bOrderCancelled As Boolean)
    m.bOrderCancelled = bOrderCancelled
End Property

Public Property Get OrderRejected() As Boolean
    OrderRejected = m.bOrderRejected
End Property
Public Property Let OrderRejected(ByVal bOrderRejected As Boolean)
    m.bOrderRejected = bOrderRejected
End Property

Public Property Get StrategyPos() As Long
    StrategyPos = m.lStrategyPos
End Property
Public Property Get StrategyPosString() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If m.lStrategyPos = 0 Then
        strReturn = ""
    ElseIf m.lStrategyPos = 1 Then
        strReturn = "Long"
    ElseIf m.lStrategyPos = -1 Then
        strReturn = "Short"
    ElseIf m.lStrategyPos > 0 Then
        strReturn = "Long " & Str(m.lStrategyPos)
    Else
        strReturn = "Short " & Str(Abs(m.lStrategyPos))
    End If
    
    StrategyPosString = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAutoTradeItem.StrategyPosString"
    
End Property

Public Property Get Name() As String
    Name = m.strName
End Property
Public Property Let Name(ByVal strName As String)
    m.strName = strName
End Property

Public Property Get ConfirmTimeout() As Long
    ConfirmTimeout = m.lConfirmTimeout
End Property
Public Property Let ConfirmTimeout(ByVal lConfirmTimeout As Long)
    m.lConfirmTimeout = lConfirmTimeout
End Property

Public Property Get Active() As Boolean
    Active = m.bActive
End Property
Private Property Let Active(ByVal bActive As Boolean)
    If bActive <> m.bActive Then
        m.bActive = bActive
        
        If m.lParentID > -1& Then
            If bActive Then
                g.TradingItems.NumActive = g.TradingItems.NumActive + 1
            Else
                g.TradingItems.NumActive = g.TradingItems.NumActive - 1
            End If
        End If
        
        ' 03/07/2016 DAJ: If we are going disabled, set the next sanity check to a large number to make sure it doesn't happen...
        If m.bActive = False Then
            NextSanityCheck = 99999#
        End If
    End If
End Property

Public Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID = 0 Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If
End Property
Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
    
    ' 08/13/2013 DAJ: If there is no data in the bars and the given symbol is different
    ' than the symbol in the bars, load the bar properties...
    If m.Bars.Size = 0 Then
        If (SymbolOrSymbolID <> m.Bars.SymbolOrSymbolID) Then
            SetBarProperties m.Bars, SymbolOrSymbolID
        End If
    End If
End Property

Public Property Get AccountPositionID() As Long
    AccountPositionID = m.lAccountPositionID
End Property
Public Property Let AccountPositionID(ByVal lAccountPositionID As Long)
    m.lAccountPositionID = lAccountPositionID
End Property

Public Property Get Deleted() As Boolean
    Deleted = m.bDeleted
End Property
Public Property Let Deleted(ByVal bDeleted As Boolean)
    m.bDeleted = bDeleted
End Property

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property

Public Property Get ParentID() As Long
    ParentID = m.lParentID
End Property
Public Property Let ParentID(ByVal lParentID As Long)
    m.lParentID = lParentID
End Property

Public Property Get StrategyBasketID() As Long
    StrategyBasketID = m.lStrategyBasketID
End Property
Public Property Let StrategyBasketID(ByVal lStrategyBasketID As Long)
    m.lStrategyBasketID = lStrategyBasketID
End Property

Public Property Get StrategyBasketItemID() As Long
    StrategyBasketItemID = m.lStrategyBasketItemID
End Property
Public Property Let StrategyBasketItemID(ByVal lStrategyBasketItemID As Long)
    m.lStrategyBasketItemID = lStrategyBasketItemID
End Property

Public Property Get StrategyBasketItemMult() As Long
    StrategyBasketItemMult = m.lStrategyBasketItemMult
End Property
Public Property Let StrategyBasketItemMult(ByVal lStrategyBasketItemMult As Long)
    m.lStrategyBasketItemMult = lStrategyBasketItemMult
End Property

Public Property Get StrategyBasketLastModified() As Double
    StrategyBasketLastModified = m.dStrategyBasketLastModified
End Property
Public Property Let StrategyBasketLastModified(ByVal dStrategyBasketLastModified As Double)
    m.dStrategyBasketLastModified = dStrategyBasketLastModified
End Property

Public Property Get Overrides() As String
    Overrides = m.strOverrides
End Property
Public Property Let Overrides(ByVal strOverrides As String)
    m.strOverrides = strOverrides
End Property

Public Property Get StrategyBasketItemKey() As String
    StrategyBasketItemKey = m.strStrategyBasketItemKey
End Property
Public Property Let StrategyBasketItemKey(ByVal strStrategyBasketItemKey As String)
    m.strStrategyBasketItemKey = strStrategyBasketItemKey
End Property

Public Property Get PositionSymbolOrID() As Variant
    If m.lPositionSymbolID = 0& Then
        PositionSymbolOrID = m.strPositionSymbol
    Else
        PositionSymbolOrID = m.lPositionSymbolID
    End If
End Property
Private Property Let PositionSymbolOrID(ByVal vPositionSymbolOrID As Variant)
On Error GoTo ErrSection:

    Dim strPositionSymbol As String     ' Position symbol

    m.lPositionSymbolID = GetSymbolID(vPositionSymbolOrID)
    strPositionSymbol = GetSymbol(vPositionSymbolOrID)
    If strPositionSymbol <> m.strPositionSymbol Then
        DumpDebug vbTab & vbTab & "Position Symbol changing from '" & m.strPositionSymbol & "' to '" & strPositionSymbol & "'"
        m.strPositionSymbol = strPositionSymbol
        g.Broker.GetContractInfo m.nBroker, m.strPositionSymbol
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAutoTradeItem.PositionSymbolOrID.Let"

End Property

Public Property Get TradeSymbolOrID() As Variant
On Error GoTo ErrSection:

    Dim dNow As Double                  ' Current time in exchange time
    Dim lSessionDate As Long            ' Last open session for the symbol
    
    dNow = CurrentTime(m.Bars.Prop(eBARS_ExchangeTimeZoneInf), m.strSymbol, True)
    lSessionDate = m.Bars.LastOpenSessionForTradeTime(dNow)
    
    TradeSymbolOrID = ConvertToTradeSymbol(SymbolOrSymbolID, lSessionDate)

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAutoTradeItem.TradeSymbolOrID.Get"
    
End Property

Public Property Get WaitingForData() As Boolean
    WaitingForData = m.bWaitingForData
End Property
Private Property Let WaitingForData(ByVal bWaitingForData As Boolean)
    m.bWaitingForData = bWaitingForData
End Property

Public Property Get DataBehind() As Boolean
    DataBehind = m.bDataBehind
End Property
Private Property Let DataBehind(ByVal bDataBehind As Boolean)
    If bDataBehind <> m.bDataBehind Then
        m.bDataBehind = bDataBehind
        RefreshUI
    End If
End Property

Private Property Get ClosingPosition() As Boolean
    ClosingPosition = m.bClosingPosition
End Property
Private Property Let ClosingPosition(ByVal bClosingPosition As Boolean)
    If m.bClosingPosition <> bClosingPosition Then
        m.bClosingPosition = bClosingPosition
        DumpDebug vbTab & vbTab & "Closing Position changed to " & Str(m.bClosingPosition)
    End If
End Property

Private Property Get ClosingQuantity() As Long
    ClosingQuantity = m.lClosingQuantity
End Property
Private Property Let ClosingQuantity(ByVal lClosingQuantity As Long)
    If m.lClosingQuantity <> lClosingQuantity Then
        m.lClosingQuantity = lClosingQuantity
        DumpDebug vbTab & vbTab & "Closing Quantity changed to " & Str(m.lClosingQuantity)
    End If
End Property

Public Property Get Pyramid() As Boolean
    Pyramid = m.Strategy.Pyramid
End Property

Private Property Get IsFractZen() As Boolean
    IsFractZen = (UCase(m.strBarPeriod) = "AUTO BREAKOUT" Or UCase(m.strBarPeriod) = "FRACTZEN")
End Property

Private Property Get DoSubmitOneOrder() As Boolean
    DoSubmitOneOrder = m.bDoSubmitOneOrder
End Property
Private Property Let DoSubmitOneOrder(ByVal bDoSubmitOneOrder As Boolean)
    If bDoSubmitOneOrder <> m.bDoSubmitOneOrder Then
        DumpDebug vbTab & vbTab & "DoSubmitOneOrder changed from " & Str(m.bDoSubmitOneOrder) & " to " & Str(bDoSubmitOneOrder)
        m.bDoSubmitOneOrder = bDoSubmitOneOrder
    End If
End Property

Public Property Get ExitAtEndOfDay() As Boolean
    ExitAtEndOfDay = m.bExitAtEndOfDay
End Property
Public Property Let ExitAtEndOfDay(ByVal bExitAtEndOfDay As Boolean)
    If bExitAtEndOfDay <> m.bExitAtEndOfDay Then
        DumpDebug vbTab & vbTab & "Changing ExitAtEndOfDay from " & Str(m.bExitAtEndOfDay) & " to " & Str(bExitAtEndOfDay)
        m.bExitAtEndOfDay = bExitAtEndOfDay
    End If
End Property

Private Property Get SessionDate() As Long
    SessionDate = m.lSessionDate
End Property
Private Property Let SessionDate(ByVal lSessionDate As Long)
    If lSessionDate <> m.lSessionDate Then
        DumpDebug vbTab & vbTab & "Session Date changing from " & Str(m.lSessionDate) & " ( " & DateFormat(m.lSessionDate, MM_DD_YYYY) & " ) to " & Str(lSessionDate) & " ( " & DateFormat(lSessionDate, MM_DD_YYYY) & " )"
        m.lSessionDate = lSessionDate
    End If
End Property

Public Property Get CurrentBarCheckReason() As String
    CurrentBarCheckReason = m.strCurrentBarCheckReason
End Property
Public Property Let CurrentBarCheckReason(ByVal strCurrentBarCheckReason As String)
    m.strCurrentBarCheckReason = strCurrentBarCheckReason
End Property

Private Property Get PreviousPosition() As Long
    PreviousPosition = m.lPreviousPosition
End Property
Private Property Let PreviousPosition(ByVal lPreviousPosition As Long)
    If lPreviousPosition <> m.lPreviousPosition Then
        DumpDebug vbTab & vbTab & "Previous Position changing from " & Str(m.lPreviousPosition) & " to " & Str(lPreviousPosition)
        m.lPreviousPosition = lPreviousPosition
    End If
End Property

Private Property Get NextSanityCheck() As Double
    NextSanityCheck = m.dNextSanityCheck
End Property
Private Property Let NextSanityCheck(ByVal dNextSanityCheck As Double)
    If dNextSanityCheck <> m.dNextSanityCheck Then
        DumpDebug vbTab & vbTab & "Next Santity check changing from " & DateFormat(m.dNextSanityCheck, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & " to " & DateFormat(dNextSanityCheck, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
        m.dNextSanityCheck = dNextSanityCheck
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the automated trading item with the given ID
'' Inputs:      Trade Item ID, Load Bars?, Recordset
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal lAutoTradeItemID As Long, Optional ByVal bLoadBars As Boolean = True, Optional ByVal rs As Recordset = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Bars As New cGdBars             ' Temporary bars structure
    Dim AcctPos As cAccountPosition     ' Account position object
    Dim bResave As Boolean              ' Re-Save the automated trading item?
    Dim strTradeSymbol As String        ' Trade symbol
    Dim lStrategyBasketItemID As Long   ' Strategy basket item ID
    
    bResave = False
    bReturn = False
    
    If rs Is Nothing Then
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblAutoTradingItem] WHERE [TradingItemID]=" & Str(lAutoTradeItemID) & ";", dbOpenDynaset)
    End If
    
    If Not (rs.BOF And rs.EOF) Then
        m.lAutoTradeItemID = rs!TradingItemID
        m.lStrategyID = rs!StrategyID
        m.strStrategyName = rs!StrategyName
        m.lSymbolID = rs!SymbolID
        m.strSymbol = rs!Symbol
        BarPeriod = rs!Period
        AccountID = rs!AccountID
        m.lCurrentPosition = rs!CurrentPosition
        m.lAccountPositionID = NullChk(rs!AccountPositionID, 0&)
        m.bDeleted = NullChk(rs!Deleted, False)
        m.bExitAtEndOfDay = NullChk(rs!ExitAtEndOfDay, False)
        
        m.lParentID = NullChk(rs!ParentID, 0&)
        m.lStrategyBasketID = NullChk(rs!StrategyBasketID, 0&)
        m.lStrategyBasketItemID = NullChk(rs!StrategyBasketItemID, 0&)
        m.lStrategyBasketItemMult = NullChk(rs!StrategyBasketItemMult, 1&)
        m.dStrategyBasketLastModified = NullChk(rs!StrategyBasketLastModified, 0#)
        m.strOverrides = NullChk(rs!Overrides, "")
        m.strStrategyBasketItemKey = NullChk(rs!StrategyBasketItemKey, "")
        lStrategyBasketItemID = mSysNav.StrategyBasketItemIdForKey(m.lStrategyBasketID, m.strStrategyBasketItemKey)
        
        PositionSymbolOrID = TradeSymbolOrID
        
        DumpDebug "Loading"
        DumpDebug vbTab & "Current Position = " & Str(m.lCurrentPosition)
        DumpDebug vbTab & "Quantity Next Entry = " & Str(rs!QtyNextEntry)
        
        m.lQtyNextEntry = rs!QtyNextEntry
        'm.dOnCloseTime = rs!OnCloseTime / 1440#
        m.dOnCloseTimeExch = NullChk(rs!OnCloseTimeExch, 0#)
        m.lMinutesBefore = rs!MinutesBeforeClose
        m.bConfirm = rs!ConfirmOrders
        m.strName = NullChk(rs!Name, "")
        m.lConfirmTimeout = NullChk(rs!ConfirmTimeout, 15&)
        
        m.Strategy.Overrides = CombineOverrides(m.strOverrides)
        m.Strategy.Load m.lStrategyID
        
        ' 06/03/2013 DAJ: Do this here so that the Secondary Markets call to the strategy
        ' will set the symbol correctly on 'Weekly', 'Monthly', etc....
        SetBarProperties m.Bars, SymbolOrSymbolID
        m.Bars.Prop(eBARS_PeriodicityStr) = m.strActualBarPeriod
        m.Strategy.Bars = m.Bars.MakeCopy
        
        Set m.astrSecondaryMarkets = m.Strategy.SecondaryMarkets(m.strActualBarPeriod)
        If m.strBarPeriod = "Daily" Then
            If AddMinuteData Then
                DumpDebug vbTab & "60 minute bars added to secondary markets"
            End If
        End If
        
        ' 06/26/2014 DAJ: If they are running the automated trading item on a synthetic
        ' future contract, add daily bars for the non-synthetic equivalent to the secondary
        ' markets to make sure it is in the stream...
        If InStr(m.strSymbol, "-") <> 0 Then
            strTradeSymbol = ConvertToTradeSymbol(m.strSymbol)
            If strTradeSymbol <> m.strSymbol Then
                If Parse(strTradeSymbol, "-", 1) <> Parse(m.strSymbol, "-", 1) Then
                    m.astrSecondaryMarkets.Add strTradeSymbol & ";Daily"
                    DumpDebug vbTab & strTradeSymbol & ";Daily added to secondary markets"
                End If
            End If
        End If
        
        If bLoadBars Then
            LoadBars
            If m.Bars.Size > 0 Then m.Strategy.FromDate = m.Bars(eBARS_DateTime, 0)
            m.Strategy.ToEndOfData = True
        'Else
            ' 06/03/2013 DAJ: Don't need to do this here because it is being done above...
            'SetBarProperties m.Bars, SymbolOrSymbolID
        End If
        
        If (m.lAccountPositionID = 0&) And (m.lParentID <> -1&) Then
            m.lAccountPositionID = g.Broker.CreateFillSummaryForAutoTrade(Me)
            bResave = True
        End If
        
        ' 12/14/2012 DAJ: If the On Close time isn't set in exchange time yet, convert
        ' the local time one that we have stored...
        If m.dOnCloseTimeExch = 0# Then
            m.dOnCloseTimeExch = ConvertTimeZone(rs!OnCloseTime / 1440#, "", m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
            m.dOnCloseTimeExch = m.dOnCloseTimeExch - Int(m.dOnCloseTimeExch)
            bResave = True
        End If
        
        ' 10/22/2015 DAJ: In some cases where the local machine time is in a foreign time zone, we are ending up with
        ' a negative exchange on-close time.  In these cases, we need to add one to get a positive number...
        If m.dOnCloseTimeExch < 0# Then
            m.dOnCloseTimeExch = m.dOnCloseTimeExch + 1#
            DumpDebug vbTab & "Exchange On-Close time fixed.  Changed from " & Str(m.dOnCloseTimeExch - 1#) & " to " & Str(m.dOnCloseTimeExch) & " ( " & DateFormat(m.dOnCloseTimeExch, NO_DATE, HH_MM_SS, AMPM_UPPER) & " )"
            bResave = True
        End If
        
        ' 03/03/2015 DAJ: When you re-import a library with strategy baskets, the strategy basket
        ' item ID gets changed, but we weren't changing it here.  We will start to do that on an import
        ' now, but we need to do this here to fix any items that got screwed up before the fix...
        If lStrategyBasketItemID <> m.lStrategyBasketItemID Then
            DumpDebug "StrategyBasketItemID changed from " & Str(m.lStrategyBasketItemID) & " to " & Str(lStrategyBasketItemID)
            m.lStrategyBasketItemID = lStrategyBasketItemID
            bResave = True
        End If
        
        If bResave Then
            Save
        End If
        
        m.nBrokerStatus = g.Broker.ConnectionStatusForAccount(m.lAccountID)
        DumpDebug "Connection status for '" & g.Broker.BrokerName(m.nBroker) & "' account '" & g.Broker.AccountNumberForID(m.lAccountID) & "' is '" & g.BrokerEnums.ConnectionStatusToString(m.nBrokerStatus) & "'"
        
        bReturn = True
    End If
    
    Load = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the automated trading item to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblAutoTradingItem] " & _
                "WHERE [TradingItemID]=" & Str(m.lAutoTradeItemID) & ";", dbOpenDynaset)
    If (rs.BOF And rs.EOF) Then
        rs.AddNew
    Else
        rs.Edit
    End If
    
    DumpDebug "Saving " & Description
        
    rs!StrategyID = m.lStrategyID
    rs!StrategyName = m.strStrategyName
    rs!SymbolID = m.lSymbolID
    rs!Symbol = m.strSymbol
    rs!Period = m.strBarPeriod
    rs!AccountID = m.lAccountID
    rs!CurrentPosition = m.lCurrentPosition
    rs!QtyNextEntry = m.lQtyNextEntry
    rs!OnCloseTime = Me.OnCloseTime * 1440 ' m.dOnCloseTime * 1440
    rs!OnCloseTimeExch = m.dOnCloseTimeExch
    rs!MinutesBeforeClose = m.lMinutesBefore
    rs!ConfirmOrders = m.bConfirm
    rs!Name = m.strName
    rs!ConfirmTimeout = m.lConfirmTimeout
    
    m.lAutoTradeItemID = rs!TradingItemID
    If (m.lAccountPositionID = 0&) And (m.lParentID <> -1&) Then
        m.lAccountPositionID = g.Broker.CreateFillSummaryForAutoTrade(Me)
    End If
    rs!AccountPositionID = m.lAccountPositionID
    rs!Deleted = m.bDeleted
    
    rs!ParentID = m.lParentID
    rs!StrategyBasketID = m.lStrategyBasketID
    rs!StrategyBasketItemID = m.lStrategyBasketItemID
    rs!StrategyBasketItemMult = m.lStrategyBasketItemMult
    rs!StrategyBasketLastModified = m.dStrategyBasketLastModified
    rs!Overrides = m.strOverrides
    rs!StrategyBasketItemKey = m.strStrategyBasketItemKey
    rs!ExitAtEndOfDay = m.bExitAtEndOfDay
    
    rs.Update
                
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Delete
'' Description: Delete this item out of the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete()
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position of the pipe in the string

    Do
        m.strName = m.strName & "|"
        
        If Len(m.strName) > 50 Then
            lPos = InStr(m.strName, "|")
            m.strName = Left(m.strName, lPos - 2) & Mid(m.strName, lPos)
        End If
    Loop While g.TradingItems.NameExistsInDb(m.strName, m.lAutoTradeItemID)
    
    Active = False
    m.bDeleted = True
    Save

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.Delete"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Description
'' Description: Create a comma delimited descriptor for the item
'' Inputs:      Include Account?
'' Returns:     Description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Description(Optional ByVal bIncludeAccount As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Description to return from the function

    strReturn = "'" & m.strStrategyName & "', '" & m.strSymbol & "', '" & m.strBarPeriod & "'"
    If bIncludeAccount Then
        strReturn = strReturn & ", " & g.Broker.AccountNumberForID(m.lAccountID)
    End If
    
    Description = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.Description"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshUI
'' Description: Refresh the user interface for the trade item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshUI()
On Error GoTo ErrSection:

    If FormIsLoaded("frmTradeItems") Then
        frmTradeItems.RefreshTradeItem m.lAutoTradeItemID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.RefreshUI"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextBarReport
'' Description: Perform a next bar report for the local strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub NextBarReport()
On Error GoTo ErrSection:

    Dim Basket As cStrategyBasket       ' Strategy basket object

    If m.lParentID < 0& Then
        Set Basket = New cStrategyBasket
        If Basket.LoadDb(m.lStrategyBasketID) Then
            Basket.Run True
        End If
    Else
        ' If there is no data loaded, try to load it here...
        If m.Bars.Size = 0 Then
            LoadBars
        End If
    
        ' If there still is no data, then display an error...
        If m.Bars.Size = 0 Then
            InfBox "Cannot load " & m.strBarPeriod & " data for " & m.strSymbol, "!", , "Automated Trading Error"
        Else
            m.Strategy.Bars = m.Bars.MakeCopy
            m.Strategy.NextBarReport
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.NextBarReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PerformanceReport
'' Description: Perform a performance report for the local strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PerformanceReport()
On Error GoTo ErrSection:

    Dim Basket As cStrategyBasket       ' Strategy basket object

    If m.lParentID < 0& Then
        Set Basket = New cStrategyBasket
        If Basket.LoadDb(m.lStrategyBasketID) Then
            Basket.Run False, , m.lAutoTradeItemID
        End If
    Else
        ' If there is no data loaded, try to load it here...
        If m.Bars.Size = 0 Then
            LoadBars
        End If
    
        ' If there still is no data, then display an error...
        If m.Bars.Size = 0 Then
            InfBox "Cannot load " & m.strBarPeriod & " data for " & m.strSymbol, "!", , "Automated Trading Error"
        Else
            If m.Strategy.SystemNumber = 0 Then
                'm.Strategy.Overrides = CombineOverrides(m.strOverrides)
                m.Strategy.Load m.lStrategyID
            End If
            m.Strategy.Overrides = CombineOverrides(m.strOverrides)
            m.Strategy.Bars = m.Bars.MakeCopy
            m.Strategy.Test
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.PerformanceReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBars
'' Description: Do realtime updates on the local Bars
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateBars()
On Error GoTo ErrSection:

    Dim bNewBar As Boolean              ' Have we hit a new bar?
    Dim dNow As Double                  ' Current time
    Dim lLastOnCloseCheck As Long       ' Last date we did an on-close check
    Dim lIndex As Long                  ' Index into a for loop
    Dim dDateTimeLastBar As Double      ' Date and time of the last bar
    Dim lBarSize As Long                ' Size of the bars before updating
    Dim strPeriod As String             ' New bar period for FractZen bars
    Dim bLoadBars As Boolean            ' Reload the bars?
    Dim lSessionDate As Long            ' Session date
    Dim bPeriodChanged As Boolean       ' Did the bar period just change?
    Dim bDoOnClose As Boolean           ' Do an on-close check?
    Dim bNewBarCheck As Boolean         ' Was a new-bar check queued?
    Dim bNewSession As Boolean          ' Are we in a new session?
    Dim bDoingCheck As Boolean          ' Are we doing some sort of check?
    Static bDumped As Boolean           ' Did we dump a line to the log?
    
    bDoingCheck = False
    
    If WaitingForData Then
        If AllRtDataAvailable Then
            ' DAJ 08/15/2012: Need to reload the bars now that we have all of the data.  Otherwise
            ' the first new tick that comes in after the Enable will cause a new bar and subsequently
            ' a new bar check which we don't necessarily want yet (unless it truly is a new bar)...
            LoadBars
            Enable m.bBelievePosition, m.AcctPos
        End If
    Else
        For lIndex = 1 To m.SecondaryMarkets.Count
            If g.RealTime.UpdateBars(m.SecondaryMarkets(lIndex), bNewBar) Then
                If bNewBar Then
                    g.RealTime.SpliceBars m.SecondaryMarkets(lIndex)
                End If
            End If
        Next lIndex
        
        CalcBetween dNow
        
        bNewBarCheck = False
        bNewSession = False
        dDateTimeLastBar = m.Bars(eBARS_DateTime, m.Bars.Size - 1)
        lBarSize = m.Bars.Size
        
        If g.RealTime.UpdateBars(m.Bars, bNewBar) = True Then
            If bNewBar Then
                bLoadBars = False
                bPeriodChanged = False
                
                If (LastDailyDownload <> m.lLDDwhenLoadBars) Or (FileExist(App.Path & "\OldAutoTrade.flg") = True) Then
                    bLoadBars = True
                ElseIf m.Bars.Prop(eBARS_Periodicity) > ePRD_Days + 1 Then
                    ' DAJ 01/26/2016: Still having issues where when a new bar starts for a weekly bar on Sunday or
                    ' Monday, the last bar is getting Monday's date instead of Friday's.  In looking at the chart and
                    ' the quote board ( which appear to be right on Monday ), they are re-loading data instead of just
                    ' a splice bars if the bar period is greater than Daily, so we will try to do that here as well...
                    bLoadBars = True
                ElseIf (m.Bars.Size > 1) And (IsFractZen = True) Then
                    strPeriod = g.FractZen.GetFractZenPeriod(m.strSymbol)
                    If strPeriod <> m.strActualBarPeriod Then
                        DumpDebug "FractZen period changed from '" & m.strActualBarPeriod & "' to '" & strPeriod & "'"
                        m.strActualBarPeriod = strPeriod
                        bLoadBars = True
                        bPeriodChanged = True
                    Else
                        DumpDebug "FractZen period stayed at '" & m.strActualBarPeriod & "'"
                    End If
                End If
                    
                ' if have done daily download since bars were loaded or the FractZen bar period changed, then
                ' reload entire data
                If bLoadBars Then
                    LoadBars
                Else ' otherwise just need to reload data since the last daily download
                    g.RealTime.SpliceBars m.Bars
                    
                    m.adLowestLimitBuy.Size = m.Bars.Size
                    m.adHighestLimitSell.Size = m.Bars.Size
                    m.anLimitBuyAtLow.Size = m.Bars.Size
                    m.anLimitSellAtHigh.Size = m.Bars.Size
                End If
                
                ' Do this in here instead of down below, because the time of the last bar is constantly
                ' changing for tick-per-bar, volume-per-bar, etc. type of data (01/30/2008 DAJ)...
                ' To be a little safer, also ensure that the size of the bars is different (in case the
                ' code to estimate if there is a new bar is ever wrong) (01/30/2008 DAJ)...
                ' Since the FractZen period changing could coincidentally change the bars to be the same
                ' size as they were, we need to run the NewBarCheck if the period just changed
                ' as well ( 05/09/2014 DAJ )...
                If (Active = True) And ((m.Bars.Size <> lBarSize) Or (bPeriodChanged = True)) Then
                    DumpDebug "NewBarCheck added to Queue: Active=" & Str(Active) & ", Bars DateTime=" & Str(m.Bars(eBARS_DateTime, m.Bars.Size - 1)) & ", TimeLastBar=" & Str(dDateTimeLastBar)
                    'NewBarCheck
                    g.TradingItems.AddNewBarCheck m.lAutoTradeItemID
                    bNewBarCheck = True
                    bDoingCheck = True
                End If
            End If
'            If Active And (m.Bars(eBARS_DateTime, m.Bars.Size - 1) <> dDateTimeLastBar) Then
'                DumpDebug "Calling NewBarCheck: Active=" & Str(Active) & ", Bars DateTime=" & Str(m.Bars(eBARS_DateTime, m.Bars.Size - 1)) & ", TimeLastBar=" & Str(dDateTimeLastBar)
'                NewBarCheck
'            End If
            
            UpdateLimitArrays m.Bars.Size - 1, True
            
            ' DAJ 02/03/2016: Brady ran into an issue running on weekly bars on a synthetic symbol where this code happened while the market was
            ' closed and didn't get him into a position like he wanted.  It turns out that g.RealTime.UpdateBars will return a True if anything
            ' changes ( including bid, ask, etc. ), so this is being hit even if we don't have a tick.  I am changing it to only happen if the
            ' market is open since this is meant kind of like a new bar check...
            If MarketIsOpen Then
                lSessionDate = m.Bars.SessionDateForTradeTime(dNow)
                If lSessionDate <> m.lSessionDate Then
                    SessionDate = lSessionDate
                    bNewSession = True
                End If
            End If
            
            ' DAJ 12/15/2014: A customer ran into a situation where we submitted an order after calling
            ' for a new bar check.  The new bar check then tried to cancel the order because it shouldn't
            ' be there anymore, but since it was pending, it didn't successfully cancel it and it ended up
            ' filling later.  Because of that, I don't think we want to do anything here if a new bar
            ' check has been queued up...
            ' DAJ 02/17/2015: Michael Cook ran into a situation where we tried to submit the closest order
            ' after a FillCheck had been queued, but before it happened.  Since the FillCheck will
            ' generate orders, don't do the SubmitOneOrder here if there is a FillCheck queued up, either...
            If (bNewBarCheck = False) And (g.TradingItems.FillCheckQueued(m.lAutoTradeItemID) = False) Then
                ' DAJ 01/07/2016: If the user is running this automated trading item on a bar periodicity greater then daily, then
                ' we need to re-run the engine when a new session that does not generate a new bar starts to re-generate orders that
                ' expired at the end of the previous session...
                If (bNewSession = True) And (m.Bars.Prop(eBARS_Periodicity) > ePRD_Days + 1) Then
                    CurrentBarCheckReason = "NewDailySession"
                    g.TradingItems.AddCurrentBarCheck m.lAutoTradeItemID
                    bDoingCheck = True
                Else
                    SubmitOneOrder
                End If
            End If
            m.bOrderPriceHit = g.Broker.OrderPriceHitForAutoTrade(Me, m.Bars(eBARS_Close, m.Bars.Size - 1))
        ElseIf DoSubmitOneOrder = True Then
            SubmitOneOrder
        End If
        
        If Active Then
            lLastOnCloseCheck = GetIniFileProperty("OnClose" & Str(m.lAutoTradeItemID), 0&, "AutoTrade", g.strIniFile)
            bDoOnClose = False
            
            If (m.Strategy.HasOnCloseRules = True) Or (ExitAtEndOfDay = True) Then
                ' If we are between the on-close time and the next start time AND we have not done the
                ' OnCloseCheck today yet AND it is not a weekend, then do the OnCloseCheck...
                If (m.bBetweenClose = True) And (Int(dNow) > lLastOnCloseCheck) Then
                    If IsWeekday(dNow) Then
                        bDoOnClose = True
                    End If
                End If
            Else
                If (m.bBetweenSessions = True) And (Int(dNow) > lLastOnCloseCheck) Then
                    If IsWeekday(dNow) Then
                        bDoOnClose = True
                    End If
                End If
            End If
            
            ' DAJ 09/08/2014: If the closest order to the market is a market order, we need to
            ' delay the On-Close check until it fills so that we do the right thing...
            If (bDoOnClose = True) Then
                If (ClosestOrderIsMarketOrder = False) Then
                    DumpDebug "OnCloseCheck added to Queue"
                    'OnCloseCheck
                    g.TradingItems.AddOnCloseCheck m.lAutoTradeItemID
                    SetIniFileProperty "OnClose" & Str(m.lAutoTradeItemID), Int(dNow), "AutoTrade", g.strIniFile
                    bDumped = False
                    bDoingCheck = True
                ElseIf bDumped = False Then
                    DumpDebug "On-Close check delayed until market order fills"
                    bDumped = True
                End If
            End If
            
#If 0 Then
            If (dNow > dOnClose) And (Int(dNow) > lLastOnCloseCheck) Then
                OnCloseCheck
                SetIniFileProperty "OnClose" & Str(m.lAutoTradeItemID), Int(dNow), "AutoTrade", g.strIniFile
            End If
#End If
    
            ' 03/01/2016 DAJ: If we aren't doing some other sort of check and we are on or after the next scheduled
            ' sanity check, then if the market is open do the next sanity check, otherwise bump it by an hour...
            ' 03/07/2016 DAJ: Moved this block up into the "If Active" block so that we don't do it if the automated
            ' trading item is not active...
            If bDoingCheck = False Then
                If dNow > NextSanityCheck Then
                    If MarketIsOpen Then
                        CurrentBarCheckReason = "SanityCheck"
                        g.TradingItems.AddCurrentBarCheck m.lAutoTradeItemID
                        bDoingCheck = True
                    Else
                        NextSanityCheck = NextSanityCheck + (1# / 24#)
                    End If
                End If
            End If
        End If
    End If
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.UpdateBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnterPosition
'' Description: Submit an order to enter into the current strategy position
'' Inputs:      Wait for it?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EnterPosition(Optional ByVal bWait As Boolean = False, Optional ByVal bAllowReverse As Boolean = False)
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Temporary order object
    Dim lTimeOut As Long                ' Timeout
    Dim strBrokerStatus As String       ' Current status of the broker
    Dim lOrderQuantity As Long          ' Quantity for the order
    
    If CanPerformOrderAction(strBrokerStatus) = False Then
        DumpDebug "Entering Position: Could not enter a position because: " & strBrokerStatus
        InfBox "You cannot enter a position in this automated trading item because you are not currently connected to the broker", "!", , "Automated Trading Error"
    
    ElseIf ((m.lCurrentPosition = 0&) Or (bAllowReverse = True)) And (m.lStrategyPos <> 0&) And (m.lQtyNextEntry > 0&) Then
        DumpDebug "Entering Position: Current Position = " & Str(m.lCurrentPosition) & ", Strategy Position = " & Str(m.lStrategyPos) & ", Quantity Next Entry = " & Str(m.lQtyNextEntry)
        
        If m.Strategy.Pyramid Then
            lOrderQuantity = Abs(m.lCurrentPosition) + (Abs(m.lQtyNextEntry) * Abs(m.lStrategyPos))
        Else
            lOrderQuantity = Abs(m.lCurrentPosition) + Abs(m.lQtyNextEntry)
        End If
        
        Set Order = New cPtOrder
        With Order
            .AccountID = m.lAccountID
            .AutoTradeItemID = m.lAutoTradeItemID
            .Buy = (m.lStrategyPos > 0)
            .Enter = True
            .LimitPrice = kNullData
            .OrderDate = .BrokerDate(CurrentTime("", m.strSymbol, True))
            .OrderType = eTT_OrderType_Market
            .Expiration = ExpirationDate * -1&
            .Quantity = lOrderQuantity
            .StopPrice = kNullData
            .SymbolOrSymbolID = TradeSymbolOrID
            .IsAutomated = True
            
            .Save
        End With
        
        PlaceOrder Order, True, "to enter position"
    
        If g.Broker.AutoJournalAutomated Then
            g.JournalBridge.CreateOrderJournal Order.OrderID, CurrentTime, "Automated Trading Strategy (" & m.strName & ") placing this order to enter into a position"
        End If
            
        If bWait Then
            lTimeOut = 0&
            Do While (Abs(m.lCurrentPosition) <> lOrderQuantity) And (lTimeOut < 30&)
                Sleep 1
                lTimeOut = lTimeOut + 1&
            Loop
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.EnterPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeEntryQuantity
'' Description: Change the entry quantity for any existing orders
'' Inputs:      Old Quantity, New Quantity
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeEntryQuantity(ByVal lOldQty As Long, ByVal lNewQty As Long)
On Error GoTo ErrSection:

    DumpDebug "Changing entry quantity from " & Str(lOldQty) & " to " & Str(lNewQty)
    
    m.lQtyNextEntry = lNewQty
    
    If Not m.bBetweenOpen Then
        SetOrderQuantities
        SyncOrders "ChangeEntryQuantity"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.ChangeEntryQuantity"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Enable
'' Description: Enable the trading item
'' Inputs:      Believe Position?, Account Position
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Enable(Optional ByVal bBelieveCurrentPosition As Boolean = False, Optional ByVal AcctPos As cAccountPosition = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As New cPtOrder           ' Temporary order object
    Dim lCurPos As Long                 ' Current position in the database
    Dim Fill As New cPtFill             ' Temporary fill object
    Dim strSymbol As String             ' Symbol to pass for the current position
    Dim Orders As cPtOrders             ' Collection of orders
    Dim iGoodOrders As Integer          ' Return value from GenerateOrders
    Dim dNow As Double                  ' Current time in exchange time
    Dim dNextSanityCheck As Double      ' Time of the next sanity check
    
    If m.lParentID = -1& Then
        Active = True
        RefreshUI
    ElseIf HasLevel(eTN5_Professional) = False Then
        DumpDebug "Couldn't enable because user doesn't have at least Professional version"
        Active = False
    ElseIf m.Bars.Size = 0 Then
        DumpDebug "Couldn't enable because there is no data"
        InfBox "There is no data available for " & m.strSymbol, "!", , "Automated Trading Error"
        Active = False
    ElseIf m.Strategy.ForceLimitThrough = False Then
        DumpDebug "Couldn't enable because ForceLimitThrough = False"
        InfBox "You cannot auto-trade a strategy that has the 'Limit Order Fills' flag turned off.||Auto trade results can differ significantly from back testing when that is turned off.||You will find the flag on the Settings tab while editing your strategy.", "!", , "Error"
        Active = False
    ElseIf AllRtDataAvailable = False Then
        m.bBelievePosition = bBelieveCurrentPosition
        Set m.AcctPos = AcctPos
        WaitingForData = True
        DumpDebug "Enable: Waiting for all data"
        Active = True
        
        RefreshUI
    Else
        WaitingForData = False
        
        DumpDebug "Begin Enable(" & Str(bBelieveCurrentPosition) & ")"
        DumpDebug vbTab & "Quantity Next Entry = " & Str(m.lQtyNextEntry) & "; Trade Symbol = '" & mDmDll.GetSymbol(TradeSymbolOrID) & "'"
        
        ' If the selected strategy does not have pyramiding turned on, warn the user that
        ' the pyramiding will not happen during automated trading...
        ' DAJ 03/19/2014: Allow this with a flag file now so that we can do some internal
        ' testing ( Joe Duffy is asking for it )...
        ' DAJ 08/25/2014: Changing from a flag file to an IB3X enablement since we have a few
        ' people with the flag file now, and don't allow at all for a live account...
        ' DAJ 11/21/2014: Changing this to allow for simulated accounts or broker demo accounts...
        ' DAJ 10/05/2015: Added a function and added live trading through a new enablement...
        'If (m.Strategy.Pyramid = True) And (FileExist(AddSlash(App.Path) & "AutoPyramid.FLG") = False) Then
        'If (m.Strategy.Pyramid = True) And ((HasModule("AUTOPRMD") = False) Or (g.Broker.IsLiveAccount(m.nBroker) = True)) Then
        If (m.Strategy.Pyramid = True) Then 'And ((HasModule("AUTOPRMD") = False) Or (mTradeTracker.TypeOfAccount(m.lAccountID) = eGDTypeOfAccount_BrokerLive)) Then
            If AllowPyramiding = False Then
                InfBox "Strategy '" & m.Strategy.SystemName & "' has pyramiding turned on, but pyramiding will not occur during automated trading", "i", , "Automated Trading Warning"
                m.Strategy.Pyramid = False
            End If
        End If
        
        Active = True
        m.iRolling = 0
        
        m.SubmittedOrders.Clear
        
        If Not AcctPos Is Nothing Then
            m.lCurrentPosition = AcctPos.CurrentPositionSnapshot
            PositionSymbolOrID = AcctPos.SymbolOrSymbolID
        End If
        
        Set Orders = g.Broker.WorkingOrders(m.lAccountID, PositionSymbolOrID, m.lAutoTradeItemID)
        If Not Orders Is Nothing Then
            For lIndex = 1 To Orders.Count
                DumpDebug vbTab & "Left Over Order: '" & Orders(lIndex).OrderText(True, True, True) & "' (Status = " & OrderStatus(Orders(lIndex).Status) & ")"
                m.SubmittedOrders.Add Orders(lIndex), Str(Orders(lIndex).OrderID)
            Next lIndex
        End If
        If m.SubmittedOrders.Count > 0 Then CancelSubmittedOrders False
                
        dNow = CurrentTime(m.Bars.Prop(eBARS_ExchangeTimeZoneInf), m.strSymbol, True)
        SessionDate = m.Bars.LastOpenSessionForTradeTime(dNow)
        
        ' 03/01/2016 DAJ: We want to start doing sanity checks during the day for automated trading items that are running on
        ' non-intraday bar periods.  We will do them once an hour at a 5 minute interval starting at the 3.5 minute of the
        ' hour ( so that they won't collide with new bar checks being done on 5-minute bar intervals ).  We also want to spread
        ' them out as best as possible, so we will use the automated trading item mod 12 as a way to put them in one of the
        ' five minute "buckets"...
        If m.Bars.Prop(eBARS_Periodicity) > ePRD_Days Then
            dNextSanityCheck = CDbl(Int(dNow)) + (Hour(dNow) / 24#) + ((3.5 + (5# * (m.lAutoTradeItemID Mod 12))) / 1440#)
            If dNextSanityCheck < dNow Then
                dNextSanityCheck = dNextSanityCheck + (1# / 24#)
            End If
            NextSanityCheck = dNextSanityCheck
        End If
        
        ' DAJ 11/20/2012: If we are between the On-Close time and the Session Start time, we need to
        ' generate the orders with the "OnClose" mode so that the engine considers the last bar to
        ' be complete...
        ' 1) Generate Current Orders and Current Position from Strategy...
        If (m.bBetweenOpen = True) And (DoOnCloseCheck = True) Then
            iGoodOrders = GenerateOrders(eGDAutoTradeOrdersMode_OnClose)
        Else
            iGoodOrders = GenerateOrders(eGDAutoTradeOrdersMode_Normal)
        End If
        DumpOrders "Enable"
        
        If iGoodOrders = 1 Then
            ' 2) Ask the user to verify their current position...
            If bBelieveCurrentPosition = False Then
                m.lCurrentPosition = frmCurrentPosition.ShowMe(m.lCurrentPosition)
                DumpDebug vbTab & "Current Position = " & Str(m.lCurrentPosition) & " (Enable)"
                Save
                If m.lCurrentPosition <> 0 Then
                    strSymbol = GetSymbol(TradeSymbolOrID)
                    
                    lCurPos = GetCurrentPositionFromBrokerInfo
                    DumpDebug vbTab & "Current Position: " & Str(m.lCurrentPosition) & ",CurPos  :" & Str(lCurPos)
                    If lCurPos <> m.lCurrentPosition Then
                        With Fill
                            .AccountID = m.lAccountID
                            .AutoTradingItemID = m.lAutoTradeItemID
                            If m.lCurrentPosition > 0 Then .Buy = True Else .Buy = False
                            .FillDate = m.Bars(eBARS_DateTime, m.Bars.Size - 1)
                            .IsManual = True
                            .IsSnapshot = True
                            .Price = m.Bars(eBARS_Close, m.Bars.Size - 1)
                            .Quantity = Abs(m.lCurrentPosition - lCurPos)
                            .SessionDate = m.Bars.SessionDateForTradeTime(.FillDate)
                            .SymbolOrSymbolID = PositionSymbolOrID
                            .Save m.lAccountID
                        End With
                        
                        ' TODO DAJ: Do something with the fake fill
                    End If
                End If
            End If
            
            PreviousPosition = CurrentPosition
            
            ' If they are flat, but the symbols don't match, then change the symbol...
            If (m.lCurrentPosition = 0) And (PositionSymbolOrID <> TradeSymbolOrID) Then
                ChangePositionSymbol TradeSymbolOrID, False
            End If
        
            ' 3) If the market is open...
            If MarketIsOpen Then
                ' If they are in a position, but the symbols don't match...
                If (m.lCurrentPosition <> 0) And (PositionSymbolOrID <> TradeSymbolOrID) Then
                    If m.SubmittedOrders.Count > 0 Then
                        m.iRolling = 1
                        DumpDebug vbTab & "Cancelling submitted orders because position is for a different symbol"
                        CancelSubmittedOrders False
                    Else
                        m.iRolling = 2
                        DumpDebug vbTab & "Closing position because position is for a different symbol"
                        ClosePosition False
                    End If
                
                ' If they're in a position but shouldn't be, exit at market
                ElseIf (PositionsMatch = False) And (m.lCurrentPosition <> 0) Then
                    DumpDebug vbTab & "Closing Position because positions do not match"
                    ClosePosition False, m.bBetweenOpen
                
                ' Set Quantity for orders and "place" them
                ElseIf m.bBetweenOpen = False Then
                    SetOrderQuantities
                    For lIndex = 0 To m.astrOrders.Size - 1
                        Set Order = EngineStringToOrder(SetQuantityOnEngineString(m.astrOrders(lIndex)))
                        If Order.Quantity > 0 Then
                            Order.Status = eTT_OrderStatus_Parked
                            Order.Save
                            
                            PlaceOrder Order, False, "enabling the automated trading strategy"
                        
                            If g.Broker.AutoJournalAutomated Then
                                g.JournalBridge.CreateOrderJournal Order.OrderID, CurrentTime, "Automated Trading Strategy (" & m.strName & ") placing this order because it is being enabled"
                            End If
                        Else
                            DumpDebug vbTab & "Didn't place order: '" & Order.OrderText(True, True, True) & "' because we are enabling the automated trading strategy but the order quantity is zero"
                        End If
                    Next lIndex
                End If
            End If
        
        ElseIf iGoodOrders = 0 Then
            DumpDebug vbTab & "GenerateOrders failed in Enable"
            ShowAlertPopup "'" & m.strName & "' could not be enabled because:||Orders could not be generated", "Automated Trading", vbCenter
            Active = False
        End If
        
        ' 4) Refresh the trade items grid...
        RefreshUI
        
        DumpDebug "End Enable"
    End If
    
    Enable = Active

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.Enable"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disable
'' Description: Disable the trading item
'' Inputs:      Confirm Position Close?, Close Position?, Reason,
''              Disabled from code?, Time out, Show Reason?, Check Parent?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Disable(Optional ByVal bConfirmPositionClose As Boolean = True, Optional ByVal bClosePosition As Boolean = True, Optional ByVal strReason As String = "", Optional ByVal bDisabledFromCode As Boolean = False, Optional ByVal lDialogTimeout As Long = 0&, Optional ByVal bShowReason As Boolean = False, Optional ByVal bCheckParent As Boolean = False)
On Error GoTo ErrSection:

    Dim dNyTime As Double               ' Current time in New York
    Dim bPositionClosed As Boolean      ' Is the position being closed?

    If m.lParentID = -1& Then
        Active = False
        RefreshUI
    Else
        DumpDebug "Disable(" & Str(bConfirmPositionClose) & ", " & Str(bClosePosition) & ", '" & strReason & "', " & Str(lDialogTimeout) & ")"
    
        If bShowReason Then
            ShowAlertPopup "'" & m.strName & "' being disabled because:||" & strReason, "Automated Trading", vbCenter
        End If
    
        Active = False
        CancelSubmittedOrders False
        If bClosePosition Then
            bPositionClosed = ClosePosition(bConfirmPositionClose, , , lDialogTimeout)
        End If
        
        RefreshUI
        
        If bDisabledFromCode Then
            ' ignore this alert if the disable is because streaming was disabled on Sun mornings (NY time)
            dNyTime = ConvertTimeZone(Now, "", "NY")
            If (mGenesis.Weekday(dNyTime) = vbSunday And Hour(dNyTime) < 12) = False Or strReason <> "Realtime streaming going inactive" Then
                g.Alerts.AutoTradeItemDisabled m.strSymbol, g.Broker.GetAccountNumber(m.lAccountID), m.strName, strReason
            End If
        End If
        
        If (bCheckParent = True) And (m.lParentID > 0&) Then
            g.TradingItems.CheckParent m.lParentID
        End If
        
        DumpDebug "End Disable (Active = " & Str(Active) & ")"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.Disable"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewBarCheck
'' Description: When a new bar starts, we need to recheck our orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub NewBarCheck()
On Error GoTo ErrSection:

    Dim bJustChangedThisBar As Boolean  ' Strategy Position changed with this bar
    Dim strMarketOrder As String        ' Market order to submit
    Dim Order As New cPtOrder           ' Temporary order object
    Dim lCurrentPos As Long             ' Current position coming into the function
    Dim vTradeSymbolOrID As Variant     ' Current back testing symbol or ID
    Dim vPositionSymbolOrID As Variant  ' Current position symbol or Id
    Dim lOldQuantity As Long            ' Old strategy quantity
    Dim lNewQuantity As Long            ' New strategy quantity
    Dim iGoodOrders As Integer          ' Return value from GenerateOrders
    Dim bRunOrders As Boolean           ' Re-run generate orders?
    Dim lCurrentBar As Long             ' Current bar
    Dim lPreviousBar As Long            ' Previous bar

    DumpDebug "Begin NewBarCheck"
    DumpDebug vbTab & "Quantity Next Entry = " & Str(m.lQtyNextEntry) & "; Symbol = " & m.Bars.Prop(eBARS_Symbol) & "; Actual Bar Period = '" & m.Bars.Prop(eBARS_PeriodicityStr) & "'; Trade Symbol = '" & mDmDll.GetSymbol(TradeSymbolOrID) & "'"
    
    vTradeSymbolOrID = TradeSymbolOrID
    vPositionSymbolOrID = PositionSymbolOrID
    
    If (MarketIsOpen = True) And (m.lCurrentPosition <> 0) And (vPositionSymbolOrID <> vTradeSymbolOrID) Then
        If m.SubmittedOrders.Count > 0 Then
            m.iRolling = 1
            DumpDebug vbTab & "Cancelling submitted orders because position is for a different symbol"
            CancelSubmittedOrders False
        Else
            m.iRolling = 2
            DumpDebug vbTab & "Closing position because position is for a different symbol"
            ClosePosition False
        End If
    
    ElseIf ClosingPosition = True Then
        DumpDebug vbTab & "NewBarCheck not done because we are in the process of closing the position"
        
    ElseIf ClosestOrderIsMarketOrder = True Then
        ' DAJ 09/08/2014: Tim ran into a situation back on 08/01/14 where a New-Bar check caused a market
        ' order to be submitted, but while it was working another New-Bar check happened right behind
        ' it causing us to try to cancel the market order ( it was simulated ).  If the closest order
        ' is a market order, we should probably just exit out here and let the Fill check handle things
        ' when it occurs...
        DumpDebug vbTab & "NewBarCheck not done because the closest order is a market order"
        
    ' DAJ 11/20/2012: If the position symbol doesn't match the trade symbol, but we are flat and not in
    ' the process of rolling a position, go ahead and change the position symbol here and allow orders to
    ' be generated...
    'ElseIf (m.bBetweenOpen = False) And (vPositionSymbolOrId = vTradeSymbolOrID) Then
    ElseIf (m.bBetweenOpen = False) And ((vPositionSymbolOrID = vTradeSymbolOrID) Or ((m.lCurrentPosition = 0) And (m.iRolling = 0))) Then
        If vPositionSymbolOrID <> vTradeSymbolOrID Then
            ChangePositionSymbol vTradeSymbolOrID, False
        End If
        
        lCurrentPos = m.lCurrentPosition
        
        ' 1) Generate Current Orders and Current Position from Strategy...
        iGoodOrders = GenerateOrders(eGDAutoTradeOrdersMode_NewBar, bJustChangedThisBar)
        
        If iGoodOrders = 1 Then
            DumpOrders "NewBarCheck --> GenerateOrders(" & Str(bJustChangedThisBar) & ")"
            
            ' 2) Check the current position against the strategy's current position...
            DumpDebug vbTab & "Current Position: " & Str(m.lCurrentPosition) & ", Strategy Position: " & Str(m.lStrategyPos)
            
            ' 03/01/2016 DAJ: If the positions don't match and a limit order got filled at either the high or the low of the bar before, but the
            ' current bar has taken out that high or low, re-run the engine in current bar mode to try to pick up the position...
            If PositionsMatch = False Then
                bRunOrders = False
                
                lCurrentBar = m.Bars.Size - 1
                lPreviousBar = m.Bars.Size - 2
                
                If (m.anLimitBuyAtLow(lPreviousBar) <> 0) And (m.Bars(eBARS_Low, lCurrentBar) < m.Bars(eBARS_Low, lPreviousBar)) Then
                    bRunOrders = True
                ElseIf (m.anLimitSellAtHigh(lPreviousBar) <> 0) And (m.Bars(eBARS_High, lCurrentBar) > m.Bars(eBARS_High, lPreviousBar)) Then
                    bRunOrders = True
                End If
                
                If bRunOrders = True Then
                    DumpDebug vbTab & "Re-running generate orders in current bar mode because positions don't match and this is not an inside bar"
                    iGoodOrders = GenerateOrders(eGDAutoTradeOrdersMode_Normal, bJustChangedThisBar)
                    If iGoodOrders = 1 Then
                        DumpOrders "NewBarCheck --> GenerateOrders(" & Str(bJustChangedThisBar) & ")"
                        DumpDebug vbTab & "Current Position: " & Str(m.lCurrentPosition) & ", Strategy Position: " & Str(m.lStrategyPos)
                    End If
                End If
            End If
            
            If iGoodOrders = 1 Then
                If m.lStrategyPos = 0& Then
                    If m.lCurrentPosition > 0 Then
                        strMarketOrder = "  1" & vbTab & "XL" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(Abs(m.lCurrentPosition))
                    ElseIf m.lCurrentPosition < 0 Then
                        strMarketOrder = "  1" & vbTab & "XS" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(Abs(m.lCurrentPosition))
                    End If
                ElseIf PositionsMatch Then
                    If m.Strategy.Pyramid = False Then
                        DumpDebug vbTab & "Cancel Submitted Entry Orders because positions match"
                        If CancelSubmittedOrders(True) > 0 Then
                            DoEvents
                        End If
                    End If
                ElseIf (bJustChangedThisBar = True) Then 'Or (IsFractZen = True) Then
                    lOldQuantity = m.lPrevStrategyPos * m.lQtyNextEntry
                    lNewQuantity = m.lStrategyPos * m.lQtyNextEntry
                    
                    DumpDebug vbTab & "Old Quantity = " & Str(lOldQuantity) & "; New Quantity = " & Str(lNewQuantity)
                    If lNewQuantity > 0 Then
                        If m.lCurrentPosition <= 0 Then
                            'strMarketOrder = "  1" & vbTab & "EL" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(Abs(m.lCurrentPosition) + lNewQuantity)
                            strMarketOrder = "  1" & vbTab & "EL" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(Abs(lNewQuantity))
                        ElseIf m.lCurrentPosition = lOldQuantity Then
                            If lNewQuantity >= lOldQuantity Then
                                strMarketOrder = "  1" & vbTab & "EL" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(lNewQuantity - lOldQuantity)
                            Else
                                strMarketOrder = "  1" & vbTab & "XL" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(lOldQuantity - lNewQuantity)
                            End If
                        End If
                    Else
                        If m.lCurrentPosition >= 0 Then
                            'strMarketOrder = "  1" & vbTab & "ES" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(Abs(m.lCurrentPosition) + lNewQuantity)
                            strMarketOrder = "  1" & vbTab & "ES" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(Abs(lNewQuantity))
                        ElseIf m.lCurrentPosition = lOldQuantity Then
                            If lNewQuantity <= lOldQuantity Then
                                strMarketOrder = "  1" & vbTab & "ES" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(lOldQuantity - lNewQuantity)
                            Else
                                strMarketOrder = "  1" & vbTab & "XS" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(lNewQuantity - lOldQuantity)
                            End If
                        End If
                    End If
                ElseIf m.lCurrentPosition <> 0 Then
                    If m.lCurrentPosition > 0 Then
                        strMarketOrder = "  1" & vbTab & "XL" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(Abs(m.lCurrentPosition))
                    Else
                        strMarketOrder = "  1" & vbTab & "XS" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(Abs(m.lCurrentPosition))
                    End If
                End If
                
                If lCurrentPos = m.lCurrentPosition Then
                    ' 3) Set order quantities for strategy orders...
                    SetOrderQuantities
                    DumpOrders "NewBarCheck --> SetOrderQuantities"
                
                    If Len(strMarketOrder) > 0 Then
                        strMarketOrder = FixEngineOrder(strMarketOrder)
                        
                        ' 4) If a market order, cancel all submitted and submit the market order...
                        DumpDebug vbTab & "Cancelling submitted orders because we are submitting a market order"
                        DumpDebug vbTab & vbTab & strMarketOrder
                        
                        If CancelSubmittedOrders(False) > 0 Then
                            DoEvents
                        End If
                        
                        If lCurrentPos = m.lCurrentPosition Then
                            Set Order = EngineStringToOrder(strMarketOrder)
                            If Order.Quantity > 0 Then
                                Order.Status = eTT_OrderStatus_Parked
                                Order.Save
                                
                                PlaceOrder Order, False, "the strategy is now in a position"
                            
                                If g.Broker.AutoJournalAutomated Then
                                    g.JournalBridge.CreateOrderJournal Order.OrderID, CurrentTime, "Automated Trading Strategy (" & m.strName & ") placing this order because the strategy is now in a position"
                                End If
                            Else
                                DumpDebug vbTab & "Didn't place order: '" & Order.OrderText(True, True, True) & "' because the strategy is now in a position but the order quantity is zero"
                            End If
                        End If
                    Else
                        ' 4) else, synchronize orders
                        SyncOrders "NewBar"
                    End If
                    
                    ' DAJ 01/30/2015: If a market order is present ( and therefore it is the closest order
                    ' to the market ) or there is only one order to be submitted, call SubmitOneOrder here
                    ' instead of waiting for a tick to happen.  In the case of a very slow moving market,
                    ' waiting for a tick to happen could cause large slippage...
                    If (ClosestOrderIsMarketOrder = True) Or (m.SubmittedOrders.Count = 1) Then
                        'SubmitOneOrder
                        DoSubmitOneOrder = True
                    End If
                End If
            End If
        ElseIf iGoodOrders = 0 Then
            DumpDebug vbTab & "GenerateOrders failed in NewBarCheck"
            Disable True, True, "Orders could not be generated during a new bar check", True, 30, True, True
        End If
    
        ' 5) Refresh the trade items grid...
        RefreshUI
    Else
        DumpDebug vbTab & "Between = " & Str(m.bBetweenOpen) & ", PositionSym = " & Str(PositionSymbolOrID) & ", TradeSym = " & Str(TradeSymbolOrID) & ", Rolling = " & Str(m.iRolling)
    End If
    
    DumpDebug "End NewBarCheck"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.NewBarCheck"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentBarCheck
'' Description: Run the engine on the current bar where we "believe" the data so far
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CurrentBarCheck()
On Error GoTo ErrSection:

    Dim iGoodOrders As Integer          ' Return value from GenerateOrders
    Dim lPreviousStrategyPos As Long    ' Strategy position before generating orders
    Dim strReason As String             ' Reason for doing the current bar check
            
    If Len(CurrentBarCheckReason) > 0 Then
        strReason = "CurrentBarCheck ( " & CurrentBarCheckReason & " )"
    Else
        strReason = "CurrentBarCheck"
    End If
            
    DumpDebug "Begin " & strReason
    DumpDebug vbTab & "Quantity Next Entry = " & Str(m.lQtyNextEntry) & "; Symbol = " & m.Bars.Prop(eBARS_Symbol) & "; Actual Bar Period = '" & m.Bars.Prop(eBARS_PeriodicityStr) & "'; Trade Symbol = '" & mDmDll.GetSymbol(TradeSymbolOrID) & "'"
    
    lPreviousStrategyPos = m.lStrategyPos
    iGoodOrders = GenerateOrders(eGDAutoTradeOrdersMode_Normal)
    If iGoodOrders = 1 Then
        DumpOrders strReason
        
        ' If the market is open...
        If MarketIsOpen Then
            ' If they are in a position, but the symbols don't match...
            If (m.lCurrentPosition <> 0) And (PositionSymbolOrID <> TradeSymbolOrID) Then
                If m.SubmittedOrders.Count > 0 Then
                    m.iRolling = 1
                    DumpDebug vbTab & "Cancelling submitted orders because position is for a different symbol"
                    CancelSubmittedOrders False
                Else
                    m.iRolling = 2
                    DumpDebug vbTab & "Closing position because position is for a different symbol"
                    ClosePosition False
                End If
                
            Else
                ' 03/11/2016 DAJ: If we get here and the position symbol doesn't match the trade symbol, then they must be flat
                ' so just change the symbol...
                If (UCase(CurrentBarCheckReason) = "NEWDAILYSESSION") And (PositionSymbolOrID <> TradeSymbolOrID) Then
                    ChangePositionSymbol TradeSymbolOrID, False
                End If
            
                ' If they're in a position but shouldn't be, exit at market
                If (m.lStrategyPos = 0) And (m.lCurrentPosition <> 0) Then
                    DumpDebug vbTab & "Closing Position because strategy no longer in a position"
                    ClosePosition False, m.bBetweenOpen
                
                ElseIf False Then ' (PositionsMatch = False) And (IsFractZen = True) Then
                    DumpDebug vbTab & "Entering Position because positions don't match but using FractZen bars"
                    EnterPosition False, True
                
                ' Otherwise if the strategey has just entered a position this bar and we are doing this check because we are in a new
                ' session of the symbol, but trading on a bar period greater than daily then enter the position...
                ElseIf (PositionsMatch = False) And (UCase(CurrentBarCheckReason) = "NEWDAILYSESSION") And (lPreviousStrategyPos <> m.lStrategyPos) And (m.Bars.Prop(eBARS_Periodicity) > ePRD_Days + 1) Then
                    DumpDebug vbTab & "Entering position because positions do not match"
                    EnterPosition False, True
                
                ' Otherwise if the strategy has just entered a position this bar and we are doing this check because a stop or stop-with-limit
                ' entry order has been rejected, then try to enter the position...
                ElseIf (PositionsMatch = False) And (UCase(CurrentBarCheckReason) = "REJECTEDENTRYORDER") And (lPreviousStrategyPos <> m.lStrategyPos) Then
                    DumpDebug vbTab & "Entering position because positions do not match"
                    EnterPosition False, True
                
                ' If they're in a position but shouldn't be, exit at market
                ' DAJ 04/07/2016: Moved this check down below the last two checks because we need them to hit if they need to first...
                ElseIf (PositionsMatch = False) And (m.lCurrentPosition <> 0) Then
                    DumpDebug vbTab & "Closing Position because positions do not match"
                    ClosePosition False, m.bBetweenOpen
                
                ' Set Quantity for orders and "place" them ( 02/03/2016 DAJ: Don't mess with orders if we are doing this because of a rejected order )...
                ElseIf (m.bBetweenOpen = False) And (UCase(CurrentBarCheckReason) <> "REJECTEDEXITORDER") And (UCase(CurrentBarCheckReason) <> "REJECTEDENTRYORDER") Then
                    SetOrderQuantities
                    SyncOrders strReason
                End If
            End If
        End If
    
    ElseIf iGoodOrders = 0 Then
        DumpDebug vbTab & "GenerateOrders failed in " & strReason
        Disable True, True, "Orders could not be generated after " & strReason, True, 30, True, True
    End If

    CurrentBarCheckReason = ""
    
    DumpDebug "End CurrentBarCheck"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.CurrentBarCheck"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillCheck
'' Description: When an order gets filled, recheck our other orders
'' Inputs:      Fill, Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FillCheck(ByVal Fill As cPtFill, ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lPrevPosition As Long           ' Previous position
    Dim dStartTime As Double            ' Starting time
    Dim bRolled As Boolean              ' Have we just rolled the symbol?
    Dim iGoodOrders As Integer          ' Return value from GenerateOrders
    Dim lNumTries As Long               ' Number of tries
    Dim dTime As Double                 ' Elapsed time
    Dim dFeedTime As Double             ' Feed time
    Dim lCurrentPosition As Long        ' Current position for the position symbol
    Dim lFillAmount As Long             ' Fill amount
    Dim dFillTimeNY As Double           ' Time of the fill in eastern time
    Dim lBarNumber As Long              ' Bar number for the fill
    Dim bClosedPosition As Boolean      ' Did this close a position because we were closing a position?
    Static lPrevOrderID As Long         ' Order ID of the previous fill
    
    If Order Is Nothing Then
        DumpDebug "Begin FillCheck ( '" & mTradeTracker.FillDisplay(Fill, True, False, True, False, True, False, False) & "', Nothing )"
        Set Order = New cPtOrder
        Order.Load Fill.OrderID
        DumpDebug vbTab & "Order '" & Order.OrderText(True, False, True) & "' loaded from database"
    Else
        DumpDebug "Begin FillCheck ( '" & mTradeTracker.FillDisplay(Fill, True, False, True, False, True, False, False) & "', '" & Order.OrderText(True, False, True) & "' )"
    End If
        
    If Fill.Buy Then
        lFillAmount = Fill.Quantity
    Else
        lFillAmount = Fill.Quantity * -1&
    End If
    
    ' Get the feedtime plus the seconds behind plus a couple of seconds.  Then,
    ' if that is less than the fill time, use it in case the fill time is wrong
    ' into the future for some reason...
    dFillTimeNY = Fill.FillDateConverted("NY")
    dFeedTime = g.RealTime.FeedTime + ((g.RealTime.GetSecondsBehind + 2#) / 86400#)
    If dFeedTime < dFillTimeNY Then
        dFillTimeNY = dFeedTime
    End If
    
    ' DAJ 09/12/2014: Moved the closing position up here because of an issue that Brady Preston
    ' ran into with partial fills.  We need to wait until we get all of the partial fills ( and
    ' subsequently the closing position goes to zero ) before doing anything else...
    lCurrentPosition = GetCurrentPositionFromBrokerInfo
    bClosedPosition = False
    If ClosingPosition = True Then
        If Sgn(lFillAmount) = Sgn(ClosingQuantity) Then
            ClosingQuantity = ClosingQuantity - lFillAmount
        End If
        
        If (ClosingQuantity = 0&) Then
            If (lCurrentPosition = 0&) Then
                ClosingPosition = False
                bClosedPosition = True
            Else
                ClosePosition False
            End If
        End If
    End If
    
    ' DAJ 02/10/2016: Moved this block of code up here from the else block below.  I need to keep it
    ' below the above block for the reason given in the comment above.  However, Carsten ran into a
    ' case where we perpetually kept submitting orders to close out the current position, but the current
    ' position was never being updated with the subsequent partial fills because we were always trying
    ' to close the position ( which also happened because the broker info current position "hopped" over
    ' being flat as well )...
    ' DAJ 02/12/2016: Utilize the class level previous position here now...
    ' Adjust current position...
    lPrevPosition = PreviousPosition ' m.lCurrentPosition
    m.lCurrentPosition = lCurrentPosition
    DumpDebug vbTab & "Quantity Next Entry = " & Str(m.lQtyNextEntry) & "; Previous Position = " & Str(lPrevPosition) & "; Current Position for '" & GetSymbol(PositionSymbolOrID) & "'=" & Str(m.lCurrentPosition) & "; Trade Symbol = '" & mDmDll.GetSymbol(TradeSymbolOrID) & "'"
    
    If ClosingPosition = True Then
        DumpDebug vbTab & "FillCheck not done because we are in the process of closing the position"
    Else
        If m.lCurrentPosition <> lPrevPosition Then
            DumpDebug vbTab & "Handling position change from " & Str(lPrevPosition) & " to " & Str(m.lCurrentPosition)
            Save
            
            If (m.iRolling = 2) And (m.lCurrentPosition = 0) Then
                m.iRolling = 0
                bRolled = True
                
                ChangePositionSymbol TradeSymbolOrID, False
            Else
                bRolled = False
            End If
            
            If (Active = True) And (m.bBetweenOpen = False) And (PositionSymbolOrID = TradeSymbolOrID) Then
                ' If real position has just changed to long or short...
                ' DAJ 06/28/2012: Also do this if we just went flat so that the strategy position is
                ' recalculated...
                ' DAJ 10/24/2014: Also do this if we are running a pyramiding strategy AND getting a fill for
                ' a different order than the previous fill AND we are not in the process of rolling...
                If ((m.lCurrentPosition > 0) And (lPrevPosition <= 0)) Or ((m.lCurrentPosition < 0) And (lPrevPosition >= 0)) Or ((m.lCurrentPosition = 0) And (lPrevPosition <> 0)) Or ((Pyramid = True) And (Fill.OrderID <> lPrevOrderID) And (m.iRolling = 0)) Then
                    ' Cancel all other submitted orders...
                    DumpDebug vbTab & "Cancelling Submitted Orders because the position has just changed"
                    If CancelSubmittedOrders(False, Fill.OrderID) > 0 Then
                        DoEvents
                    End If
                    
                    ' Generate Current Orders...
                    dStartTime = gdTickCount
                    ' DAJ 10/29/2014: We are passing a flag down to the engine to handle limit orders
                    ' on the last bar better, so we no longer need to do this...
                    'AdjustBarBasedOnLimitPrice Order
                    
                    ' DAJ 07/31/2012: With the fix that was put in on 6/28, we are now getting in here
                    ' if the position just went flat as well.  Since the position could be going flat
                    ' because we are rolling, we need to do that check here as well...
                    If bRolled = True Then
                        iGoodOrders = GenerateOrders(eGDAutoTradeOrdersMode_Normal)
                        If iGoodOrders = 1 Then
                            If (m.lStrategyPos <> 0&) And (m.lQtyNextEntry > 0&) Then
                                DumpDebug vbTab & "Entering a position because we are rolling the position"
                                EnterPosition
                            Else
                                SetOrderQuantities
                                SyncOrders "Fill"
                            End If
                        
                        ElseIf iGoodOrders = 0 Then
                            DumpDebug vbTab & "GenerateOrders failed in FillCheck (1)"
                            Disable True, True, "Orders could not be generated during a fill check", True, 30, True, True
                        End If
                    Else
                        ' DAJ 04/30/2014: Check the fill time against the feed time.  If the feed time
                        ' is less than the fill time ( meaning that the stream is behind ), wait until
                        ' the feed time >= the fill time OR The stream is disconnected ( which it will
                        ' eventually if the feed doesn't catch back up )...
                        DataBehind = False
                        Do
                            dFeedTime = g.RealTime.FeedTime(m.strSymbol)
                            If dFeedTime < dFillTimeNY Then
                                DataBehind = True
                                DumpDebug vbTab & "Feed Time ( " & DateFormat(dFeedTime, MM_DD_YYYY, HH_MM_SS) & " ) behind Fill Time ( " & DateFormat(dFillTimeNY, MM_DD_YYYY, HH_MM_SS) & " )"
                                Sleep 0.25
                            End If
                        Loop Until ((dFeedTime >= dFillTimeNY) Or (g.RealTime.ConnectionStatus <> eGDConnectionStatus_Connected))
                        
                        ' If the order for this fill was a Limit order, update the lowest buy or
                        ' highest sell for the correct bar if appropriate...
                        If Order.OrderType = eTT_OrderType_Limit Then
                            lBarNumber = BarNumberForFill(Fill)
                            UpdateLimitPriceArrays lBarNumber, Order.LimitPrice, Order.Buy, True
                        End If
                                                
                        If g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
                            lNumTries = 0&
                            Do
                                iGoodOrders = GenerateOrders(eGDAutoTradeOrdersMode_Normal)
                                lNumTries = lNumTries + 1&
                                
                                If iGoodOrders = 1 Then
                                    DumpOrders "FillCheck"
                                    If (PositionsMatch = False) And (bClosedPosition = False) Then
                                        dTime = gdTickCount - dStartTime
                                        
                                        DumpDebug vbTab & "Positions do not match ( Time = " & Str(dTime) & "; NumTries = " & Str(lNumTries) & " )"
                                        
                                        ' DAJ 09/26/2013: Do the sleep after the time-out check so that if the sleep
                                        ' ends up taking more than 5 seconds ( which we are seeing in some cases ),
                                        ' we will loop around and check at least one more time...
                                        ' DAJ 09/26/2013: Tim and I also decided to try at least 5 times...
                                        'Sleep 0.25
                                        If (dTime >= 5000#) And (lNumTries >= 5&) Then Exit Do
                                        Sleep 0.25
                                    End If
                                End If
                            Loop Until (PositionsMatch = True) Or (bClosedPosition = True) Or (iGoodOrders <> 1)
                            
                            If iGoodOrders = 1 Then
                                ' If still position mismatch, then get out of position...
                                If (PositionsMatch = False) And (bClosedPosition = False) Then
                                    DumpDebug vbTab & "Close position because positions do not match"
                                    ClosePosition False
                                    
                                ' Otherwise, synchronize orders...
                                Else
                                    SyncOrders "Fill"
                                End If
                            
                            ElseIf iGoodOrders = 0 Then
                                DumpDebug vbTab & "GenerateOrders failed in FillCheck (2)"
                                Disable True, True, "Orders could not be generated during a fill check", True, 30, True, True
                            End If
                        End If
                    End If
                    
                ElseIf bRolled = True Then
                    iGoodOrders = GenerateOrders(eGDAutoTradeOrdersMode_Normal)
                    If iGoodOrders = 1 Then
                        If (m.lStrategyPos <> 0&) And (m.lQtyNextEntry > 0&) Then
                            DumpDebug vbTab & "Entering a position because we are rolling the position"
                            EnterPosition
                        Else
                            SetOrderQuantities
                            SyncOrders "Fill"
                        End If
                    
                    ElseIf iGoodOrders = 0 Then
                        DumpDebug vbTab & "GenerateOrders failed in FillCheck (3)"
                        Disable True, True, "Orders could not be generated during a fill check", True, 30, True, True
                    End If
                    
                ' Otherwise adjust quantity of existing orders...
                Else
                    SetOrderQuantities
                    SyncOrders "Fill"
                End If
            End If
            
            PreviousPosition = CurrentPosition
        
            ' Refresh the trade items grid...
            RefreshUI
            m.bOrderFilled = True
        End If
    End If
    
    lPrevOrderID = Fill.OrderID
    
    DumpDebug "End FillCheck"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.FillCheck"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPosition
'' Description: Refresh the actual position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshPosition()
On Error GoTo ErrSection:

    Dim lPrevPosition As Long           ' Previous position
    Dim dStartTime As Double            ' Starting time
    Dim iGoodOrders As Integer          ' Return value from GenerateOrders

    DumpDebug "Begin RefreshPosition()"
    DumpDebug vbTab & "Quantity Next Entry = " & Str(m.lQtyNextEntry) & "; Trade Symbol = '" & mDmDll.GetSymbol(TradeSymbolOrID) & "'"
    
    ' Adjust current position...
    lPrevPosition = m.lCurrentPosition
    m.lCurrentPosition = GetCurrentPositionFromBrokerInfo
    
    If m.lCurrentPosition <> lPrevPosition Then
        DumpDebug vbTab & " Changing current position from " & Str(lPrevPosition) & " to " & Str(m.lCurrentPosition) & " because position refresh called"
        Save
        
        If (Active = True) And (m.bBetweenOpen = False) And (PositionSymbolOrID = TradeSymbolOrID) Then
            ' If real position has just changed to long or short...
            If ((m.lCurrentPosition > 0) And (lPrevPosition <= 0)) Or ((m.lCurrentPosition < 0) And (lPrevPosition >= 0)) Then
                ' Cancel all other submitted orders...
                DumpDebug vbTab & "Cancelling Submitted Orders because the position has just changed"
                If CancelSubmittedOrders(False) > 0 Then
                    DoEvents
                End If
                
                ' Generate Current Orders...
                dStartTime = gdTickCount
                Do
                    iGoodOrders = GenerateOrders(eGDAutoTradeOrdersMode_Normal)
                    If iGoodOrders = 1 Then
                        DumpOrders "RefreshPosition"
                        If Not PositionsMatch Then
                            DumpDebug vbTab & "Positions do not match"
                            Sleep 0.25
                            If (gdTickCount - dStartTime) >= 5000# Then Exit Do
                        End If
                    End If
                Loop Until (PositionsMatch = True) Or (iGoodOrders <> 1)
                
                If iGoodOrders = 1 Then
                    ' If still position mismatch, then get out of position...
                    If Not PositionsMatch Then
                        DumpDebug vbTab & "Close position because positions do not match"
                        ClosePosition False
                        
                    ' Otherwise, synchronize orders...
                    Else
                        SyncOrders "RefreshPosition"
                    End If
                
                ElseIf iGoodOrders = 0 Then
                    DumpDebug vbTab & "GenerateOrders failed in RefreshPosition"
                    Disable True, True, "Orders could not be generated during a position check", True, 30, True, True
                End If
                
            ' Otherwise adjust quantity of existing orders...
            Else
                SetOrderQuantities
                SyncOrders "RefreshPosition"
            End If
        End If
            
        PreviousPosition = CurrentPosition
    
        ' Refresh the trade items grid...
        RefreshUI
    End If
    
    DumpDebug "End RefreshPosition"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.RefreshPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OnCloseCheck
'' Description: When we hit the "On-Close Check Time", we need to recheck orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OnCloseCheck()
On Error GoTo ErrSection:

    Dim lPrevStrategyPos As Long        ' Strategy position before GenerateOrders call
    Dim lStrategyPosition As Long       ' Strategy position before GenerateOrders call
    Dim strMarketOrder As String        ' Market order to submit
    Dim Order As New cPtOrder           ' Temporary order object
    Dim lNewQuantity As Long            ' New strategy quantity
    Dim iGoodOrders As Integer          ' Return value from GenerateOrders

    DumpDebug "Begin OnCloseCheck"
    DumpDebug vbTab & "Quantity Next Entry = " & Str(m.lQtyNextEntry) & "; Strategy Position = " & Str(m.lStrategyPos) & "; Trade Symbol = '" & mDmDll.GetSymbol(TradeSymbolOrID) & "'"
    
    lPrevStrategyPos = m.lStrategyPos
    
    If ClosingPosition = True Then
        DumpDebug vbTab & "OnCloseCheck not done because we are in the process of closing the position"
        
    ElseIf DoOnCloseCheck = False Then
        DumpDebug vbTab & "OnCloseCheck not done because not closing session of bar"
        
    ElseIf (ExitAtEndOfDay = True) And (CurrentPosition <> 0) Then
        DumpDebug vbTab & "Cancelling submitted orders because ExitAtEndOfDay = " & Str(ExitAtEndOfDay) & " and CurrentPosition = " & Str(CurrentPosition)
        CancelSubmittedOrders False
        
        DumpDebug vbTab & "Closing position because ExitAtEndOfDay = " & Str(ExitAtEndOfDay) & " and CurrentPosition = " & Str(CurrentPosition)
        ClosePosition False
        
        DumpDebug vbTab & vbTab & "Changing strategy position from " & Str(m.lStrategyPos) & " to 0"
        m.lStrategyPos = 0&
    
    ElseIf m.Strategy.HasOnCloseRules = False Then
        ' DAJ 06/26/2014: As per Tim, if the strategy doesn't have any on-close rules active,
        ' all we need to do is cancel the submitted orders after the close...
        ' DAJ 09/08/2014: Terry Knight ran into a situation where a New-Bar check caused a market
        ' order to go out instants before the On-Close check.  This caused his market order to be
        ' cancelled ( because it was simulated ) even though the strategy doesn't have On-Close
        ' orders.  As per Tim, we don't want to do anything here now if there are no On-Close
        ' orders in the strategy...
        'DumpDebug vbTab & "Cancelling Submitted Orders because close is coming"
        'CancelSubmittedOrders False
        DumpDebug vbTab & "OnCloseCheck not done because there are no On-Close orders in the strategy"
    
    ' 1) Call normal backtest and act as if this bar is complete...
    Else
        iGoodOrders = GenerateOrders(eGDAutoTradeOrdersMode_OnClose)
        If iGoodOrders = 1 Then
            DumpOrders "OnCloseCheck"
            
            strMarketOrder = ""
            
            ' 2) Cancel all submitted orders...
            DumpDebug vbTab & "Cancelling Submitted Orders because close is coming"
            If CancelSubmittedOrders(False) > 0 Then
                DoEvents
            End If
            
            lStrategyPosition = m.lStrategyPos * m.lQtyNextEntry
            
            ' 3) If in a position that we shouldn't be, submit MOC order to exit...
            If (m.lCurrentPosition <> 0&) And (m.lStrategyPos = 0&) Then
                DumpDebug vbTab & "Closing Position because the positions do not match"
                ClosePosition False, True
            
            ' 4) Otherwise, if strategy position just changed...
            ElseIf (lPrevStrategyPos = 0&) And (m.lStrategyPos <> 0&) Then
                ' If the strategy just entered, get in with a MOC order...
                If m.lCurrentPosition = 0& Then
                    lNewQuantity = Abs(lStrategyPosition)
                    
                    DumpDebug vbTab & vbTab & "On-Close check just got into a position"
                    If m.lStrategyPos > 0 Then
                        strMarketOrder = "  1" & vbTab & "EL" & vbTab & "0" & vbTab & "MOC" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(lNewQuantity)
                    Else
                        strMarketOrder = "  1" & vbTab & "ES" & vbTab & "0" & vbTab & "MOC" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(lNewQuantity)
                    End If
                
                ' Otherwise, synchronize the positions...
                Else
                    If lStrategyPosition > m.lCurrentPosition Then
                        lNewQuantity = lStrategyPosition - m.lCurrentPosition
                        
                        If (m.lStrategyPos > 0) And (m.lCurrentPosition > 0) Then
                            ' Going further long...
                            DumpDebug vbTab & vbTab & "On-Close check going further long"
                            strMarketOrder = "  1" & vbTab & "EL" & vbTab & "0" & vbTab & "MOC" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(lNewQuantity)
                        ElseIf m.lStrategyPos > 0 Then
                            ' Reversing position...
                            DumpDebug vbTab & vbTab & "On-Close check reversing position"
                            strMarketOrder = "  1" & vbTab & "EL" & vbTab & "0" & vbTab & "MOC" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(lStrategyPosition)
                        Else
                            ' Exiting a portion of a short position...
                            DumpDebug vbTab & vbTab & "On-Close check exiting a portion of the position"
                            strMarketOrder = "  1" & vbTab & "XS" & vbTab & "0" & vbTab & "MOC" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(lNewQuantity)
                        End If
                    
                    ElseIf lStrategyPosition < m.lCurrentPosition Then
                        lNewQuantity = Abs(lStrategyPosition - m.lCurrentPosition)
                        
                        If m.lStrategyPos > 0 Then
                            ' Exiting a portion of a long position...
                            DumpDebug vbTab & vbTab & "On-Close check exiting a portion of the position"
                            strMarketOrder = "  1" & vbTab & "XL" & vbTab & "0" & vbTab & "MOC" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(lNewQuantity)
                        ElseIf m.lCurrentPosition < 0 Then
                            ' Going further short...
                            DumpDebug vbTab & vbTab & "On-Close check going further short"
                            strMarketOrder = "  1" & vbTab & "ES" & vbTab & "0" & vbTab & "MOC" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(lNewQuantity)
                        Else
                            ' Reversing position...
                            DumpDebug vbTab & vbTab & "On-Close check reversing position"
                            strMarketOrder = "  1" & vbTab & "ES" & vbTab & "0" & vbTab & "MOC" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(lStrategyPosition)
                        End If
                    End If
                End If
            End If
            
            If Len(strMarketOrder) > 0 Then
                DumpDebug vbTab & vbTab & strMarketOrder
                
                Set Order = EngineStringToOrder(strMarketOrder, False)
                If Order.Quantity > 0 Then
                    Order.Status = eTT_OrderStatus_Parked
                    Order.Save
                    
                    PlaceOrder Order, False, "quantity changed on an on-close signal"
                
                    If g.Broker.AutoJournalAutomated Then
                        g.JournalBridge.CreateOrderJournal Order.OrderID, CurrentTime, "Automated Trading Strategy (" & m.strName & ") placing this order because strategy got into the market on a market-on-close signal"
                    End If
                Else
                    DumpDebug vbTab & "Didn't place order: '" & Order.OrderText(True, True, True) & "' because strategy got into the market on a market-on-close signal but the order quantity is zero"
                End If
            End If
        ElseIf iGoodOrders = 0 Then
            DumpDebug vbTab & "GenerateOrders failed in OnCloseCheck"
            Disable True, True, "Orders could not be generated during a on-close check", True, 30, True, True
        End If
    End If

    ' 5) Refresh the trade items grid...
    RefreshUI
    
    DumpDebug "End OnCloseCheck"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.OnCloseCheck"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RollContracts
'' Description: Allow the user to roll contracts (if trading futures)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RollContracts()
On Error GoTo ErrSection:

    Dim strNewContract As String        ' Contract to roll into
    Dim lCurrentPosition As Long        ' Current position (before disable)
    Dim strMarketOrder As String        ' Market order to submit
    Dim Order As New cPtOrder           ' Temporary order object
    
    DumpDebug "Begin RollContracts"
    
    If frmRollContracts.ShowMe(m.strSymbol, strNewContract) = True Then
        lCurrentPosition = m.lCurrentPosition
        
        ' Cancel existing orders and close position...
        DumpDebug vbTab & "Cancelling submitted orders and closing position because we are rolling position"
        CancelSubmittedOrders False
        ClosePosition False
        
        ' Change the symbol and symbol ID, save the auto trading item, and reload the bars...
        m.strSymbol = strNewContract
        m.lSymbolID = GetSymbolID(strNewContract)
        Save
        LoadBars
        
        ' Submit market order to get into the same position with new contract...
        If lCurrentPosition <> 0 Then
            Select Case lCurrentPosition
                Case Is > 0
                    strMarketOrder = "  1" & vbTab & "EL" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(Abs(lCurrentPosition))
                Case Is < 0
                    strMarketOrder = "  1" & vbTab & "ES" & vbTab & "0" & vbTab & "M" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Str(Abs(lCurrentPosition))
            End Select
            strMarketOrder = FixEngineOrder(strMarketOrder)
            Set Order = EngineStringToOrder(strMarketOrder)
            If Order.Quantity > 0 Then
                Order.Status = eTT_OrderStatus_Parked
                Order.Save
                
                PlaceOrder Order, False, "we are rolling the position"
            
                If g.Broker.AutoJournalAutomated Then
                    g.JournalBridge.CreateOrderJournal Order.OrderID, CurrentTime, "Automated Trading Strategy (" & m.strName & ") placing this order because the position is being rolled into a new contract"
                End If
            Else
                DumpDebug vbTab & "Didn't place order: '" & Order.OrderText(True, True, True) & "' because we are rolling the position but the order quantity is zero"
            End If
        End If
    End If
    
    DumpDebug "End RollContracts"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.RollContracts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Send the given string to the test form and the debug log
'' Inputs:      String to Send
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DumpDebug(ByVal strMessage As String)
On Error GoTo ErrSection:

    GenesisLog "(" & Str(m.lAutoTradeItemID) & ", '" & m.strName & "', '" & m.strSymbol & "'): " & strMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.DumpDebug"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderStatusChange
'' Description: Handle an order status change appropriately
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderStatusChange(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    DumpDebug "Begin OrderStatusChange: '" & Order.OrderText(True, True, True) & "' New Status = '" & OrderStatus(Order.Status) & "'"

    Select Case Order.Status
        Case eTT_OrderStatus_Filled, eTT_OrderStatus_Cancelled, eTT_OrderStatus_Expired, eTT_OrderStatus_Rejected, eTT_OrderStatus_Error, eTT_OrderStatus_BalCancelled, eTT_OrderStatus_OverFilled
            DumpDebug vbTab & "Order Removed: '" & Order.OrderText(True, True, True) & "' because order status has changed to " & OrderStatus(Order.Status)
            
            m.SubmittedOrders.Remove Str(Order.OrderID)
            If InStr(m.strConfirmedSubmit, "," & Str(Order.OrderID) & ",") <> 0 Then
                m.strConfirmedSubmit = Replace(m.strConfirmedSubmit, "," & Str(Order.OrderID) & ",", ",")
            End If
            
            If ((Order.Status = eTT_OrderStatus_Cancelled) Or (Order.Status = eTT_OrderStatus_BalCancelled)) Then
                m.bOrderCancelled = True
            
            ElseIf Order.Status = eTT_OrderStatus_Rejected Then
                m.bOrderRejected = True
                
                ' DAJ 02/03/2016: If an order gets rejected, we want to do a current bar check to see if the engine now thinks they are
                ' out of the position ( e.g. wrong side of market situation )...
                ' DAJ 06/17/2016: Don't try to do anything on a rejected order if it was a market order, because if a market order is
                ' being rejected, they likely won't be able to flatten the position with another market order...
                If ((m.lCurrentPosition <> 0) And (Order.OrderType <> eTT_OrderType_Market)) Then
                    CurrentBarCheckReason = "RejectedExitOrder"
                    g.TradingItems.AddCurrentBarCheck m.lAutoTradeItemID
                
                ' DAJ 03/11/2016: If an order gets rejected AND the automated trading item is not in a position AND the order was a stop or
                ' stop-with-limit, then run a current bar check to see if the strategy got into a position ( use this to handle cases where
                ' the order got rejected for being on the wrong side of the market -- however, if they were rejected for insuffient funds, the
                ' corresponding market order to get them into the position will get rejected as well and nothing will happen due to that
                ' reject...
                ElseIf (m.lCurrentPosition = 0) And ((Order.OrderType = eTT_OrderType_Stop) Or (Order.OrderType = eTT_OrderType_StopWithLimit)) Then
                    CurrentBarCheckReason = "RejectedEntryOrder"
                    g.TradingItems.AddCurrentBarCheck m.lAutoTradeItemID
                End If
            End If
            
        Case Else
            DumpDebug vbTab & "Order Updated: '" & Order.OrderText(True, True, True) & "' because order status has changed to " & OrderStatus(Order.Status)
            DumpDebug vbTab & vbTab & Order.EngineString
            Set m.SubmittedOrders(Str(Order.OrderID)) = Order
    
    End Select
    
    If (m.iRolling = 1) And (m.SubmittedOrders.Count = 0) And (m.lCurrentPosition <> 0) Then
        m.iRolling = 2
        DumpDebug vbTab & "Closing position because position is for a different symbol"
        ClosePosition False
    End If
    
    RefreshUI

    DumpDebug "End OrderStatusChange"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.OrderStatusChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReloadData
'' Description: Reload all of the data (likely just after a daily download)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReloadData()
On Error GoTo ErrSection:

    DumpDebug "Begin ReloadData"
    LoadBars
    
    If ParentID >= 0 Then
        FixOnCloseTime
        
        If ClosingPosition = True Then
            DumpDebug vbTab & "ReloadDataCheck not done because we are in the process of closing the position"
        
        ' Generate Current Orders and Current Position from Strategy...
        Else
            ' 08/28/2015 DAJ: Moved this block up here from down under DumpOrders because we need to cancel
            ' orders on the old symbol if the position is flat whether or not we have data for the new session...
            ' If they are flat, but the symbols don't match, then change the symbol...
            If (m.lCurrentPosition = 0) And (PositionSymbolOrID <> TradeSymbolOrID) Then
                ChangePositionSymbol TradeSymbolOrID, True
            End If
            
            CurrentBarCheckReason = "ReloadData"
            CurrentBarCheck
        End If
        
        ' 4) Refresh the trade items grid...
        RefreshUI
    Else
        DumpDebug vbTab & "ReloadData not run because this is a parent"
    End If

    DumpDebug "End ReloadData"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.ReloadData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllRtDataAvailable
'' Description: Is all of the streaming data available?
'' Inputs:      None
'' Returns:     True if all data available, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllRtDataAvailable() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Static bAlreadyLogged As Boolean    ' Have we already logged information?
    Static bLastKnown As Boolean        ' Last known value
    
    If g.RealTime.RtDataAvailableForBars(m.Bars) Then
        bReturn = True
        For lIndex = 1 To m.SecondaryMarkets.Count
            If g.RealTime.RtDataAvailableForBars(m.SecondaryMarkets(lIndex)) = False Then
                bReturn = False
                Exit For
            End If
        Next lIndex
    Else
        bReturn = False
    End If
    
    If (bAlreadyLogged = False) Or (bReturn <> bLastKnown) Then
        DumpDebug "AllRtDataAvailable = " & Str(bReturn)
        bAlreadyLogged = True
        bLastKnown = bReturn
    End If
        
    AllRtDataAvailable = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.AllRtDataAvailable"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectionStatusChanged
'' Description: Notification that the connection status for the given broker
''              has just changed
'' Inputs:      Broker, Connection Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ConnectionStatusChanged(ByVal nBroker As eTT_AccountType, ByVal nNewStatus As eGDConnectionStatus)
On Error GoTo ErrSection:

    If nBroker = m.nBroker Then
        nNewStatus = g.Broker.ConnectionStatusForAccount(m.lAccountID)
        If nNewStatus <> m.nBrokerStatus Then
            DumpDebug "Connection status for '" & g.Broker.BrokerName(m.nBroker) & "' account '" & g.Broker.AccountNumberForID(m.lAccountID) & "' has changed from '" & g.BrokerEnums.ConnectionStatusToString(m.nBrokerStatus) & "' to '" & g.BrokerEnums.ConnectionStatusToString(nNewStatus) & "'"
            m.nBrokerStatus = nNewStatus
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.ConnectionStatusChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsGuru
'' Description: Does this item automate a guru strategy or basket?
'' Inputs:      None
'' Returns:     True if guru, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsGuru() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim Basket As cStrategyBasket       ' Strategy basket object

    bReturn = False
    If m.lStrategyBasketID > 0 Then
        Set Basket = New cStrategyBasket
        If Basket.LoadDb(m.lStrategyBasketID, , True) Then
            bReturn = Basket.IsGuru
        End If
    ElseIf m.lStrategyID > 0 Then
        bReturn = m.Strategy.IsGuru
    End If
    
    IsGuru = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.IsGuru"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsGuruOwner
'' Description: Is this client the owner of the guru basket or strategy?
'' Inputs:      None
'' Returns:     True if guru owner, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsGuruOwner() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = True
    If IsGuru Then
        If m.lStrategyBasketID > 0 Then
            bReturn = mSysNav.IsOwnerOfGuruBasket(m.lStrategyBasketID)
        ElseIf m.lStrategyID > 0 Then
            bReturn = mSysNav.IsOwnerOfGuruStrategy(m.lStrategyID)
        End If
    End If
    
    IsGuruOwner = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.IsGuruOwner"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CombineOverrides
'' Description: Combine the automated trading item overrides and the strategy
''              basket overrides into one override string to send to the engine
'' Inputs:      Override string
'' Returns:     Combined override string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CombineOverrides(ByVal strOverrides As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrOverrides As cGdArray       ' Overrides split into an array
    Dim astrReturn As cGdArray          ' Overrides to return
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position in the array
    
    If InStr(strOverrides, "|") = 0 Then
        strReturn = strOverrides
    Else
        Set astrReturn = New cGdArray
        astrReturn.Create eGDARRAY_Strings
        
        Set astrOverrides = New cGdArray
        astrOverrides.SplitFields Parse(strOverrides, "|", 1), ","
        For lIndex = 0 To astrOverrides.Size - 1
            If astrReturn.BinarySearch(Parse(astrOverrides(lIndex), "=", 1) & "=", lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                astrReturn(lPos) = astrOverrides(lIndex)
            Else
                astrReturn.Add astrOverrides(lIndex), lPos
            End If
        Next lIndex
        
        astrOverrides.SplitFields Parse(strOverrides, "|", 2), ","
        For lIndex = 0 To astrOverrides.Size - 1
            If astrReturn.BinarySearch(Parse(astrOverrides(lIndex), "=", 1) & "=", lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                astrReturn(lPos) = astrOverrides(lIndex)
            Else
                astrReturn.Add astrOverrides(lIndex), lPos
            End If
        Next lIndex
        
        strReturn = astrReturn.JoinFields(",")
    End If
    
    CombineOverrides = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.CombineOverrides"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetCurrentPositionFromBrokerInfo
'' Description: Get the current position for this automated trading item from
''              the appropriate broker info object
'' Inputs:      None
'' Returns:     Current Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCurrentPositionFromBrokerInfo() As Long
On Error GoTo ErrSection:

    GetCurrentPositionFromBrokerInfo = g.Broker.CurrentPosition(m.lAccountID, PositionSymbolOrID, m.lAutoTradeItemID)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.GetCurrentPositionFromBrokerInfo"
    
End Function

Private Sub LoadBars()
On Error GoTo ErrSection:

    Dim dStartTime#, nFrom&, lIndex&
    Dim Bars As cGdBars
    Dim strSymbol As String             ' Symbol to load
    Dim strPeriod As String             ' Period to load
    Dim lActualPeriodicity As Long      ' Periodicity for the bars

    dStartTime = gdTickCount
    
    If m.strSymbol <> m.Bars.Prop(eBARS_Symbol) Then
        g.RealTime.RemoveTickBuffer m.Bars
        SetBarProperties m.Bars, m.strSymbol
        m.Bars.Size = 0
    End If
    
    If IsFractZen = True Then
        g.FractZen.SetFractZen m.Bars
        strPeriod = m.Bars.Prop(eBARS_PeriodicityStr)
        If strPeriod <> m.strActualBarPeriod Then
            If Len(m.strActualBarPeriod) = 0 Then
                DumpDebug vbTab & "FractZen period set to '" & strPeriod & "'"
            Else
                DumpDebug vbTab & "FractZen period changed from '" & m.strActualBarPeriod & "' to '" & strPeriod & "'"
            End If
            m.strActualBarPeriod = strPeriod
        Else
            DumpDebug vbTab & "FractZen period stayed at '" & strPeriod & "'"
        End If
    End If
    
    ' load primary bars
    If FileExist(App.Path & "\OldAutoTrade.flg") Then
        nFrom = 1
        m.lLDDwhenLoadBars = 0
    Else
        If IsIntraday(GetPeriodicity(m.strActualBarPeriod)) And Not m.Strategy Is Nothing Then
            nFrom = -1.2 * (m.Strategy.BarsLoadedBeforeTrading + m.Strategy.BarsTradedBeforeOrders)
        Else
            nFrom = 0
        End If
        m.lLDDwhenLoadBars = LastDailyDownload
    End If
    
    ' DAJ 01/11/2016: If you just load and splice data in bar periods greater than Daily, the last bar ends up with
    ' the current session date instead of the ending session date of the bar.  If we load daily data here and then
    ' compress the data after we splice it, we seem to get the ending session of the bar instead ( which is what we
    ' want for comparisons later )...
    lActualPeriodicity = mGdDll.GetPeriodicity(m.strActualBarPeriod)
    If lActualPeriodicity > ePRD_Days + 1 Then
        DumpDebug vbTab & "Requesting Market1 data for '" & m.strSymbol & ";" & "Daily" & "' from " & Str(nFrom)
        DM_GetBars m.Bars, m.strSymbol, "Daily", nFrom, m.lLDDwhenLoadBars
    Else
        DumpDebug vbTab & "Requesting Market1 data for '" & m.strSymbol & ";" & m.strActualBarPeriod & "' from " & Str(nFrom)
        DM_GetBars m.Bars, m.strSymbol, m.strActualBarPeriod, nFrom, m.lLDDwhenLoadBars
    End If
    
    g.RealTime.AddTickBuffer m.Bars
    g.RealTime.SpliceBars m.Bars
    
    If lActualPeriodicity > ePRD_Days + 1 Then
        DumpDebug vbTab & vbTab & "Building '" & m.strActualBarPeriod & "' bars for Market1 from '" & m.strSymbol & ";Daily'"
        m.Bars.BuildBars m.strActualBarPeriod
    End If
    
    LoadLimitArrays
    
    ' load secondary markets
    m.SecondaryMarkets.Clear
    For lIndex = 0 To m.astrSecondaryMarkets.Size - 1
        Set Bars = New cGdBars
        If InStr(m.astrSecondaryMarkets(lIndex), " MinuteXX") > 0 Then
            nFrom = -2 ' just need to load a very minimal amount for the "unused" intraday data
        ElseIf IsIntraday(GetPeriodicity(Parse(m.astrSecondaryMarkets(lIndex), ";", 2))) And Not m.Strategy Is Nothing Then
            nFrom = -1.2 * (m.Strategy.BarsLoadedBeforeTrading + m.Strategy.BarsTradedBeforeOrders)
        Else
            nFrom = 0
        End If
        
        strSymbol = Parse(m.astrSecondaryMarkets(lIndex), ";", 1)
        strPeriod = Parse(m.astrSecondaryMarkets(lIndex), ";", 2)
        
        DumpDebug vbTab & "Requesting data for '" & strSymbol & ";" & strPeriod & "' from " & Str(nFrom)
        DM_GetBars Bars, strSymbol, strPeriod, nFrom
        g.RealTime.AddTickBuffer Bars
        g.RealTime.SpliceBars Bars
        m.SecondaryMarkets.Add Bars
    Next lIndex
    Set Bars = Nothing
    
    CalcBetween
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.LoadBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateOrders
'' Description: Have the strategy generate orders and current position
'' Inputs:      Mode, Just Changed This Bar?
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenerateOrders(ByVal nMode As eGDAutoTradeOrdersMode, Optional bJustChangedThisBar As Boolean) As Integer
On Error GoTo ErrSection:

    Dim iReturn As Integer              ' Return value for the function
    Dim bReturn As Boolean              ' Return value from the order generation
    Dim astrTrades As New cGdArray      ' Trade-by-trade array from the strategy
    Dim astrNextBar As New cGdArray     ' Next bar report from the strategy
    Dim astrRuleBased As New cGdArray   ' Rule based next bar report from the strategy
    Dim lIndex As Long                  ' Index into a for loop
    Dim strLastTrade As String          ' Last trade from the trades file
    Dim lLastTradeIndex As Long         ' Index of the last completed trade in the trades file
    Dim dLastDate As Double             ' Date/Time of last trade
    Dim nPeriodType As eBarsPeriodType  ' Bar period type
    Dim strDebug As String              ' Text to send to the debug log
    Dim strError As String              ' Error string back from the engine
    Dim strText As String               ' Text to send to log and display
    Dim strMode As String               ' Mode we are running in
    Dim Bars As cGdBars                 ' Bars object
    Dim astrOrder As cGdArray           ' Array of order information
    Dim bCalcJustChanged As Boolean     ' Do we need to calculate the just changed?
    Dim dLastEntry As Double            ' Time of the last entry
    Dim dLastExit As Double             ' Time of the last exit
    Dim strLimitFlag As String          ' Limit flag to send to the engine
    Dim lBarsSession As Long            ' Session date for the last known data
    Dim lCurrentSession As Long         ' Session date for the current time
    Dim dChkTime As Double
    
    Select Case nMode
        Case eGDAutoTradeOrdersMode_Normal
            strMode = "Normal"
        Case eGDAutoTradeOrdersMode_NewBar
            strMode = "NewBar"
        Case eGDAutoTradeOrdersMode_OnClose
            strMode = "OnClose"
        Case Else
            strMode = Str(nMode)
    End Select
        
    lBarsSession = m.Bars.SessionDate(m.Bars.Size - 1)
    lCurrentSession = m.Bars.SessionDateForTradeTime(CurrentTime(m.Bars.Prop(eBARS_ExchangeTimeZoneInf), m.strSymbol, True))
    
    If (nMode = eGDAutoTradeOrdersMode_Normal) And (lBarsSession < lCurrentSession) Then
        DumpDebug "Didn't Generate Orders('" & strMode & "', " & Str(bJustChangedThisBar) & ") - Previous Strategy Position: " & Str(m.lPrevStrategyPos) & ", Current Strategy Position: " & Str(m.lStrategyPos) & ", Bars Session: " & DateFormat(lBarsSession, MM_DD_YYYY) & ", Current Session: " & DateFormat(lCurrentSession, MM_DD_YYYY)
        
        ' DAJ 12/18/2014: Returning false from this function causes the automated trading item to
        ' be disabled because it means that the engine could not generate the orders.  In this case,
        ' there is not an error, we just chose not to generate the orders...
        'bReturn = True
        iReturn = 2
    Else
        DumpDebug "Generate Orders('" & strMode & "', " & Str(bJustChangedThisBar) & ") - Previous Strategy Position: " & Str(m.lPrevStrategyPos) & ", Current Strategy Position: " & Str(m.lStrategyPos) & ", Bars Session: " & DateFormat(lBarsSession, MM_DD_YYYY) & ", Current Session: " & DateFormat(lCurrentSession, MM_DD_YYYY)
        
        If nMode = eGDAutoTradeOrdersMode_NewBar Then
            strLimitFlag = ""
        Else
            strLimitFlag = LimitFlag
        End If
        
        ' Call the engine to generate the trades and orders...
If IsIDE Then
    dChkTime = gdTickCount
    ''frmTradeItems.Caption = "Generating orders ..."
End If
        bCalcJustChanged = False
        m.Strategy.Bars = m.Bars.MakeCopy
        bReturn = m.Strategy.AutoTradeOrders(nMode, astrTrades, astrNextBar, astrRuleBased, m.lAutoTradeItemID, strError, strLimitFlag)
If IsIDE Then
    dChkTime = gdTickCount - dChkTime
    ''frmTradeItems.Caption = "Generate orders = " & Str(Round(dChkTime)) & " ms"
End If
        If (bReturn = False) Then
            If Len(strError) > 0 Then
                strText = "Error Generating Orders: " & strError
            Else
                strText = "Error Generating Orders"
            End If
            
            DumpDebug vbTab & Replace(strText, "|", "")
            ShowAlertPopup strText, m.strName, vbLeftJustify
            
            iReturn = 0
        
        ElseIf nMode = eGDAutoTradeOrdersMode_OnClose Then
            astrTrades.ToFile AddSlash(App.Path) & "AutoTrade\AT_T_" & Str(m.lAutoTradeItemID) & ".TXT"
        
            ' Get the current position...
            m.lPrevStrategyPos = m.lStrategyPos
            m.lStrategyPos = CurrentStrategyPositionFromTrades(astrTrades)
            If m.lStrategyPos <> m.lPrevStrategyPos Then
                DumpDebug vbTab & "Strategy Position just changed from " & Str(m.lPrevStrategyPos) & " to " & m.lStrategyPos
            End If
            
            m.astrOrders.Size = 0&
            
            bCalcJustChanged = True
            iReturn = 1
        
        ElseIf astrNextBar(0) = "N/A" Then
            DumpDebug vbTab & "Consolidated orders could not be generated"
            astrRuleBased.ToFile AddSlash(App.Path) & "AutoTrade\AT_RB_" & Str(m.lAutoTradeItemID) & ".TXT"
            'bReturn = False
            iReturn = 0
                
        Else
            astrTrades.ToFile AddSlash(App.Path) & "AutoTrade\AT_T_" & Str(m.lAutoTradeItemID) & ".TXT"
            astrNextBar.ToFile AddSlash(App.Path) & "AutoTrade\AT_NB_" & Str(m.lAutoTradeItemID) & ".TXT"
            
            ' Get the current position...
            m.lPrevStrategyPos = m.lStrategyPos
            m.lStrategyPos = CurrentStrategyPosition(astrNextBar(2))
            DumpDebug vbTab & "Previous Strategy Position: " & Str(m.lPrevStrategyPos) & "; Strategy Position: " & Str(m.lStrategyPos) & "; Position Line: '" & astrNextBar(2) & "'"
            
            ' Save the orders...
            DumpDebug vbTab & "Orders:"
            m.astrOrders.Size = 0
            If astrNextBar.Size = 0 Then
                DumpDebug vbTab & "No Orders"
            Else
                For lIndex = 0 To astrNextBar.Size - 1
                    Set astrOrder = New cGdArray
                    astrOrder.SplitFields astrNextBar(lIndex), vbTab
                    
                    If astrOrder(0) = "  1" Then
                        If (astrOrder(OrderField(eGDOrderField_OrderType)) = "M") Or (astrOrder(OrderField(eGDOrderField_OrderType)) = "SWL") Or (astrOrder(OrderField(eGDOrderField_OrderType)) = "SWLCO") Then
                            astrNextBar(lIndex) = FixEngineOrder(astrNextBar(lIndex))
                            astrOrder.SplitFields astrNextBar(lIndex), vbTab
                        End If
                                                
                        If astrOrder(1) = "EL" Or astrOrder(1) = "ES" Then
                            If (m.Strategy.Pyramid = False) Or (m.lCurrentPosition = (m.lPrevStrategyPos * m.lQtyNextEntry)) Then
                                m.astrOrders.Add astrNextBar(lIndex)
                                DumpDebug vbTab & astrNextBar(lIndex)
                            Else
                                DumpDebug vbTab & astrNextBar(lIndex) & " ( Not added ) "
                            End If
                        
                        ElseIf (astrOrder(1) = "XL") And (m.lCurrentPosition <= 0) Then
                            DumpDebug vbTab & astrNextBar(lIndex) & " ( Not added because not in a long position ) "
                        
                        ElseIf (astrOrder(1) = "XS") And (m.lCurrentPosition >= 0) Then
                            DumpDebug vbTab & astrNextBar(lIndex) & " ( Not added because not in a short position ) "
                        
                        Else
                            m.astrOrders.Add astrNextBar(lIndex)
                            DumpDebug vbTab & astrNextBar(lIndex)
                        End If
                    End If
                Next lIndex
            End If
            
            bCalcJustChanged = True
            iReturn = 1
        End If
        
        If bCalcJustChanged = True Then
            ' Figure out if it just changed this bar...
            bJustChangedThisBar = False
            If astrTrades.Size > 1 Then
                strLastTrade = astrTrades(astrTrades.Size - 1)
                lLastTradeIndex = astrTrades.Size - 1
                dLastEntry = Val(Parse(strLastTrade, vbTab, 2))
                dLastExit = kNullData
                
                For lIndex = astrTrades.Size - 1 To 1 Step -1
                    strLastTrade = astrTrades(lIndex)
                    lLastTradeIndex = lIndex
                    
                    If Parse(strLastTrade, vbTab, 13) <> "(null)" Then
                        dLastExit = Val(Parse(strLastTrade, vbTab, 5))
                        Exit For
                    End If
                Next lIndex
                
                If dLastEntry > dLastExit Then
                    strDebug = "Entry: "
                    dLastDate = dLastEntry
                ElseIf dLastEntry < dLastExit Then
                    strDebug = "Exit: "
                    dLastDate = dLastExit
                Else
                    strDebug = "Both: "
                    dLastDate = dLastExit
                End If
                
                'strDebug = strDebug & "LastDate=" & DateFormat(dLastDate, MM_DD_YYYY, HH_MM) & "; Bars=" & DateFormat(m.Bars(eBARS_DateTime, m.Bars.Size - 2), MM_DD_YYYY, HH_MM)
                strDebug = strDebug & "LastDate=" & Str(dLastDate) & "; Bars(-2)=" & DateLogString(m.Bars(eBARS_DateTime, m.Bars.Size - 2)) & "; Bars(-1)=" & DateLogString(m.Bars(eBARS_DateTime, m.Bars.Size - 1)) & "; Last=" & m.Bars.PriceDisplay(m.Bars(eBARS_Close, m.Bars.Size - 1))
                
                ' DAJ 07/31/2015: Back on 01/06/2010, I had split out the intraday bar periods with the >= check meaning that we were going to allow any fills
                ' on the bar before the new bar to be considered in the "just changed this bar" check.  This change was made as a result of a customer who was trying
                ' to do automated trading on range bars and missing some fills.  Now, we had another customer who ran into an issue with this relating to minute
                ' bars.  Tim decided that we shouldn't include the bar before the new bar if minute bars are involved...
                nPeriodType = GetPeriodType(GetPeriodicity(m.strActualBarPeriod))
                If (nPeriodType >= ePRD_Days) Or (nPeriodType = ePRD_Minutes) Then
                    If dLastDate > Round(m.Bars(eBARS_DateTime, m.Bars.Size - 2), 6) Then
                        bJustChangedThisBar = True
                    End If
                Else
                    If dLastDate >= Round(m.Bars(eBARS_DateTime, m.Bars.Size - 2), 6) Then
                        bJustChangedThisBar = True
                    End If
                End If
                
                strDebug = strDebug & "; JustChangedThisBar=" & Str(bJustChangedThisBar)
                DumpDebug strDebug
                
                ' DAJ 07/31/2015: If the position just changed, spit out the last line from the trades file...
                If bJustChangedThisBar = True Then
                    DumpDebug vbTab & "Last Trade Line: " & strLastTrade
                    For lIndex = lLastTradeIndex + 1 To astrTrades.Size - 1
                        DumpDebug vbTab & vbTab & vbTab & astrTrades(lIndex)
                    Next lIndex
                End If
                
                For lIndex = 1 To m.SecondaryMarkets.Count
                    Set Bars = m.SecondaryMarkets(lIndex)
                    If Bars.Size > 1 Then
                        DumpDebug vbTab & Bars.Prop(eBARS_Symbol) & ";" & Bars.Prop(eBARS_PeriodicityStr) & ": Bars(-2)=" & DateLogString(Bars(eBARS_DateTime, Bars.Size - 2)) & "; Bars(-1)=" & DateLogString(Bars(eBARS_DateTime, Bars.Size - 1)) & "; Last=" & Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1))
                    ElseIf Bars.Size = 1 Then
                        DumpDebug vbTab & Bars.Prop(eBARS_Symbol) & ";" & Bars.Prop(eBARS_PeriodicityStr) & ": Bars(-1)=" & DateLogString(Bars(eBARS_DateTime, Bars.Size - 1)) & "; Last=" & Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1))
                    Else
                        DumpDebug vbTab & Bars.Prop(eBARS_Symbol) & ";" & Bars.Prop(eBARS_PeriodicityStr) & ": No Bars"
                    End If
                Next lIndex
            End If
        End If
    End If
        
    ' 03/01/2016 DAJ: If this was done within 15 minutes of the next scheduled sanity check ( or this happens to be the next scheduled sanity check ),
    ' then add one hour to the next sanity check...
    ' 03/07/2016 DAJ: Moved this out of the if block above so that we update the next sanity check even if we didn't generate orders...
    If Active Then
        If CurrentTime(m.Bars.Prop(eBARS_ExchangeTimeZoneInf), m.strSymbol, True) > (NextSanityCheck - (15# / 1440#)) Then
            NextSanityCheck = NextSanityCheck + (1# / 24#)
        End If
    End If
    
    GenerateOrders = iReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.GenerateOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SyncPosition
'' Description: Synchronize the current position with the strategy position
'' Inputs:      Triggered by a New Bar?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SyncPosition(ByVal bNewBar As Boolean)
On Error GoTo ErrSection:

    Dim lStrategyPos As Long            ' Strategy Position

    DumpDebug "SyncPosition - NewBar: " & Str(bNewBar) & ", Current Position: " & Str(m.lCurrentPosition) & ", Current Strategy Position: " & Str(m.lStrategyPos) & ", Previous Strategy Position: " & Str(m.lPrevStrategyPos)

    If bNewBar And (m.lStrategyPos <> m.lPrevStrategyPos) Then
        lStrategyPos = m.lPrevStrategyPos
    Else
        lStrategyPos = m.lStrategyPos
    End If
    
    If m.lCurrentPosition = 0 And lStrategyPos <> 0 Then
        m.bIgnoreExits = True
    ElseIf m.lCurrentPosition <> lStrategyPos Then
        m.bIgnoreExits = False
        ClosePosition False
    Else
        m.bIgnoreExits = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.SyncPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClosePosition
'' Description: Submit an order to close out the current position
'' Inputs:      Confirm first?, On Close?, Wait for it?, Timeout
'' Returns:     True if position close attempted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ClosePosition(Optional ByVal bConfirm As Boolean = True, Optional ByVal bOnClose As Boolean = False, Optional ByVal bWait As Boolean = False, Optional ByVal lDialogTimeout As Long = 0&) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim Order As New cPtOrder           ' Temporary order object
    Dim strReturn As String             ' Return from the infbox
    Dim lTimeOut As Long                ' Timeout
    Dim strMessage As String            ' Message to show to the user
                    
    bReturn = False
    If m.lCurrentPosition <> 0 Then
        strReturn = "E"
        If bConfirm Then
            strMessage = "Do you wish to exit your|" & g.Broker.TextPosition(m.lCurrentPosition) & "|position in|" & m.strName & "?"
            strReturn = InfBox(strMessage, "?", "+-KEEP Pos.|EXIT Pos.", "Position Exit Confirmation", , lDialogTimeout)
            DumpDebug "User answered '" & strReturn & "' to exiting the position"
        End If
        
        If strReturn = "E" Then
            Set Order = New cPtOrder
            With Order
                .AccountID = m.lAccountID
                .AutoTradeItemID = m.lAutoTradeItemID
                .Buy = (m.lCurrentPosition < 0)
                .Enter = False
                .LimitPrice = kNullData
                .OrderDate = .BrokerDate(CurrentTime("", m.strSymbol, True))
                'If bOnClose Then
                '    .OrderType = eTT_OrderType_MarketOnClose
                '    .Expiration = ExpirationDate * -1&
                'Else
                    .OrderType = eTT_OrderType_Market
                    .Expiration = ExpirationDate * -1&
                'End If
                .Quantity = Abs(m.lCurrentPosition)
                .StopPrice = kNullData
                .SymbolOrSymbolID = PositionSymbolOrID
                .IsAutomated = True
                
                .Save
            End With
            
            ClosingPosition = True
            ClosingQuantity = m.lCurrentPosition * -1&
            
            PlaceOrder Order, True, "to close position"
                        
            If g.Broker.AutoJournalAutomated Then
                g.JournalBridge.CreateOrderJournal Order.OrderID, CurrentTime, "Automated Trading Strategy (" & m.strName & ") placing this order to close the position"
            End If
            
            If bWait Then
                lTimeOut = 0&
                Do While (m.lCurrentPosition <> 0) And (lTimeOut < 30&)
                    Sleep 1
                    lTimeOut = lTimeOut + 1&
                Loop
            End If
            
            bReturn = True
        End If
    End If
    
    ClosePosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.ClosePosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EngineStringToOrder
'' Description: Convert an engine order to a cPtOrder
'' Inputs:      Engine Order, Allow OnClose Orders?
'' Returns:     Order Object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EngineStringToOrder(ByVal strEngineString As String, Optional ByVal bAllowOnClose As Boolean = True) As cPtOrder
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order to return
    Dim astrOrder As New cGdArray       ' Array of order information from string
    Dim lQuantity As Long               ' Quantity for the engine order
    
    astrOrder.SplitFields strEngineString, vbTab
    
    With Order
        .AccountID = m.lAccountID
        .AutoTradeItemID = m.lAutoTradeItemID
        If UseGtcOrder Then
            .Expiration = 0
        Else
            .Expiration = ExpirationDate * -1&
        End If
        .OrderDate = .BrokerDate(CurrentTime("", m.strSymbol))
        .SymbolOrSymbolID = TradeSymbolOrID
        .IsAutomated = True
        
        lQuantity = CLng(Val(astrOrder(OrderField(eGDOrderField_NumContracts))))
        
        Select Case astrOrder(OrderField(eGDOrderField_Position))
            Case "EL"
                .Buy = True
                .Enter = True
                If m.lCurrentPosition < 0 Then
                    .Quantity = Abs(m.lCurrentPosition) + lQuantity
                Else
                    .Quantity = lQuantity
                End If
                
            Case "XL"
                .Buy = False
                .Enter = False
                If lQuantity <= Abs(m.lCurrentPosition) Then
                    .Quantity = lQuantity
                Else
                    .Quantity = Abs(m.lCurrentPosition)
                End If
                
            Case "ES"
                .Buy = False
                .Enter = True
                If m.lCurrentPosition > 0 Then
                    .Quantity = Abs(m.lCurrentPosition) + lQuantity
                Else
                    .Quantity = lQuantity
                End If
            
            Case "XS"
                .Buy = True
                .Enter = False
                If lQuantity <= Abs(m.lCurrentPosition) Then
                    .Quantity = lQuantity
                Else
                    .Quantity = Abs(m.lCurrentPosition)
                End If
        
        End Select
        
        Select Case astrOrder(OrderField(eGDOrderField_OrderType))
            Case "M"
                .OrderType = eTT_OrderType_Market
                .LimitPrice = kNullData
                .StopPrice = kNullData
                
            Case "S"
                If astrOrder(OrderField(eGDOrderField_Offset1)) = "0" And astrOrder(OrderField(eGDOrderField_Price1)) = "ONB" Then
                    .OrderType = eTT_OrderType_Market
                    .LimitPrice = kNullData
                    .StopPrice = kNullData
                Else
                    .OrderType = eTT_OrderType_Stop
                    .LimitPrice = kNullData
                    .StopPrice = m.Bars.RoundToPrice(Val(astrOrder(OrderField(eGDOrderField_Price1))))
                End If
                
            Case "L"
                If astrOrder(OrderField(eGDOrderField_Offset1)) = "0" And astrOrder(OrderField(eGDOrderField_Price1)) = "ONB" Then
                    .OrderType = eTT_OrderType_Market
                    .LimitPrice = kNullData
                    .StopPrice = kNullData
                Else
                    .OrderType = eTT_OrderType_Limit
                    .LimitPrice = m.Bars.RoundToPrice(Val(astrOrder(OrderField(eGDOrderField_Price1))))
                    .StopPrice = kNullData
                End If
            
            Case "MOC"
                If bAllowOnClose Then
                    .OrderType = eTT_OrderType_MarketOnClose
                Else
                    .OrderType = eTT_OrderType_Market
                End If
                .LimitPrice = kNullData
                .StopPrice = kNullData
                
            Case "SCO"
                If astrOrder(OrderField(eGDOrderField_Offset1)) = "0" And astrOrder(OrderField(eGDOrderField_Price1)) = "ONB" Then
                    If bAllowOnClose Then
                        .OrderType = eTT_OrderType_MarketOnClose
                    Else
                        .OrderType = eTT_OrderType_Market
                    End If
                    .LimitPrice = kNullData
                    .StopPrice = kNullData
                Else
                    If bAllowOnClose Then
                        .OrderType = eTT_OrderType_StopCloseOnly
                    Else
                        .OrderType = eTT_OrderType_Stop
                    End If
                    .LimitPrice = kNullData
                    .StopPrice = m.Bars.RoundToPrice(Val(astrOrder(OrderField(eGDOrderField_Price1))))
                End If
                
            Case "LCO"
                If astrOrder(OrderField(eGDOrderField_Offset1)) = "0" And astrOrder(OrderField(eGDOrderField_Price1)) = "ONB" Then
                    If bAllowOnClose Then
                        .OrderType = eTT_OrderType_MarketOnClose
                    Else
                        .OrderType = eTT_OrderType_Market
                    End If
                    .LimitPrice = kNullData
                    .StopPrice = kNullData
                Else
                    If bAllowOnClose Then
                        .OrderType = eTT_OrderType_LimitCloseOnly
                    Else
                        .OrderType = eTT_OrderType_Limit
                    End If
                    .LimitPrice = m.Bars.RoundToPrice(Val(astrOrder(OrderField(eGDOrderField_Price1))))
                    .StopPrice = kNullData
                End If
            
            Case "SWL"
                .OrderType = eTT_OrderType_StopWithLimit
                .LimitPrice = m.Bars.RoundToPrice(Val(astrOrder(OrderField(eGDOrderField_Price2))))
                .StopPrice = m.Bars.RoundToPrice(Val(astrOrder(OrderField(eGDOrderField_Price1))))
                
            Case "SWLCO"
                If bAllowOnClose Then
                    .OrderType = eTT_OrderType_StopWithLimitCloseOnly
                Else
                    .OrderType = eTT_OrderType_StopWithLimit
                End If
                .LimitPrice = m.Bars.RoundToPrice(Val(astrOrder(OrderField(eGDOrderField_Price2))))
                .StopPrice = m.Bars.RoundToPrice(Val(astrOrder(OrderField(eGDOrderField_Price1))))
        
        End Select
    End With
    
    Set EngineStringToOrder = Order

ErrExit:
    Set astrOrder = Nothing
    Set Order = Nothing
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.EngineStringToOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentStrategyPosition
'' Description: Convert the current position line out of the next bar report
'' Inputs:      Current Position String
'' Returns:     Current Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CurrentStrategyPosition(ByVal strCurrentPositions As String) As Long
On Error GoTo ErrSection:

    Dim lQuantity As Long               ' Quantity long or short
    Dim lReturn As Long                 ' Return value
    
#If 1 Then
    Dim astrEntries As cGdArray         ' Entries
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrEntry As cGdArray           ' Entry
    
    Set astrEntries = New cGdArray
    astrEntries.SplitFields strCurrentPositions, "|"
    
    lReturn = 0&
    For lIndex = 0 To astrEntries.Size - 1
        Set astrEntry = New cGdArray
        astrEntry.SplitFields astrEntries(lIndex), vbTab
        
        lQuantity = CLng(Val(astrEntry(1)))
        Select Case UCase(astrEntry(0))
            Case "L"
                lReturn = lReturn + lQuantity
            Case "S"
                lReturn = lReturn - lQuantity
        End Select
    Next lIndex
#Else
    Select Case UCase(Parse(strCurrentPositions, vbTab, 1))
        Case "N"
            lReturn = 0
        Case "L"
            lQuantity = CLng(Val(Parse(strCurrentPositions, vbTab, 2)))
            lReturn = lQuantity
        Case "S"
            lQuantity = CLng(Val(Parse(strCurrentPositions, vbTab, 2)))
            lReturn = lQuantity * -1&
    End Select
#End If
    
    CurrentStrategyPosition = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.CurrentStrategyPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentStrategyPositionFromTrades
'' Description: Determine the current position based on the trade-by-trade array
'' Inputs:      Trade-by-Trade array
'' Returns:     Current Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CurrentStrategyPositionFromTrades(ByVal astrTrades As cGdArray) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrLine As cGdArray            ' Line in the array split out into fields
    
    lReturn = 0&
    For lIndex = astrTrades.Size - 1 To 1 Step -1&
        Set astrLine = New cGdArray
        astrLine.SplitFields astrTrades(lIndex), vbTab
        
        If astrLine(12) = "(null)" Then
            If UCase(astrLine(0)) = "L" Then
                lReturn = lReturn + 1&
            Else
                lReturn = lReturn - 1&
            End If
        Else
            Exit For
        End If
    Next lIndex
    
    CurrentStrategyPositionFromTrades = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.CurrentStrategyPositoinFromTrades"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MarketIsOpen
'' Description: Is the market open for the symbol?
'' Inputs:      None
'' Returns:     True if Open, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MarketIsOpen() As Boolean
On Error GoTo ErrSection:

    Dim dNow As Double                  ' Time right now
    Dim dSessionNow As Double           ' Time right now with session date
    Dim dSessionStart As Double         ' Time the session started
    Dim dSessionEnd As Double           ' Time the session ends
    Dim dCrossOver As Double            ' Crossover time

    dNow = ConvertTimeZone(CurrentTime, "", m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
    dSessionNow = dNow
    dCrossOver = Val(Int(dSessionNow)) + (m.Bars.Prop(eBARS_CrossoverTime) / 1440#)
    If dSessionNow > dCrossOver Then dSessionNow = dSessionNow + 1
    
    dSessionStart = Val(Int(dSessionNow)) + (m.Bars.Prop(eBARS_StartTime) / 1440#)
    dSessionEnd = Val(Int(dSessionNow)) + (m.Bars.Prop(eBARS_EndTime) / 1440#)
    If dSessionStart >= dSessionEnd Then dSessionStart = dSessionStart - 1

    If IsWeekday(dSessionNow) And (dNow >= dSessionStart) And (dNow <= dSessionEnd) Then
        MarketIsOpen = True
    Else
        MarketIsOpen = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.MarketIsOpen"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetOrderQuantities
'' Description: Set the quantity on the engine orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetOrderQuantities()
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of the delimited string
    Dim lIndex As Long                  ' Index into a for loop
    Dim lQuantity As Long               ' Quantity the order should be
    
    If False Then
        For lIndex = 0 To m.astrOrders.Size - 1
            astrOrder.SplitFields m.astrOrders(lIndex), vbTab
            Select Case astrOrder(1)
                Case "EL", "ES"
                    lQuantity = Abs(m.lCurrentPosition) + m.lQtyNextEntry
                Case "XL", "XS"
                    lQuantity = Abs(m.lCurrentPosition)
            End Select
            astrOrder(8) = Str(lQuantity)
            m.astrOrders(lIndex) = astrOrder.JoinFields(vbTab)
        Next lIndex
    End If

ErrExit:
    Set astrOrder = Nothing
    Exit Sub
    
ErrSection:
    Set astrOrder = Nothing
    RaiseError "cAutoTradeItem.SetOrderQuantities"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanPerformOrderAction
'' Description: Determine if we can perform an order action
'' Inputs:      Broker status
'' Returns:     True if OK to perform action, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CanPerformOrderAction(strBrokerStatus As String) As Boolean
On Error GoTo ErrSection:

    Dim nStatus As eGDConnectionStatus  ' Connection status for the account
    Dim bPositionVerify As Boolean      ' Position verify for the broker
    
    nStatus = g.Broker.ConnectionStatusForAccount(m.lAccountID)
    bPositionVerify = g.Broker.PositionVerify(m.nBroker)
    strBrokerStatus = "Status = '" & g.BrokerEnums.ConnectionStatusToString(nStatus) & "'; Waiting for Position Verify = '" & Str(bPositionVerify) & "'"
    
    CanPerformOrderAction = (nStatus = eGDConnectionStatus_Connected) And (bPositionVerify = False)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.CanPerformOrderAction"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrder
'' Description: Submit the given order
'' Inputs:      Order, Reason, Log to add to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SubmitOrder(ByVal Order As cPtOrder, ByVal strReason As String, Optional strLog As String = "")
On Error GoTo ErrSection:

    Dim strText As String               ' Text to send to the logs
    Dim strStatus As String             ' Broker status
    Dim strOrderText As String          ' Order text for the order
    Dim dCurrentPrice As Double         ' Current market price for the order symbol
    Dim bWrongSide As Boolean           ' Are we on the wrong side of the market?

    strOrderText = Order.OrderText(True, True, True)
    strText = vbTab & "Submit Order: '" & strOrderText & "' because '" & strReason & "' (Status = " & OrderStatus(Order.Status) & ")"
    If Len(strLog) = 0 Then
        DumpDebug strText
    End If
    
    If CanPerformOrderAction(strStatus) = False Then
        If Len(strLog) = 0 Then
            DumpDebug vbTab & vbTab & "Could not submit order: " & strStatus
        Else
            strText = strText & vbCrLf & vbTab & vbTab & "Could not submit order: " & strStatus
        End If
    Else
        ' DAJ 09/16/2015: If the order is a stop order, but is on the wrong side of the market, change it to a market order.  If the order
        ' is a stop-with-limit order, but is on the wrong side of the market, change it to a limit order...
        If (Order.StopPrice <> 0) Then
            dCurrentPrice = g.RealTime.LastKnownPrice(Order.SymbolOrSymbolID)
            
            If Order.Buy Then
                bWrongSide = (dCurrentPrice >= Order.StopPrice)
            Else
                bWrongSide = (dCurrentPrice <= Order.StopPrice)
            End If
            
            If bWrongSide Then
                If (Order.OrderType = eTT_OrderType_Stop) Then
                    Order.StopPrice = 0#
                    Order.LimitPrice = 0#
                    Order.OrderType = eTT_OrderType_Market
                    
                    DumpDebug vbTab & vbTab & "Order changed from '" & strOrderText & "' to '" & Order.OrderText(True, True, True) & "' because it was on the wrong side of the market"
                
                ElseIf (Order.OrderType = eTT_OrderType_StopWithLimit) Then
                    Order.StopPrice = 0#
                    Order.OrderType = eTT_OrderType_Limit
                    
                    DumpDebug vbTab & vbTab & "Order changed from '" & strOrderText & "' to '" & Order.OrderText(True, True, True) & "' because it was on the wrong side of the market"
                End If
            End If
        End If
        
        g.Broker.BrokerDebug m.nBroker, m.strName & strText
        mTradeTracker.SubmitOrder Order
    End If
    
    If Len(strLog) > 0 Then
        strLog = strLog & vbCrLf & strText & vbCrLf
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.SubmitOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Cancel the given order
'' Inputs:      Order, Reason, Log to add to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelOrder(ByVal Order As cPtOrder, ByVal strReason As String, Optional strLog As String = "")
On Error GoTo ErrSection:

    Dim strText As String               ' Text to send to the logs
    Dim strStatus As String             ' Broker status

    strText = vbTab & "Cancel Order: '" & Order.OrderText(True, True, True) & "' because '" & strReason & "' (Status = " & OrderStatus(Order.Status) & ")"
    If Len(strLog) = 0 Then
        DumpDebug strText
    End If
    
    If (CanPerformOrderAction(strStatus) = False) And (HasBeenSent(Order.Status) = True) Then
        If Len(strLog) = 0 Then
            DumpDebug vbTab & vbTab & "Could not cancel order: " & strStatus
        Else
            strText = strText & vbCrLf & vbTab & vbTab & "Could not cancel order: " & strStatus
        End If
    ElseIf (Order.OrderType = eTT_OrderType_Market) And (g.Broker.IsLiveAccount(Order.Broker) = True) And (HasBeenSent(Order.Status) = True) Then
        If Len(strLog) = 0 Then
            DumpDebug vbTab & vbTab & "Could not cancel order because it is a market order"
        Else
            strText = strText & vbCrLf & vbTab & vbTab & "Could not cancel order because it is a market order"
        End If
    Else
        g.Broker.BrokerDebug m.nBroker, m.strName & strText
        mTradeTracker.CancelOrder Order, False
    End If
    
    If Len(strLog) > 0 Then
        strLog = strLog & vbCrLf & strText & vbCrLf
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.CancelOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Park the given order
'' Inputs:      Order, Reason, Log to add to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ParkOrder(ByVal Order As cPtOrder, ByVal strReason As String, Optional strLog As String = "")
On Error GoTo ErrSection:

    Dim strText As String               ' Text to send to the logs
    Dim strStatus As String             ' Broker status

    strText = vbTab & "Park Order: '" & Order.OrderText(True, True, True) & "' because '" & strReason & "' (Status = " & OrderStatus(Order.Status) & ")"
    If Len(strLog) = 0 Then
        DumpDebug strText
    End If
    
    If CanPerformOrderAction(strStatus) = False Then
        If Len(strLog) = 0 Then
            DumpDebug vbTab & vbTab & "Could not park order: " & strStatus
        Else
            strText = strText & vbCrLf & vbTab & vbTab & "Could not park order: " & strStatus
        End If
    Else
        g.Broker.BrokerDebug m.nBroker, m.strName & strText
        mTradeTracker.ParkOrder Order
    End If
    
    If Len(strLog) > 0 Then
        strLog = strLog & vbCrLf & strText & vbCrLf
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.ParkOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelSubmittedOrders
'' Description: Cancel submitted orders
'' Inputs:      Just Non-Reversal Entries?, ID not to Cancel, Wait until done?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CancelSubmittedOrders(ByVal bJustNonReversalEntries As Boolean, Optional ByVal lNotThisID As Long = -1&, Optional ByVal bWait As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Orders As cPtOrders             ' Collection of working orders
    Dim lTimeOut As Long                ' Timeout variable
    
    DumpDebug "Begin CancelSubmittedOrders(" & Str(bJustNonReversalEntries) & ", " & Str(lNotThisID) & ")"
    
    lReturn = 0&
    Set Orders = g.Broker.WorkingOrders(m.lAccountID, PositionSymbolOrID, m.lAutoTradeItemID)
    If Not Orders Is Nothing Then
        For lIndex = Orders.Count To 1 Step -1
            If (OrderIsNonReversalEntry(Orders(lIndex)) = True) Or (bJustNonReversalEntries = False) Then
                If Orders(lIndex).OrderID <> lNotThisID Then
                    If ConfirmOrder(Orders(lIndex), True) Then
                        lReturn = lReturn + 1&
                        CancelOrder Orders(lIndex), "submitted orders are being cancelled"
                        
                        If g.Broker.AutoJournalAutomated Then
                            g.JournalBridge.CreateOrderJournal Orders(lIndex).OrderID, CurrentTime, "Automated Trading Strategy (" & m.strName & ") cancelled the order"
                        End If
                    End If
                End If
            End If
        Next lIndex
        
        If bWait Then
            lTimeOut = 0&
            Do While (g.Broker.HasWorkingOrdersForAutoTrade(Me) > 0) And (lTimeOut < 30#)
                Sleep 1#
                lTimeOut = lTimeOut + 1&
            Loop
        End If
    End If
    
    DumpDebug "End CancelSubmittedOrders"
    
    CancelSubmittedOrders = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.CancelSubmittedOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionsMatch
'' Description: Check the current position against the strategy position
'' Inputs:      None
'' Returns:     True if Match, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PositionsMatch() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    If ((m.lStrategyPos >= 0) And (m.lCurrentPosition < 0)) Then
        bReturn = False
    ElseIf ((m.lStrategyPos <= 0) And (m.lCurrentPosition > 0)) Then
        bReturn = False
    ElseIf (m.lStrategyPos <> 0) And (m.lCurrentPosition = 0) Then
        bReturn = False
    ElseIf m.Strategy.Pyramid = False Then
        bReturn = True
    Else
        bReturn = (m.lCurrentPosition = (m.lStrategyPos * m.lQtyNextEntry))
    End If
    
    PositionsMatch = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.PositionsMatch"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SyncOrders
'' Description: Synchronize the orders with the engine orders
'' Inputs:      Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SyncOrders(ByVal strSource As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Order As New cPtOrder           ' Temporary order object
    Dim bFound As Boolean               ' Did we find the order?
    Dim lQuantity As Long               ' Quantity from the string
    Dim Orders As cPtOrders             ' Working orders collection
    Dim strOrder As String              ' Order from the array

    DumpDebug "Begin SyncOrders( '" & strSource & "' )"

    ' 1) Cancel obsolete orders and edit quantity on existing orders...
    Set Orders = g.Broker.WorkingOrders(m.lAccountID, PositionSymbolOrID, m.lAutoTradeItemID)
    If Not Orders Is Nothing Then
        For lIndex = Orders.Count To 1 Step -1
            bFound = False
            For lIndex2 = 0 To m.astrOrders.Size - 1
                strOrder = SetQuantityOnEngineString(m.astrOrders(lIndex2))
                
                If SameOrder(Orders(lIndex).EngineString(True), strOrder, True, Orders(lIndex).SymbolOrSymbolID) Then
                    lQuantity = CLng(Val(Parse(strOrder, vbTab, 9)))
                    If lQuantity = 0 Then
                        If ConfirmOrder(Orders(lIndex), True) Then
                            CancelOrder Orders(lIndex), "quantity is now zero"
                        
                            If g.Broker.AutoJournalAutomated Then
                                g.JournalBridge.CreateOrderJournal Orders(lIndex).OrderID, CurrentTime, "Automated Trading Strategy (" & m.strName & ") cancelled the order because the quantity is zero"
                            End If
                        End If
                    'ElseIf Orders(lIndex).Quantity <> lQuantity Then
                    ElseIf Orders(lIndex).RemainingQuantity <> lQuantity Then
                        Orders(lIndex).Quantity = lQuantity
                        Orders(lIndex).Save
                        
                        PlaceOrder Orders(lIndex), False, "quantity has changed to " & Str(lQuantity)
                        
                        If g.Broker.AutoJournalAutomated Then
                            g.JournalBridge.CreateOrderJournal Orders(lIndex).OrderID, CurrentTime, "Automated Trading Strategy (" & m.strName & ") updating quantity of this order to " & Str(lQuantity)
                        End If
                    End If
                    bFound = True
                    Exit For
                End If
            Next lIndex2
                
            If Not bFound Then
                If (UCase(strSource) <> "FILL") Or (Orders(lIndex).Status <> eTT_OrderStatus_Partial) Then
                    If ConfirmOrder(Orders(lIndex), True) Then
                        CancelOrder Orders(lIndex), "because it no longer exists"
                        
                        If g.Broker.AutoJournalAutomated Then
                            g.JournalBridge.CreateOrderJournal Orders(lIndex).OrderID, CurrentTime, "Automated Trading Strategy (" & m.strName & ") cancelled the order because strategy no longer has the order working"
                        End If
                    End If
                End If
            End If
        Next lIndex
    
        ' 2) Add New Orders...
        For lIndex = 0 To m.astrOrders.Size - 1
            strOrder = SetQuantityOnEngineString(m.astrOrders(lIndex))
            
            bFound = False
            For lIndex2 = 1 To Orders.Count
                If SameOrder(strOrder, Orders(lIndex2).EngineString(True), True, Orders(lIndex2).SymbolOrSymbolID) Then
                    bFound = True
                    Exit For
                End If
            Next lIndex2
            
            If Not bFound Then
                Set Order = EngineStringToOrder(strOrder)
                If Order.Quantity > 0 Then
                    Order.Status = eTT_OrderStatus_Parked
                    Order.Save
                    
                    PlaceOrder Order, False, "strategy has generated the order"
                
                    If g.Broker.AutoJournalAutomated Then
                        g.JournalBridge.CreateOrderJournal Order.OrderID, CurrentTime, "Automated Trading Strategy (" & m.strName & ") placing this order because strategy has generated the order"
                    End If
                Else
                    DumpDebug vbTab & "Didn't place order: '" & Order.OrderText(True, True, True) & "' because strategy has generated the order but the order quantity is zero"
                End If
            End If
        Next lIndex
    End If
    
    DumpDebug "End SyncOrders"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.SyncOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SameOrder
'' Description: Determine whether the two orders passed in are the "same"
'' Inputs:      Order 1, Order 2, Check Quantity?, Order Symbol
'' Returns:     True if "Same", False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SameOrder(ByVal strOrder1 As String, ByVal strOrder2 As String, ByVal bCheckQuantity As Boolean, ByVal vOrderSymbolOrId As Variant) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim vTradeSymbolOrID As Variant     ' Trade symbol or ID
    
    bReturn = False
    vTradeSymbolOrID = TradeSymbolOrID
    
    ' Check the symbol...
    If vOrderSymbolOrId = vTradeSymbolOrID Then
        ' Check the position (EL, ES, XL, XS)...
        If Parse(strOrder1, vbTab, 2) = Parse(strOrder2, vbTab, 2) Then
            ' Check the Order Type (M, S, L, etc)...
            If Parse(strOrder1, vbTab, 4) = Parse(strOrder2, vbTab, 4) Then
                ' Check Price1...
                If SameOrderPrice(Parse(strOrder1, vbTab, 5), Parse(strOrder2, vbTab, 5)) Then
                    ' Check Price2...
                    If SameOrderPrice(Parse(strOrder1, vbTab, 7), Parse(strOrder2, vbTab, 7)) Then
                        If bCheckQuantity Then
                            ' Check Quantity...
                            If Parse(strOrder1, vbTab, 9) = Parse(strOrder2, vbTab, 9) Then
                                bReturn = True
                            End If
                        Else
                            bReturn = True
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If bReturn = False Then
        DebugLog "---------------"
        DebugLog "Order Symbol = '" & GetSymbol(vOrderSymbolOrId) & "'; Trade Symbol = '" & GetSymbol(vTradeSymbolOrID) & "'"
        DebugLog "Order1 = '" & strOrder1 & "'"
        DebugLog "Order2 = '" & strOrder2 & "'"
        DebugLog "---------------"
    End If
    
    SameOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.SameOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SameOrderPrice
'' Description: Determine whether the two orders passed in are the "same"
'' Inputs:      Order 1, Order 2, Check Quantity?
'' Returns:     True if "Same", False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SameOrderPrice(ByVal strPrice1 As String, ByVal strPrice2 As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim dPrice1 As Double               ' Value of Price1 rounded to the nearest tick
    Dim dPrice2 As Double               ' Value of Price2 rounded to the nearest tick
    
    bReturn = False
    If strPrice1 = strPrice2 Then
        bReturn = True
    ElseIf (Len(strPrice1) = 0) And (strPrice2 = "-999999") Then
        bReturn = True
    ElseIf (strPrice1 = "-999999") And (Len(strPrice2) = 0) Then
        bReturn = True
    Else
        dPrice1 = m.Bars.RoundToPrice(Val(strPrice1))
        dPrice2 = m.Bars.RoundToPrice(Val(strPrice2))
        
        bReturn = (dPrice1 = dPrice2)
    End If
    
    SameOrderPrice = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.SameOrderPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PlaceOrder
'' Description: Either submit or park the order (unless quantity is zero)
'' Inputs:      Order to Place, Submit?, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PlaceOrder(Order As cPtOrder, ByVal bSubmit As Boolean, ByVal strReason As String)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position to put order in the array
    Dim lCurrentPosition As Long        ' Current position
    Dim lResultingPosition As Long      ' Position after this order
    Dim strError As String              ' Reason for not placing order

    strError = ""

    DumpDebug "PlaceOrder(" & Str(bSubmit) & "): '" & Order.OrderText(True, True, True) & "' because '" & strReason & "'"
    If Order.Quantity > 0 Then
        lCurrentPosition = GetCurrentPositionFromBrokerInfo
        If Order.Buy Then
            lResultingPosition = lCurrentPosition + Order.Quantity
        Else
            lResultingPosition = lCurrentPosition - Order.Quantity
        End If
        
        If (Pyramid = True) Or (lResultingPosition <= QtyNextEntry) Then
            If Len(Order.GenesisOrderID) = 0 Then
                Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Order.AccountID))
                Order.Save
            End If
            
            If bSubmit Then
                If (Order.Buy = True And m.lCurrentPosition <= 0) Or (Order.Buy = False And m.lCurrentPosition >= 0) Then
                    If IsOnCloseOrder(Order) = False Then
                        ChangeMarketToLimit Order
                        SubmitOrder Order, strReason
                        m.bOrderSubmitted = True
                    Else
                        strError = "it is an on-close order"
                    End If
                End If
            Else
                If m.SubmittedOrders.Exists(Str(Order.OrderID)) Then
                    Set m.SubmittedOrders(Str(Order.OrderID)) = Order
                Else
                    m.SubmittedOrders.Add Order, Str(Order.OrderID)
                End If
                
                RefreshOrder Order
            
                g.Broker.AddOrder Order
            End If
        Else
            strError = "The resulting position is greater than the quantity next entry"
            Disable False, False, strError, True
        End If
    Else
        strError = "the order quantity is zero"
    End If
    
    If Len(strError) > 0 Then
        DumpDebug vbTab & "Didn't submit order: '" & Order.OrderText(True, True, True) & "' because '" & strError & " '; Status = '" & OrderStatus(Order.Status) & "'; QtyNextEntry = " & Str(m.lQtyNextEntry) & "; Resulting Position = " & Str(lResultingPosition)
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAutoTradeItem.PlaceOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExpirationDate
'' Description: Determine the expiration date for a non-GTC order
'' Inputs:      None
'' Returns:     ExpirationDate
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ExpirationDate() As Long
On Error GoTo ErrSection:

    Dim strExchangeTimeInf As String    ' String of exchange time zone information
    Dim dNowInExchangeTime As Double    ' The current time in exchange time zone
    Dim dSessionEnd As Double           ' Session End time in exchange time zone
    Dim lReturn As Long                 ' Return value
    
    strExchangeTimeInf = m.Bars.Prop(eBARS_ExchangeTimeZoneInf)
    dNowInExchangeTime = ConvertTimeZone(CurrentTime, , strExchangeTimeInf)
    dSessionEnd = Int(dNowInExchangeTime) + (m.Bars.Prop(eBARS_EndTime) / 1440#)
    
    If dNowInExchangeTime > dSessionEnd Then
        lReturn = Int(ConvertTimeZone(dNowInExchangeTime + 1, strExchangeTimeInf, "NY"))
        Do While Not IsWeekday(lReturn)
            lReturn = lReturn + 1
        Loop
    Else
        lReturn = Int(ConvertTimeZone(dNowInExchangeTime, strExchangeTimeInf, "NY"))
    End If
    
    ExpirationDate = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.ExpirationDate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitString
'' Description: Determine the submit string from the order
'' Inputs:      Order
'' Returns:     Submit String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SubmitString(ByVal Order As cPtOrder) As String
On Error GoTo ErrSection:

    Dim astrReturn As New cGdArray      ' Array of information to join
    
    astrReturn(0) = Str(Order.OrderID)
    Select Case Order.OrderType
        Case eTT_OrderType_Stop
            If Order.Buy Then
                astrReturn(1) = ">=" & Str(Order.StopPrice)
            Else
                astrReturn(1) = "<=" & Str(Order.StopPrice)
            End If
            
        Case eTT_OrderType_Limit
            If Order.Buy Then
                astrReturn(1) = "<=" & Str(Order.LimitPrice)
            Else
                astrReturn(1) = ">=" & Str(Order.LimitPrice)
            End If
        
        Case eTT_OrderType_StopWithLimit
            If Order.Buy Then
                astrReturn(1) = "<=" & Str(Order.LimitPrice) & ",>=" & Str(Order.StopPrice)
            Else
                astrReturn(1) = "<=" & Str(Order.StopPrice) & ",>=" & Str(Order.LimitPrice)
            End If
    End Select
    astrReturn(2) = Order.OrderText
    
    SubmitString = astrReturn.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.SubmitString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpOrders
'' Description: Dump the orders array to a file and to the test form
'' Inputs:      Calling Function
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpOrders(ByVal strCallingFunction As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim fh As Integer                   ' File handle
    
    If FileExist(AddSlash(App.Path) & "DumpOrders.FLG") = True Then
        fh = FreeFile
        Open AddSlash(App.Path) & "Orders.DAJ" For Append As #fh
        
        Print #fh, "========================"
        frmTest2.AddList "========================"
        
        Print #fh, strCallingFunction
        frmTest2.AddList strCallingFunction
        
        For lIndex = 0 To m.astrOrders.Size - 1
            Print #fh, m.astrOrders(lIndex)
            frmTest2.AddList m.astrOrders(lIndex)
        Next lIndex
        
        Print #fh, "========================"
        frmTest2.AddList "========================"
        
        Close #fh
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.DumpOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOneOrder
'' Description: If there is more than one order, make sure only one is
''              submitted at one time
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SubmitOneOrder()
On Error GoTo ErrSection:

    Dim Closest As New cPtOrder         ' Closest order to the market
    Dim NextClosest As New cPtOrder     ' Second closest order to the market
    Dim dMarket As Double               ' Current market price
    Dim lIndex As Long                  ' Index into a for loop
    Dim bCancelPending As Boolean       ' Is there a cancel pending?
    Dim SubmittedOrder As New cPtOrder  ' Currently submitted order
    Dim strOrderText As String          ' Text description of the order
    Dim strToLog As String              ' String to send to the log
    Static strPrevToLog As String       ' Last string sent to the log
    
    DoSubmitOneOrder = False
    If (m.SubmittedOrders.Count > 0) And (m.iRolling = 0) Then
        strToLog = "Begin SubmitOneOrder" & vbCrLf
    
        bCancelPending = False
        dMarket = m.Bars(eBARS_Close, m.Bars.Size - 1)
        
        ' Find the two closest orders to the market...
        Set Closest = m.SubmittedOrders(1)
        If (Closest.Status = eTT_OrderStatus_CancelPending) Or (Closest.Status = eTT_OrderStatus_ParkPending) Then bCancelPending = True
        If Closest.Status <> eTT_OrderStatus_Parked Then Set SubmittedOrder = Closest
        
        For lIndex = 2 To m.SubmittedOrders.Count
            If DiffFromMarket(m.SubmittedOrders(lIndex), dMarket) < DiffFromMarket(Closest, dMarket) Then
                Set NextClosest = Closest
                Set Closest = m.SubmittedOrders(lIndex)
            ElseIf (NextClosest.OrderID = 0&) Or (DiffFromMarket(m.SubmittedOrders(lIndex), dMarket) < DiffFromMarket(NextClosest, dMarket)) Then
                Set NextClosest = m.SubmittedOrders(lIndex)
            End If
            
            If (m.SubmittedOrders(lIndex).Status = eTT_OrderStatus_CancelPending) Or (m.SubmittedOrders(lIndex).Status = eTT_OrderStatus_ParkPending) Then bCancelPending = True
            If m.SubmittedOrders(lIndex).Status <> eTT_OrderStatus_Parked Then Set SubmittedOrder = m.SubmittedOrders(lIndex)
        Next lIndex
        
        strOrderText = Closest.OrderText(True, True, True)
        If (Len(strOrderText) = 0) Or (Len(Closest.GenesisOrderID) = 0) Then
            strToLog = strToLog & vbTab & "Closest:" & vbTab & "Nothing" & vbCrLf
        Else
            strToLog = strToLog & vbTab & "Closest:" & vbTab & "'" & strOrderText & "' (Status = " & OrderStatus(Closest.Status) & ")" & vbCrLf
        End If
        
        strOrderText = NextClosest.OrderText(True, True, True)
        If (Len(strOrderText) = 0) Or (Len(NextClosest.GenesisOrderID) = 0) Then
            strToLog = strToLog & vbTab & "Next Closest:" & vbTab & "Nothing" & vbCrLf
        Else
            strToLog = strToLog & vbTab & "Next Closest:" & vbTab & "'" & strOrderText & "' (Status = " & OrderStatus(Closest.Status) & ")" & vbCrLf
        End If
        
        strOrderText = SubmittedOrder.OrderText(True, True, True)
        If (Len(strOrderText) = 0) Or (Len(SubmittedOrder.GenesisOrderID) = 0) Then
            strToLog = strToLog & vbTab & "Submitted:" & vbTab & "Nothing" & vbCrLf & vbCrLf
        Else
            strToLog = strToLog & vbTab & "Submitted:" & vbTab & "'" & strOrderText & "' (Status = " & OrderStatus(SubmittedOrder.Status) & ")" & vbCrLf & vbCrLf
        End If
        
        ' If there is a cancel pending...
        If bCancelPending Then
            strToLog = strToLog & vbTab & "There is a Cancel Pending" & vbCrLf
            
        ' Otherwise if Closest is already submitted...
        ElseIf Closest.Status <> eTT_OrderStatus_Parked Then
            strToLog = strToLog & vbTab & "Closest is already submitted" & vbCrLf
            
        ' Otherwise if there are currently no submitted orders...
        ElseIf Len(SubmittedOrder.GenesisOrderID) = 0 Then
            strToLog = strToLog & vbTab & "No currently submitted orders" & vbCrLf
            If IsOnCloseOrder(Closest) = False Then
                If ConfirmOrder(Closest, False) Then
                    strToLog = strToLog & vbTab & "Submit Order: '" & Closest.OrderText(True, True, True) & "' because it is the closest to the market" & vbCrLf
                    SubmitOrder Closest, "it is the closest to the market", strToLog
                    m.bOrderSubmitted = True
                End If
            Else
                strToLog = strToLog & vbTab & "Didn't Submit Order: '" & Closest.OrderText(True, True, True) & "' because it is the closest to the market but it is an on-close order" & vbCrLf
            End If
            
        ElseIf m.SubmittedOrders.Count > 1 Then
            ' Otherwise if the next closest order is not the one submitted...
            If NextClosest.Status = eTT_OrderStatus_Parked Then
                If mTradeTracker.OrderIsPending(SubmittedOrder) Then
                    strToLog = strToLog & vbTab & "Next Closest order is not the one submitted, but submitted order is still pending"
                Else
                    strToLog = strToLog & vbTab & "Next Closest order is not the one submitted" & vbCrLf
                    strToLog = strToLog & vbTab & "Park Order: '" & SubmittedOrder.OrderText(True, True, True) & "' because next closest order is not the one submitted (Status = " & OrderStatus(SubmittedOrder.Status) & ")" & vbCrLf
                    ParkOrder SubmittedOrder, "next closest order is not the one submitted", strToLog
                End If
                
            ' Otherwise if the market is not in between the two closest orders...
            ElseIf MarketBetweenOrders(dMarket, Closest, NextClosest) = False Then
                If mTradeTracker.OrderIsPending(SubmittedOrder) Then
                    strToLog = strToLog & vbTab & "Market is not in between the two closest orders but submitted order is still pending"
                Else
                    strToLog = strToLog & vbTab & "Market is not in between the two closest orders" & vbCrLf
                    strToLog = strToLog & vbTab & "Park Order: '" & SubmittedOrder.OrderText(True, True, True) & "' because market is not in between the two closest orders (Status = " & OrderStatus(SubmittedOrder.Status) & ")" & vbCrLf
                    ParkOrder SubmittedOrder, "market is not in between the two closest orders", strToLog
                End If
                
            ' Otherwise if the closest order is closer by more than half...
            ElseIf DiffFromMarket(Closest, dMarket) < (DiffFromMarket(NextClosest, dMarket) * kDiffPct) Then
                If mTradeTracker.OrderIsPending(SubmittedOrder) Then
                    strToLog = strToLog & vbTab & "Closest order is closer by more than half but submitted order is still pending"
                Else
                    strToLog = strToLog & vbTab & "Closest order is closer by more than half" & vbCrLf
                    strToLog = strToLog & vbTab & "Park Order: '" & SubmittedOrder.OrderText(True, True, True) & "' because closest order is closer by more than half (Status = " & OrderStatus(SubmittedOrder.Status) & ")" & vbCrLf
                    ParkOrder SubmittedOrder, "closest order is closer by more than half", strToLog
                End If
            End If
        End If
        
        strToLog = strToLog & "End SubmitOneOrder"
        
        If strToLog <> strPrevToLog Then
            DumpDebug vbCrLf & strToLog
            strPrevToLog = strToLog
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeitem.SubmitOneOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DiffFromMarket
'' Description: Difference between an order price and the market price
'' Inputs:      Order, Market Price
'' Returns:     Difference
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DiffFromMarket(ByVal Order As cPtOrder, ByVal dMarket As Double) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function

    Select Case Order.OrderType
        Case eTT_OrderType_Market
            dReturn = 0#
        
        Case eTT_OrderType_Stop
            dReturn = Abs(Order.StopPrice - dMarket)
        
        Case eTT_OrderType_Limit
            dReturn = Abs(Order.LimitPrice - dMarket)
        
        Case eTT_OrderType_StopWithLimit
            If Order.Buy Then
                If dMarket > Order.StopPrice And dMarket < Order.LimitPrice Then
                    dReturn = 0#
                ElseIf Abs(Order.StopPrice - dMarket) <= Abs(Order.LimitPrice - dMarket) Then
                    dReturn = Abs(Order.StopPrice - dMarket)
                Else
                    dReturn = Abs(Order.LimitPrice - dMarket)
                End If
            Else
                If dMarket > Order.LimitPrice And dMarket < Order.StopPrice Then
                    dReturn = 0#
                ElseIf Abs(Order.StopPrice - dMarket) <= Abs(Order.LimitPrice - dMarket) Then
                    dReturn = Abs(Order.StopPrice - dMarket)
                Else
                    dReturn = Abs(Order.LimitPrice - dMarket)
                End If
            End If
        
        ' Try to make sure that on close orders are never the closest to the market...
        Case eTT_OrderType_MarketOnClose, eTT_OrderType_LimitCloseOnly, eTT_OrderType_StopCloseOnly, eTT_OrderType_StopWithLimitCloseOnly
            dReturn = 99999
    End Select
    
    DiffFromMarket = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.DiffFromMarket"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MarketBetweenOrders
'' Description: Determine if the market is in between orders
'' Inputs:      Market Price, Order1, Order2
'' Returns:     True if Between, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MarketBetweenOrders(ByVal dMarket As Double, ByVal Order1 As cPtOrder, ByVal Order2 As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim dPrice1 As Double               ' Order price of Order1
    Dim dPrice2 As Double               ' Order price of Order2
    
    MarketBetweenOrders = False
    dPrice1 = DiffFromMarket(Order1, 0#)
    dPrice2 = DiffFromMarket(Order2, 0#)
    
    If dPrice1 > 0 And dPrice2 > 0 Then
        If dPrice1 > dPrice2 Then
            If (dMarket < dPrice1) And (dMarket > dPrice2) Then MarketBetweenOrders = True
        Else
            If (dMarket < dPrice2) And (dMarket > dPrice1) Then MarketBetweenOrders = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.MarketBetweenOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConfirmOrder
'' Description: If we want to confirm the order, have the user confirm it
'' Inputs:      Order, Whether it is a Cancel
'' Returns:     True if go ahead, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ConfirmOrder(ByVal Order As cPtOrder, ByVal bCancel As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return from infbox

    ConfirmOrder = True
    If m.bConfirm And Active Then
        If bCancel Then
            strReturn = "A"
            'strReturn = InfBox("Canceling the following order:|" & Order.OrderText & "||Do you want to accept the cancel or stop auto-trading?|", "?", "+-Accept|Stop Trading", "Order Confirmation", , m.lConfirmTimeout)
        Else
            If InStr(m.strConfirmedSubmit, "," & Str(Order.OrderID) & ",") = 0 Then
                strReturn = InfBox("Submitting the following order:|" & Order.OrderText & "||Do you want to accept the trade or stop auto-trading?|", "?", "+-Accept|Stop Trading", "Order Confirmation", , m.lConfirmTimeout)
                DumpDebug "User answered '" & strReturn & "' to submitting '" & Order.OrderText(True, True, True) & "'"
                If strReturn = "A" Then
                    m.strConfirmedSubmit = m.strConfirmedSubmit & "," & Str(Order.OrderID) & ","
                End If
            Else
                strReturn = "A"
            End If
        End If
        
        If strReturn <> "A" Then
            ConfirmOrder = False
            Disable , , "User chose to stop auto trading instead of submitting order", False, , , True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.ConfirmOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisLog
'' Description: Send a string to the log file for the day
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GenesisLog(ByVal strMessage As String)
On Error Resume Next

#If 0 Then

    Dim fh As Integer                   ' File handle to open file with
    fh = FreeFile
    Open AddSlash(App.Path) & "AutoTrade\" & Format(Now, "YYYYMMDD") & ".LOG" For Append Shared As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ") - " & strMessage
        Close #fh
    End If
    
#Else

    Static LogFile As cLogFile
    If LogFile Is Nothing Then
        Set LogFile = New cLogFile
        LogFile.OpenFile AddSlash(App.Path) & "AutoTrade\*.LOG"
    End If
    LogFile.WriteText strMessage
    
#End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CleanGenesisLogs
'' Description: Clean out any Genesis logs that are older than 7 days
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CleanGenesisLogs()
On Error GoTo ErrSection:

    Dim strFile As String               ' Name of the file that was found
    
    strFile = Dir(AddSlash(App.Path) & "AutoTrade\*.LOG")
    Do While Len(strFile) > 0
        If FileDate(strFile) < Date - 7 Then
            KillFile strFile
        End If
        
        strFile = Dir
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.CleanGenesisLogs"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AdjustBarBasedOnLimitPrice
'' Description: Adjust the high/low on the last bar based on the limit price
''              of the filled order in such a way that the engine will believe
''              that we should be filled
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AdjustBarBasedOnLimitPrice(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim dHigh As Double                 ' Current high of the last bar
    Dim dLow As Double                  ' Current low of the last bar
    
    ' If the strategy is marked to force a price through the limit ...
    If m.Strategy.ForceLimitThrough = True Then
        ' ... and the order that we received the fill for was a limit order ...
        If Order.OrderType = eTT_OrderType_Limit Then
            dHigh = m.Bars(eBARS_High, m.Bars.Size - 1)
            dLow = m.Bars(eBARS_Low, m.Bars.Size - 1)
            
            ' ... and the limit price is at the high or the low of the bar (as appropriate) ...
            If (Order.LimitPrice = dHigh) And (Order.Buy = False) Then
                ' ... then fudge the bar so that the engine will recognize a trade through
                m.Bars(eBARS_High, m.Bars.Size - 1) = dHigh + m.Bars.MinMove
            ElseIf (Order.LimitPrice = dLow) And (Order.Buy = True) Then
                ' ... then fudge the bar so that the engine will recognize a trade through
                m.Bars(eBARS_Low, m.Bars.Size - 1) = dLow - m.Bars.MinMove
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.AdjustBarBasedOnLimitPrice"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIsNonReversalEntry
'' Description: Is the given order a non-reversal entry?
'' Inputs:      Order
'' Returns:     True if non-reversal entry, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderIsNonReversalEntry(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for this function
    
    If (Order.Enter = False) Then
        bReturn = False
    ElseIf (m.lCurrentPosition > 0) And (Order.Buy = True) Then
        bReturn = True
    ElseIf (m.lCurrentPosition < 0) And (Order.Buy = False) Then
        bReturn = True
    Else
        bReturn = False
    End If
    
    OrderIsNonReversalEntry = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.OrderIsNonReversalEntry"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOnCloseOrder
'' Description: Is the given order a OnClose order?
'' Inputs:      Order
'' Returns:     True if OnClose, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsOnCloseOrder(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    Select Case Order.OrderType
        Case eTT_OrderType_LimitCloseOnly
            bReturn = True
        Case eTT_OrderType_MarketOnClose
            bReturn = True
        Case eTT_OrderType_StopCloseOnly
            bReturn = True
        Case eTT_OrderType_StopWithLimitCloseOnly
            bReturn = True
        Case Else
            bReturn = False
    End Select
    
    IsOnCloseOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.IsOnCloseOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddMinuteData
'' Description: Add minute data to the secondary markets if not there
'' Inputs:      None
'' Returns:     True if added, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddMinuteData() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Is there minute data
    Dim strSymbol As String             ' Symbol for the secondary market
    Dim strPeriod As String             ' Period for the secondary market
    
    bReturn = False
    bFound = False
    
    For lIndex = 0 To m.astrSecondaryMarkets.Size - 1
        strSymbol = Parse(m.astrSecondaryMarkets(lIndex), ";", 1)
        strPeriod = Parse(m.astrSecondaryMarkets(lIndex), ";", 2)
        
        If strSymbol = m.strSymbol Then
            If InStr(UCase(strPeriod), "MINUTE") <> 0 Then
                bFound = True
                Exit For
            End If
        End If
    Next lIndex
    
    If bFound = False Then
        ' TLB: the XX at the end just flags our LoadBars routine to not load much for this "unused" data
        m.astrSecondaryMarkets.Add m.strSymbol & ";60 MinuteXX"
        bReturn = True
    End If
    
    AddMinuteData = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.AddMinuteData"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcBetween
'' Description: Determine if we are between the 'On Close' time and the start
''              of the next session
'' Inputs:      Current exchange time (out)
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcBetween(Optional dNow As Double)
On Error GoTo ErrSection:

    Dim lSessionDate As Long            ' Session date for the current time
    Dim dOnCloseTimeExch As Double      ' "On-close" time in exchange time
    Dim dStartTimeExch As Double        ' Session start time in exchange time
    Dim dNextStart As Double            ' Next start
    Dim dEndTimeExch As Double          ' Session end time in exchange time
    Dim bBetweenOpen As Boolean         ' Are we between the "on-close" time and the next start time?
    Dim bBetweenClose As Boolean        ' Are we between the "on-close" time and the end time?
    Dim bBetweenSessions As Boolean     ' Are we between the end time and the next start time?
    Static bLogged As Boolean           ' Have we logged the change at least once?
    
    ' Get current time in exchange time...
    dNow = CurrentTime(m.Bars.Prop(eBARS_ExchangeTimeZoneInf), m.Bars.Prop(eBARS_Symbol), True)
    
    ' Get current session date...
    lSessionDate = m.Bars.SessionDateForTime(dNow, True)
    If lSessionDate <= 0 Then
        ' If not a valid normal trading time, then we must be between...
        bBetweenOpen = True
        bBetweenClose = True
        bBetweenSessions = True
    
        If (bBetweenOpen <> m.bBetweenOpen) Or (bLogged = False) Then
            m.bBetweenOpen = bBetweenOpen
            DumpDebug vbTab & "Between ( Open ) changed to " & Str(m.bBetweenOpen) & " ( Current Exchange Time = " & DateFormat(dNow, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", On Close Time = " & DateFormat(m.dOnCloseTimeExch, NO_DATE, HH_MM_SS, AMPM_UPPER) & ", Start Time = " & DateFormat(m.Bars.Prop(eBARS_StartTime) / 1440#, NO_DATE, HH_MM_SS, AMPM_UPPER) & " ) -- Not valid trading time"
        End If
        If (bBetweenClose <> m.bBetweenClose) Or (bLogged = False) Then
            m.bBetweenClose = bBetweenClose
            DumpDebug vbTab & "Between ( Close ) changed to " & Str(m.bBetweenClose) & " ( Current Exchange Time = " & DateFormat(dNow, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", On Close Time = " & DateFormat(m.dOnCloseTimeExch, NO_DATE, HH_MM_SS, AMPM_UPPER) & ", Start Time = " & DateFormat(m.Bars.Prop(eBARS_StartTime) / 1440#, NO_DATE, HH_MM_SS, AMPM_UPPER) & " ) -- Not valid trading time"
        End If
        If (bBetweenSessions <> m.bBetweenSessions) Or (bLogged = False) Then
            DumpDebug vbTab & "Between ( Sessions ) changed to " & Str(bBetweenSessions) & " ( Current Exchange Time = " & DateFormat(dNow, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", End Time = " & DateFormat(m.Bars.Prop(eBARS_EndTime) / 1440#, NO_DATE, HH_MM_SS, AMPM_UPPER) & ", Start Time = " & DateFormat(m.Bars.Prop(eBARS_StartTime) / 1440#, NO_DATE, HH_MM_SS, AMPM_UPPER) & " ) -- Not valid trading time"
            
            If bBetweenSessions <> m.bBetweenSessions Then
                m.bBetweenSessions = bBetweenSessions
                
                If m.bBetweenSessions = True Then
                    ' DAJ 04/14/2015: I have seen a few trade logs where the closing position flag gets
                    ' stuck forever because the market order failed for some reason.  Tim and I agree that
                    ' we can at least clear it here when the session ends...
                    If ClosingPosition = True Then
                        DumpDebug vbTab & "Clearing closing position flag because we are now between sessions"
                        ClosingPosition = False
                    End If
                    
                    DumpDebug vbTab & "Cancelling submitted orders because we are now between sessions"
                    CancelSubmittedOrders False
                End If
            End If
        End If
        
        bLogged = True
    Else
        ' Find the on-close time for the current trading session...
        ' (TLB 2/8/2013: subtract 5 seconds -- just to be on the safe side, esp. in case exiting right as market is closing)
        dOnCloseTimeExch = Val(lSessionDate) + m.dOnCloseTimeExch - 5 / 86400#
        
        ' Find the start and end times for the current trading session ...
        dStartTimeExch = lSessionDate + (m.Bars.Prop(eBARS_StartTime) / 1440#)
        If m.Bars.Prop(eBARS_StartTime) > m.Bars.Prop(eBARS_EndTime) Then
            dStartTimeExch = dStartTimeExch - 1
        End If
        dEndTimeExch = lSessionDate + (m.Bars.Prop(eBARS_EndTime) / 1440#)
        
        dNextStart = dStartTimeExch
        Do While dNextStart < dNow
            dNextStart = dNextStart + 1
        Loop
        
        bBetweenOpen = ((dNow < dStartTimeExch) Or (dNow >= dOnCloseTimeExch))
        bBetweenClose = ((dNow < dEndTimeExch) And (dNow >= dOnCloseTimeExch))
        bBetweenSessions = ((dNow < dNextStart) And (dNow > dEndTimeExch))
        
        If (bBetweenOpen <> m.bBetweenOpen) Or (bLogged = False) Then
            m.bBetweenOpen = bBetweenOpen
            DumpDebug vbTab & "Between ( Open ) changed to " & Str(m.bBetweenOpen) & " ( Current Exchange Time = " & DateFormat(dNow, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", On Close Time = " & DateFormat(dOnCloseTimeExch, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", Start Time = " & DateFormat(dStartTimeExch, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", End Time = " & DateFormat(dEndTimeExch, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & " )"
        End If
                    
        If (bBetweenClose <> m.bBetweenClose) Or (bLogged = False) Then
            m.bBetweenClose = bBetweenClose
            DumpDebug vbTab & "Between ( Close ) changed to " & Str(m.bBetweenClose) & " ( Current Exchange Time = " & DateFormat(dNow, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", On Close Time = " & DateFormat(dOnCloseTimeExch, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", Start Time = " & DateFormat(dStartTimeExch, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", End Time = " & DateFormat(dEndTimeExch, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & " )"
        End If
    
        If (bBetweenSessions <> m.bBetweenSessions) Or (bLogged = False) Then
            DumpDebug vbTab & "Between ( Sessions ) changed to " & Str(bBetweenSessions) & " ( Current Exchange Time = " & DateFormat(dNow, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", End Time = " & DateFormat(m.Bars.Prop(eBARS_EndTime) / 1440#, NO_DATE, HH_MM_SS, AMPM_UPPER) & ", Start Time = " & DateFormat(m.Bars.Prop(eBARS_StartTime) / 1440#, NO_DATE, HH_MM_SS, AMPM_UPPER) & " ) -- Not valid trading time"
            
            If bBetweenSessions <> m.bBetweenSessions Then
                m.bBetweenSessions = bBetweenSessions
                
                If m.bBetweenSessions = True Then
                    ' DAJ 04/14/2015: I have seen a few trade logs where the closing position flag gets
                    ' stuck forever because the market order failed for some reason.  Tim and I agree that
                    ' we can at least clear it here when the session ends...
                    If ClosingPosition = True Then
                        DumpDebug vbTab & "Clearing closing position flag because we are now between sessions"
                        ClosingPosition = False
                    End If
                    
                    DumpDebug vbTab & "Cancelling submitted orders because we are now between sessions"
                    CancelSubmittedOrders False
                End If
            End If
        End If
    
        bLogged = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.CalcBetween"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcBetween
'' Description: Determine if we are between the 'On Close' time and the start
''              of the next session
'' Inputs:      On-Close Time (out)
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcBetweenOld(Optional dOnClose As Double)
On Error GoTo ErrSection:

    Dim dNow As Double                  ' Time right now
    Dim strTimeZone As String           ' Time zone information from the bars
    Dim dSessionStart As Double         ' Session start time
    Dim dSessionEnd As Double           ' Session end time
    Static dLastOnCloseTime As Double   ' Last on-close time

    strTimeZone = m.Bars.Prop(eBARS_ExchangeTimeZoneInf)
    dNow = CurrentTime(strTimeZone)
    
    dSessionStart = Val(Int(dNow)) + (m.Bars.Prop(eBARS_StartTime) / 1440#)
    dSessionEnd = Val(Int(dNow)) + (m.Bars.Prop(eBARS_EndTime) / 1440#)
    dOnClose = Val(Int(dSessionEnd)) + m.dOnCloseTimeExch
    If dSessionStart < dOnClose Then
        dSessionStart = dSessionStart + 1
        Do While Not IsWeekday(m.Bars.SessionDateForTradeTime(dSessionStart))
            dSessionStart = dSessionStart + 1
        Loop
    End If
        
    m.bBetweenOpen = ((dNow > dOnClose) And (dNow <= dSessionStart))
    
    If dOnClose <> dLastOnCloseTime Then
        DumpDebug vbTab & "OnClose = " & DateFormat(dOnClose, MM_DD_YYYY, HH_MM_SS) & "; SessionStart = " & DateFormat(dSessionStart, MM_DD_YYYY, HH_MM_SS) & "; Between = " & Str(m.bBetweenOpen)
        dLastOnCloseTime = dOnClose
    End If
    
    dOnClose = ConvertTimeZone(dOnClose, strTimeZone, "")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.CalcBetweenOld"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UseGtcOrder
'' Description: Determine if we want to submit orders as GTC
'' Inputs:      None
'' Returns:     True if GTC, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function UseGtcOrder() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strSymbol As String             ' Symbol to check
    
    strSymbol = GetSymbol(TradeSymbolOrID)
    
    If InStr(strSymbol, "@PFG") <> 0 Then
        bReturn = True
    ElseIf InStr(strSymbol, "@IB") <> 0 Then
        bReturn = True
    ElseIf InStr(strSymbol, "@CNX") <> 0 Then
        bReturn = True
    Else
        bReturn = False
    End If
    
    UseGtcOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.UseGtcOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateLogString
'' Description: Build a log string for the given date
'' Inputs:      Date
'' Returns:     Date String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DateLogString(ByVal dDate As Double) As String
On Error GoTo ErrSection:

    DateLogString = DateFormat(dDate, MM_DD_YY, HH_MM_SS) & " ( " & Format(dDate, "0.000000") & " )"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.DateLogString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetQuantityOnEngineString
'' Description: Set quantity on the engine string
'' Inputs:      Engine String
'' Returns:     New Engine String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetQuantityOnEngineString(ByVal strEngineString As String) As String
On Error GoTo ErrSection:

    Dim astrOrder As cGdArray           ' Array of order information from string
    Dim lQuantity As Long               ' Quantity for the engine order
    
    Set astrOrder = New cGdArray
    astrOrder.SplitFields strEngineString, vbTab
    
    ' For an entry, the quantity on the order should be the engine-generated quantity times the quantity of the next entry...
    If (UCase(astrOrder(OrderField(eGDOrderField_Position))) = "ES") Or (UCase(astrOrder(OrderField(eGDOrderField_Position))) = "EL") Then
        lQuantity = CLng(Val(astrOrder(OrderField(eGDOrderField_NumContracts)))) * m.lQtyNextEntry
        
    ' For an exit if the user is NOT pyramiding, the quantity on the order should be the current position so that the order
    ' exits the entire position...
    ElseIf Pyramid = False Then
        lQuantity = Abs(m.lCurrentPosition)
        
    ' For an exit if the user is pyramiding, we are going to exit the entire position for now, but this is not correct because
    ' pyramiding can exit out of a fraction of a position -- this needs to be thought out some more...
    Else
        lQuantity = Abs(m.lCurrentPosition)
    End If
    
    astrOrder(OrderField(eGDOrderField_NumContracts)) = Str(lQuantity)
    
    SetQuantityOnEngineString = astrOrder.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.SetQuantityOnEngineString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClosestOrderIsMarketOrder
'' Description: Determine if the order closest to the market is a market order
'' Inputs:      None
'' Returns:     True if Market Order, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ClosestOrderIsMarketOrder() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim Order As cPtOrder               ' Closest order
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 1 To m.SubmittedOrders.Count
        Set Order = m.SubmittedOrders(lIndex)
        If Order.OrderType = eTT_OrderType_Market Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    ClosestOrderIsMarketOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.ClosestOrderIsMarketOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BarNumberForFill
'' Description: Determine the bar number for the given fill
'' Inputs:      Fill
'' Returns:     Bar Number
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BarNumberForFill(ByVal Fill As cPtFill) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    If m.Bars.Prop(eBARS_PeriodType) >= ePRD_Days Then
        lReturn = m.Bars.FindDateTime(Fill.SessionDate)
    Else
        lReturn = m.Bars.FindDateTime(Fill.FillDateConverted(m.Bars.Prop(eBARS_ExchangeTimeZoneInf)))
    End If
    
    BarNumberForFill = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.BarNumberForFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateLimitPriceArrays
'' Description: Update the limit price arrays for the given bar number
'' Inputs:      Bar Number, Price, Buy?, Log if Changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateLimitPriceArrays(ByVal lBarNumber As Long, ByVal dPrice As Double, ByVal bBuy As Boolean, ByVal bLogIfChanged As Boolean)
On Error GoTo ErrSection:

    If bBuy = True Then
        If (dPrice < m.adLowestLimitBuy(lBarNumber)) Or (m.adLowestLimitBuy(lBarNumber) = kNullData) Then
            m.adLowestLimitBuy(lBarNumber) = dPrice
            
            If bLogIfChanged = True Then
                DumpDebug vbTab & vbTab & "Lowest Limit Buy ( " & DateFormat(m.Bars(eBARS_DateTime, lBarNumber), MM_DD_YYYY, HH_MM_SS) & " ) changed to " & Str(dPrice)
            End If
            
            UpdateLimitArrays lBarNumber, bLogIfChanged
        End If
    Else
        If dPrice > m.adHighestLimitSell(lBarNumber) Then
            m.adHighestLimitSell(lBarNumber) = dPrice
            
            If bLogIfChanged = True Then
                DumpDebug vbTab & vbTab & "Highest Limit Sell ( " & DateFormat(m.Bars(eBARS_DateTime, lBarNumber), MM_DD_YYYY, HH_MM_SS) & " ) changed to " & Str(dPrice)
            End If
            
            UpdateLimitArrays lBarNumber, bLogIfChanged
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.UpdateLimitPriceArrays"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateLimitArrays
'' Description: Update the limit arrays for the given bar number
'' Inputs:      Bar Number, Log if Changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateLimitArrays(ByVal lBarNumber As Long, ByVal bLogIfChanged As Boolean)
On Error GoTo ErrSection:

    Dim iValue As Integer               ' Value for the array

    iValue = Abs(Int(m.adLowestLimitBuy(lBarNumber) = m.Bars(eBARS_Low, lBarNumber)))
    If m.anLimitBuyAtLow(lBarNumber) <> iValue Then
        m.anLimitBuyAtLow(lBarNumber) = iValue
        If bLogIfChanged = True Then
            DumpDebug vbTab & vbTab & "Limit Buy At Low ( " & DateFormat(m.Bars(eBARS_DateTime, lBarNumber), MM_DD_YYYY, H_MM_SS) & " ) changed to " & Str(iValue) & " ( Low = " & Str(m.Bars(eBARS_Low, lBarNumber)) & "; Lowest Fill = " & Str(m.adLowestLimitBuy(lBarNumber)) & " )"
        End If
    End If
    
    iValue = Abs(Int(m.adHighestLimitSell(lBarNumber) = m.Bars(eBARS_High, lBarNumber)))
    If m.anLimitSellAtHigh(lBarNumber) <> iValue Then
        m.anLimitSellAtHigh(lBarNumber) = iValue
        If bLogIfChanged = True Then
            DumpDebug vbTab & vbTab & "Limit Sell At High ( " & DateFormat(m.Bars(eBARS_DateTime, lBarNumber), MM_DD_YYYY, H_MM_SS) & " ) changed to " & Str(iValue) & " ( High = " & Str(m.Bars(eBARS_High, lBarNumber)) & "; Highest Fill = " & Str(m.adHighestLimitSell(lBarNumber)) & " )"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.UpdateLimitArrays"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadLimitArrays
'' Description: Load the limit arrays
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadLimitArrays()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As cPtFill                 ' Fill object
    Dim vTradeSymbol As Variant         ' Trade Symbol or ID
    Dim lBarNumber As Long              ' Bar number for the fill date
    Dim lSessionDate As Long            ' Session date to check
    
    vTradeSymbol = TradeSymbolOrID
    If IsIntraday(GetPeriodicity(m.strActualBarPeriod)) Then
        lSessionDate = m.Bars.SessionDate(m.Bars.Size - 1) - 1
    ElseIf m.Bars.Size > 1 Then
        lSessionDate = m.Bars.SessionDate(m.Bars.Size - 2)
    End If
    
    m.anLimitBuyAtLow.Clear
    m.anLimitBuyAtLow.Size = m.Bars.Size
    m.anLimitSellAtHigh.Clear
    m.anLimitSellAtHigh.Size = m.Bars.Size
    m.adHighestLimitSell.Clear
    m.adHighestLimitSell.Size = m.Bars.Size
    m.adLowestLimitBuy.Clear
    m.adLowestLimitBuy.Size = m.Bars.Size
    
    Set rs = g.dbPaper.OpenRecordset("SELECT tblFills.*, tblOrders.OrderType, tblOrders.LimitPrice " & _
                    "FROM tblFills INNER JOIN tblOrders ON tblFills.OrderID=tblOrders.OrderID " & _
                    "WHERE [AutoTradeID]=" & Str(m.lAutoTradeItemID) & " AND [OrderType]=" & Str(eTT_OrderType_Limit) & " " & _
                    "ORDER BY [FillDate] DESC;", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!SessionDate > lSessionDate Then
            If rs!AccountID = m.lAccountID Then
                Set Fill = New cPtFill
                If Fill.Load(rs!FillID, rs) Then
                    If Fill.SymbolOrSymbolID = vTradeSymbol Then
                        lBarNumber = BarNumberForFill(Fill)
                        UpdateLimitPriceArrays lBarNumber, rs!LimitPrice, Fill.Buy, False
                    End If
                End If
            End If
        Else
            Exit Do
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.LoadLimitArrays"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LimitFlag
'' Description: Determine the limit flag to send to the engine
'' Inputs:      None
'' Returns:     Limit Flag
''
'' "" - Force market through limit price, "High" - Allow fill at high on last bar,
'' "Low" - Allow fill at low on last bar, "Both" - Don't force market through on last bar
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LimitFlag() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lBarNumber As Long              ' Bar number to check
    
    strReturn = ""
    lBarNumber = m.Bars.Size - 1
    
    If (m.anLimitBuyAtLow(lBarNumber) = 1) And (m.anLimitSellAtHigh(lBarNumber) = 1) Then
        strReturn = "Both"
    ElseIf m.anLimitBuyAtLow(lBarNumber) = 1 Then
        strReturn = "Low"
    ElseIf m.anLimitSellAtHigh(lBarNumber) = 1 Then
        strReturn = "High"
    End If
    
    LimitFlag = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.LimitFlag"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangePositionSymbol
'' Description: Change the position symbol to the one given
'' Inputs:      New Position Symbol, Cancel submitted orders?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangePositionSymbol(ByVal vNewPositionSymbolOrId As Variant, ByVal bCancelSubmittedOrders As Boolean)
On Error GoTo ErrSection:

    If vNewPositionSymbolOrId <> PositionSymbolOrID Then
        ' Cancel all submitted orders for the old symbol first...
        If (m.SubmittedOrders.Count > 0) And (bCancelSubmittedOrders = True) Then
            DumpDebug vbTab & "Cancelling submitted orders because symbol has changed"
            CancelSubmittedOrders False
        End If
        
        DumpDebug vbTab & "Changing position symbol from '" & GetSymbol(PositionSymbolOrID) & "' to '" & GetSymbol(vNewPositionSymbolOrId) & "'"
        PositionSymbolOrID = vNewPositionSymbolOrId
        m.lAccountPositionID = g.Broker.CreateFillSummaryForAutoTrade(Me)
        m.lCurrentPosition = GetCurrentPositionFromBrokerInfo
        DumpDebug vbTab & "Current Position changed to " & Str(m.lCurrentPosition)
                
        Save
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.ChangePositionSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixEngineOrder
'' Description: Fix a stop with limit order coming from the engine
'' Inputs:      Engine order string
'' Returns:     Fixed order string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FixEngineOrder(ByVal strEngineOrder As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrOrder As cGdArray           ' Order string split out into an array
    Dim Order As cPtOrder               ' Order object built from the engine string
    Dim strMessage As String            ' Message to dump to the log
    
    strReturn = strEngineOrder
    Set astrOrder = New cGdArray
    astrOrder.SplitFields strEngineOrder, vbTab
    Set Order = EngineStringToOrder(strEngineOrder)
    
    If (Order.OrderType = eTT_OrderType_StopWithLimit) Or (Order.OrderType = eTT_OrderType_StopWithLimitCloseOnly) Then
        If Order.WrongSideOfMarket = True Then
            astrOrder(OrderField(eGDOrderField_OrderType)) = Replace(astrOrder(OrderField(eGDOrderField_OrderType)), "SW", "")
            astrOrder(OrderField(eGDOrderField_Price1)) = astrOrder(OrderField(eGDOrderField_Price2))
            astrOrder(OrderField(eGDOrderField_Price2)) = ""
                        
            strReturn = astrOrder.JoinFields(vbTab)
            
            strMessage = "Engine Order changed from '" & strEngineOrder & "' to '" & strReturn & "'"
            DumpDebug vbTab & vbTab & strMessage
        End If
    
    ElseIf Order.OrderType = eTT_OrderType_Market Then
        If NeedToChangeMarketToLimit(PositionSymbolOrID, m.Bars) = True Then
            ChangeMarketToLimit Order
            astrOrder(OrderField(eGDOrderField_OrderType)) = "L"
            astrOrder(OrderField(eGDOrderField_Price1)) = Str(Order.LimitPrice)
            astrOrder(OrderField(eGDOrderField_Price2)) = ""
                        
            strReturn = astrOrder.JoinFields(vbTab)
            
            strMessage = "Engine Order changed from '" & strEngineOrder & "' to '" & strReturn & "'"
            DumpDebug vbTab & vbTab & strMessage
        End If
    End If
    
    FixEngineOrder = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.FixEngineOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeMarketToLimit
'' Description: Change a market order to a limit order on the "wrong" side of the market
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeMarketToLimit(Order As cPtOrder)
On Error GoTo ErrSection:

    Dim dLimitPrice As Double           ' Limit price for the order

    If Order.OrderType = eTT_OrderType_Market Then
        If NeedToChangeMarketToLimit(Order.SymbolOrSymbolID, m.Bars) = True Then
            Order.OrderType = eTT_OrderType_Limit
            Order.LimitPrice = mTradeTracker.LimitPriceForMarketOrder(Order.Buy, Order.Symbol, m.Bars)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.ChangeMarketToLimit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixOnCloseTime
'' Description: If the on-close time is an invalid trading time, fix it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixOnCloseTime()
On Error GoTo ErrSection:

    Dim dOnCloseTime As Double          ' On-Close time
    Dim dSessionStart As Double         ' Session start time
    Dim dSessionEnd As Double           ' Session end time
    Dim bFix As Boolean                 ' Does the on-close time need to be fixed?
    Dim strMessage As String            ' Message to show to the user
    
    dOnCloseTime = m.dOnCloseTimeExch
    dSessionStart = m.Bars.Prop(eBARS_DefaultStartTime) / 1440#
    dSessionEnd = m.Bars.Prop(eBARS_DefaultEndTime) / 1440#
    
    If dSessionStart > dSessionEnd Then
        bFix = ((dOnCloseTime > dSessionEnd) And (dOnCloseTime < dSessionStart))
    Else
        bFix = ((dOnCloseTime > dSessionEnd) Or (dOnCloseTime < dSessionStart))
    End If
    
    If bFix = True Then
        m.dOnCloseTimeExch = dSessionEnd - (m.lMinutesBefore / 1440#)
        Save
        
        strMessage = "On-Close time for '" & m.strName & "' has been changed from " & DateFormat(dOnCloseTime, NO_DATE, H_MM, AMPM_UPPER) & " to " & DateFormat(m.dOnCloseTimeExch, NO_DATE, H_MM, AMPM_UPPER) & " because the previous value was outside of market hours"
        mMain.ShowAlertPopup strMessage, "Automated Trading Message"
        DumpDebug strMessage
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAutoTradeItem.FixOnCloseTime"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllowPyramiding
'' Description: Determine if pyramiding is allowed for automated trading for this user
'' Inputs:      None
'' Returns:     True if allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AllowPyramiding() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    
    Select Case mTradeTracker.TypeOfAccount(m.lAccountID)
        Case eGDTypeOfAccount_Simulated
            bReturn = HasModule("AUTOPRMD")
        Case eGDTypeOfAccount_BrokerDemo
            bReturn = HasModule("AUTOPRMD")
        Case eGDTypeOfAccount_BrokerLive
            bReturn = HasModule("AUTOPRMDL")
    End Select
    
    AllowPyramiding = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.AllowPyramiding"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoOnCloseCheck
'' Description: Determine whether or not we want to do an on-close check
'' Inputs:      None
'' Returns:     True if we want to do an on-close check, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOnCloseCheck() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lNextTradingDay As Long         ' Next trading day after current session
    Dim lLastBarDate As Long            ' Date of the last bar
    
    lNextTradingDay = m.Bars.NextTradingDay(SessionDate)
    lLastBarDate = Int(m.Bars(eBARS_DateTime, m.Bars.Size - 1))
    bReturn = (m.Bars.Prop(eBARS_Periodicity) <= ePRD_Days + 1) Or (lNextTradingDay > lLastBarDate)
    
    DumpDebug vbTab & vbTab & "DoOnCloseCheck = " & Str(bReturn) & ": Period = '" & m.Bars.Prop(eBARS_PeriodicityStr) & "'; LastBarDate = " & DateFormat(lLastBarDate, MM_DD_YYYY) & "; Next Trading Day = " & DateFormat(lNextTradingDay, MM_DD_YYYY)
    
    DoOnCloseCheck = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItem.DoOnCloseCheck"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Setup and initialize class members
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    m.bConfirm = True
    BarPeriod = "Daily"
    m.lConfirmTimeout = 10
    m.iRolling = 0
    m.dOnCloseTimeExch = 0#
    m.bClosingPosition = False
    m.lClosingQuantity = 0&
    m.bDoSubmitOneOrder = False
    m.strCurrentBarCheckReason = ""
    m.dNextSanityCheck = 99999#
    
    Set m.Strategy = New cSystem
    Set m.Bars = New cGdBars
    Set m.adLowestLimitBuy = New cGdArray
    m.adLowestLimitBuy.Create eGDARRAY_Doubles, , kNullData
    Set m.adHighestLimitSell = New cGdArray
    m.adHighestLimitSell.Create eGDARRAY_Doubles, , kNullData
    Set m.astrOrders = New cGdArray
    Set m.anLimitBuyAtLow = New cGdArray
    m.anLimitBuyAtLow.Create eGDARRAY_TinyInts, , 0
    Set m.anLimitSellAtHigh = New cGdArray
    m.anLimitSellAtHigh.Create eGDARRAY_TinyInts, , 0
    m.astrOrders.Create eGDARRAY_Strings
    Set m.SubmittedOrders = New cGdTree
    Set m.astrSecondaryMarkets = New cGdArray
    m.astrSecondaryMarkets.Create eGDARRAY_Strings
    Set m.SecondaryMarkets = New cGdTree
    
    If Not DirExist(AddSlash(App.Path) & "AutoTrade") Then MkDir AddSlash(App.Path) & "AutoTrade"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up upon destruction
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Strategy = Nothing
    Set m.Bars = Nothing
    Set m.adLowestLimitBuy = Nothing
    Set m.adHighestLimitSell = Nothing
    Set m.anLimitBuyAtLow = Nothing
    Set m.anLimitSellAtHigh = Nothing
    Set m.astrOrders = Nothing
    Set m.SubmittedOrders = Nothing
    Set m.astrSecondaryMarkets = Nothing
    Set m.SecondaryMarkets = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItem.Class_Terminate"
    
End Sub
