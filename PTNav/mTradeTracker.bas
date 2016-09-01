Attribute VB_Name = "mTradeTracker"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mTradeTracker.bas
'' Description: Global routines for trading
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/26/2009   DAJ         Allow for cancelling a market order if Parked
'' 01/26/2009   DAJ         Changed some of the Interactive Brokers messages
'' 02/03/2009   DAJ         Only allow a flatten for a symbol that is not currently
''                          in a position mismatch
'' 02/06/2009   DAJ         Display "Mismatch" for position if in a mismatch
'' 02/23/2009   DAJ         Prepend WeekNum to PFG orders instead of calendar date
'' 05/06/2009   DAJ         Refresh edit order form in RefreshOrder call
'' 05/18/2009   DAJ         Added the IsManExpressDemoAccount function
'' 05/19/2009   DAJ         Fixed PostFixVersion28 to handle database changes
'' 05/19/2009   DAJ         Fixed IsManExpressDemoAccount to take number or ID
'' 05/26/2009   DAJ         Set StatusDate on Xpress orders when Amend Pending
'' 06/01/2009   DAJ         Added FO and SO to security type mask for trading
'' 06/05/2009   DAJ         Added TypeOfAccount stuff
'' 07/28/2009   DAJ         Allow Cancel of TriggerPending Market order on live
'' 08/21/2009   DAJ         Make chanes for Parking/Submitting/Cancelling OCOs
'' 09/01/2009   DAJ         Use new Parked order status, Fix OCO stuff
'' 09/02/2009   DAJ         Added the bAskUserAboutOtherside arg to SubmitOrder
'' 10/07/2009   DAJ         Added support for Linked Orders held at broker, reworked amend
'' 10/07/2009   DAJ         Implement price shaving for filling option orders
'' 10/07/2009   DAJ         Broker Held Linked Orders for simple auto exits
'' 10/08/2009   DAJ         Fix for advanced order trailing stop
'' 11/06/2009   DAJ         Fix for amending Xpress order by clicking on same price
'' 12/01/2009   DAJ         Enhancements for Auto Exits held at broker
'' 12/16/2009   DAJ         Fixes for Trigger Pending orders and parking a link
'' 12/21/2009   DAJ         Fix for submitting a parked broker OCO link
'' 12/21/2009   DAJ         Give user the option to Park simulated order on shutdown
'' 12/22/2009   DAJ         Fix for conditional AND trigger by orders
'' 12/29/2009   DAJ         Reject an order of quantity zero before submitting
'' 01/04/2010   DAJ         Ensure that HandleDemoOrders sleeps even if unloading
'' 01/08/2010   DAJ         Clean up triggered orders where triggering is closed
'' 02/09/2010   DAJ         Log when confirming a cancel
'' 03/02/2010   DAJ         Use specific list of TransAct simulated accounts
'' 03/03/2010   DAJ         Make sure TransAct object exists when checking sim accounts
'' 03/11/2010   DAJ         Use global collections, added global routines
'' 03/11/2010   DAJ         Don't combine orders @ same price with different expirations
'' 03/16/2010   DAJ         Create sim account if no non-stream replay sim accounts
'' 05/06/2010   DAJ         Handle parking/submitting order with triggered by orders (#5715),
''                          Handle moving order with triggered by to a price with another order
'' 05/11/2010   DAJ         Fix merging when already one non-triggered and one triggered
'' 05/17/2010   DAJ         Added enumerations for TradeSense orders
'' 05/20/2010   DAJ         Fixed issue with amending an order over another order
'' 05/20/2010   DAJ         Fixed check to see if TIF was the same
'' 05/25/2010   DAJ         Fixes and enhancements to conditional orders
'' 05/25/2010   DAJ         Added routine to delete blank accounts out of the database
'' 06/03/2010   DAJ         Changes for new Trade Sense Order Groups
'' 06/15/2010   DAJ         Don't allow flatten until data available (#5768)
'' 06/25/2010   DAJ         Fix for orders triggered by a conditional order (#5820)
'' 06/25/2010   DAJ         Fix for OrderIsEntry when multiple levels of trigger (#5823)
'' 06/30/2010   DAJ         Fix for using TradeSense order groups through PFG
'' 07/06/2010   DAJ         Fix for flattening position before streaming connected (#5768)
'' 07/15/2010   DAJ         Added ShowAdvancedTSOG function
'' 07/20/2010   DAJ         Added in the Daniel Code stuff
'' 07/21/2010   DAJ         Changed over to the Provided.INI for Daniel Code stuff
'' 07/21/2010   DAJ         Fix for Broker OCO when one side is conditional (#5840)
'' 07/21/2010   DAJ         Added message timeout to HandleDemoOrders
'' 08/04/2010   DAJ         Added flag file for DanielCode/TradeSense Orders/Groups
'' 08/05/2010   DAJ         Disable DanielCode button when stand alone is running
'' 08/10/2010   DAJ         When starting Daniel Code button, kill file if exist
'' 08/12/2010   DAJ         New method for checking streaming availability for a symbol
'' 08/23/2010   DAJ         New method for deactivating trading related items
'' 09/13/2010   DAJ         Added code for Rithmic
'' 09/17/2010   DAJ         More information in alert for order status change (#5891)
'' 09/20/2010   DAJ         Fix for auto exits/TradeSense order groups when stream stops
'' 09/22/2010   DAJ         Don't show deactivate exits message if all broker held (#5929)
'' 09/28/2010   DAJ         Take triggered by price into account in CanMoveOrder (#5947)
'' 09/29/2010   DAJ         Changed global order confirmation flags
'' 09/30/2010   DAJ         Send arguments file over to the DanielCode process
'' 10/01/2010   DAJ         Fixed the argument call to the DanielCode process
'' 10/05/2010   DAJ         Additional logging in CancelOrder
'' 10/08/2010   DAJ         Don't merge TradeSense orders with manual orders
'' 10/28/2010   DAJ         Added PFG spread calls
'' 11/01/2010   DAJ         Added Optimus, OpVest, and Vision (Rithmic Brokers)
'' 11/04/2010   DAJ         Don't show position verification for SimTrade if drop/reconnect to stream
'' 11/18/2010   DAJ         Send new parked orders over to OptNav if applicable
'' 11/29/2010   DAJ         If unsent order gets cancelled, send to Opt Nav if applicable
'' 12/01/2010   DAJ         Require Gold for TradeSense auto exits/order groups instead of flag file
'' 12/10/2010   DAJ         Added Zen-Fire, Changed over to the IsBrokerUser function
'' 01/04/2011   DAJ         Don't Set OTO Negative if OrderID is zero
'' 01/12/2011   DAJ         Reload order after Cancel confirm dialog up in case order is now closed
'' 01/28/2011   DAJ         Utilized new Order.ChangeOrderStatus call
'' 03/02/2011   DAJ         If conditional order amended, don't mark as AmendPending
'' 03/07/2011   DAJ         Added Oec/OptionsExpress, Added Change Password for LindXpress,
''                          IB/I-Deal/Rithmic/Gain now utilize cBroker, Added unsolicited fill/position
''                          for IB/Ideal
'' 03/17/2011   DAJ         Added CreateMarketOrder, Pass Daniel Code enablement
'' 03/18/2011   DAJ         Utilize the CreateMarketOrder function
'' 04/20/2011   DAJ         Added PreSubmitted order status
'' 04/29/2011   DAJ         Added Inactive order status
'' 05/02/2011   DAJ         Set the order date for new park order and create market order
'' 05/16/2011   DAJ         Set OrderDate correctly in case delayed streaming
'' 06/15/2011   DAJ         When cancelling a trigger order, make sure it is still open
'' 06/21/2011   DAJ         Separate out Simulated trading types
'' 06/22/2011   DAJ         Don't wait for data on a flatten unless it is SimStream
'' 06/24/2011   DAJ         Moved NextSimTradeAccount functionality to simulated objects
'' 07/11/2011   DAJ         Fix for cancelling OTO conditional orders when parent order is closed
'' 07/19/2011   DAJ         Added PositionToString function
'' 07/28/2011   DAJ         Default new simulated account to SimBroker
'' 08/02/2011   DAJ         When parking sim orders at stream shutdown, don't do auto trade orders
'' 08/25/2011   DAJ         Added Suspended order status, mods for CQG/TT
'' 08/25/2011   DAJ         Updated the broker message enumeration string functions
'' 09/07/2011   DAJ         Save original order on an amend with CQG
'' 09/08/2011   DAJ         Call deactivate on trading items upon stream loss
'' 09/20/2011   DAJ         Added price messages for brokers, some amend fixes for TT
'' 09/21/2011   DAJ         If user has RTG, setup default sim account as sim stream
'' 09/28/2011   DAJ         Fixes for automated journaling for automated orders
'' 10/04/2011   DAJ         Added the ShowJournals function
'' 10/14/2011   DAJ         Don't do auto journal on cancel unless user initiated the cancel
'' 10/17/2011   DAJ         Added the auto breakout for TradeSense order groups function
'' 10/24/2011   DAJ         Tweaked the CreateAccountFromNumber routine
'' 11/02/2011   DAJ         Added Amp Trading and RJ O'Brien as CQG brokers
'' 12/02/2101   DAJ         Added RJO (PATS)
'' 12/09/2011   DAJ         Perform a post-fix for database version 57, Added GFT Forex & OptionsHouse
'' 12/13/2011   DAJ         Added Capital Trading Group for PATS and CQG
'' 12/14/2011   DAJ         Added Capital Trading Group and Fintec for PFG
'' 12/19/2011   DAJ         Added some logging into OrderIsEntry
'' 01/18/2012   DAJ         Enhanced logging for automated trading
'' 01/30/2012   DAJ         Option Nav Journal Image
'' 01/31/2012   DAJ         Add option to OrderIsEntry whether to dump to log
'' 02/14/2012   DAJ         New status alerts for position mismatch / auto trade disabled
'' 02/14/2012   DAJ         Added multi-leg order support
'' 02/28/2012   DAJ         Put in user warning for trying to live trade against delayed data
'' 03/06/2012   DAJ         Mods to the CanMoveOrder routine
'' 03/14/2012   DAJ         Added Alpari(Currenex), Alpari(PATS), Penson(Currenex), Penson(CQG)
'' 03/19/2012   DAJ         Added Symbol/Symbol ID to the date journals table
'' 04/05/2012   DAJ         Added new broker message enums, broker mode stuff
'' 04/12/2012   DAJ         Optionally take broker on the NextGenesisOrderID call
'' 05/31/2012   DAJ         Turnkey implementation
'' 06/04/2012   DAJ         Turnkey administration / Split from broker stand-alone
'' 06/05/2012   DAJ         Changed CreateOrder to take the Lot number
'' 06/11/2012   DAJ         Make Turnkey work with all brokers
'' 06/25/2012   DAJ         Visible Columns mode for Turnkey Admin
'' 07/13/2012   DAJ         Added in new calls for getting history from IB
'' 07/16/2012   DAJ         ZanerCqg, ZanerPats, ZanerRithmic, ZanerZenFire, KnightCnx, KnightCqg
'' 07/17/2012   DAJ         AlpariZenFire
'' 07/17/2012   DAJ         RobbinsCqg
'' 07/18/2012   DAJ         RCG (New PATS)
'' 07/27/2012   DAJ         Demo (PATS), New Currenex Messages, GmajPro
'' 08/02/2012   DAJ         If GmajPro or DanielCode are running, disable both toolbar buttons
'' 08/03/2012   DAJ         Remove Gain, FXCM, Photon, OptionsHouse, Alaron, Cadent, Lotus, OptXpress, Oec, ManChicago, ManLondon, Robbins
'' 08/17/2012   DAJ         Added more brokers to the SubmitAmend routine
'' 08/23/2012   DAJ         Born (PATS), RJO Hong Kong (PATS), Get symbols from PATS server
'' 08/29/2012   DAJ         Zaner (Currenex)
'' 08/31/2012   DAJ         Load different INI file properties for GmajPro
'' 09/06/2012   DAJ         Changed "Cancel market order" error to use infbox off of frmOnlineBroker
'' 09/11/2012   DAJ         New Turnkey Enums, Key on Lot ID for CreateOrder
'' 09/12/2012   DAJ         Added Currenex, FXDD (Currenex), and VanKar (Currenex)
'' 09/12/2012   DAJ         Removed Rosenthal (Old PATS), Changed Generic PATS to New PATS
'' 09/20/2012   DAJ         Changed DeleteTrades to DeleteTrade (only delete one trade)
'' 10/03/2012   DAJ         Fixed broker enumerations that were out of order and duplicated
'' 10/05/2012   DAJ         Added Account threshold for CQG brokers
'' 10/23/2012   DAJ         New enablements for DCPro Futures / Forex
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 12/11/2012   DAJ         Use the flatten queue for position reversals
'' 12/11/2012   DAJ         Contingency Orders
'' 12/11/2012   DAJ         Vision (CQG)
'' 01/18/2013   DAJ         Don't allow automated trading for spreads
'' 01/18/2013   DAJ         Broker held OCO for Interactive Brokers
'' 01/18/2013   DAJ         Show message for Daniel Code buttons if less than Gold
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 02/12/2013   DAJ         When submitting prices, don't check for another order with null order price
'' 02/20/2013   DAJ         Changed trade report filter, added filter enumerations
'' 03/11/2013   DAJ         "Reject Message" from CQG
'' 03/19/2013   DAJ         When cancelling all SimTrade orders, don't always confirm each cancel
'' 04/01/2013   DAJ         Changed SubmitAmend function to work for parked orders with CQG
'' 04/02/2013   DAJ         Allow 'Trigger Pending' orders to be parked in 'ParkOrderFromOrder' ( #6808 )
'' 04/03/2013   DAJ         Log when adding order to CQG amend orders array
'' 04/15/2013   DAJ         Check for broker connection before attempting a cancel, change timeout
''                          on confirm order InfBoxes to 20 seconds
'' 04/15/2013   DAJ         Only allow auto exits, TSOG, and automated trading if enabled for streaming
'' 05/01/2013   DAJ         Shadow Trading
'' 05/06/2013   DAJ         Message message from Interactive Brokers stand-alone
'' 05/08/2013   DAJ         Added CreateStopOrder, CreateLimitOrder, Allow Merge flag on SubmitOrder/SubmitAmend
'' 05/13/2013   DAJ         Allow a user to park a market order if it is a brand new order
'' 06/07/2013   DAJ         Allow date journals with an enablement OR Gold and above
'' 06/12/2013   DAJ         Symbol and quantity validation for automated trading items
'' 07/16/2013   DAJ         Trade Routes for Rithmic
'' 07/30/2013   DAJ         Automatic Journal for a Fill, Data Pending order status for a conditional order
'' 08/01/2013   DAJ         Change to whether or not order is considered conditional
'' 08/01/2013   DAJ         Fix for expiring non-submitted orders
'' 08/08/2013   DAJ         Journal category types
'' 09/04/2013   DAJ         Extra logging in the RefreshAccountCombos function
'' 10/04/2013   DAJ         Fixes for OCO when the non-filling order is pending
'' 10/10/2013   DAJ         Moved conditional order add from cWorkingOrdersUI
'' 10/16/2013   DAJ         Removed PFG/Xpress/OrderLinks, Added Oec/FptOec/FptCqg
'' 11/04/2013   DAJ         Show delayed data warning if they try to trade on SimBroker
'' 11/15/2013   DAJ         Changed/Added Turnkey enumerations
'' 12/03/2013   DAJ         Expand/Collapse Level; Default visible lot columns; New turnkey customer types
'' 12/04/2013   DAJ         Detail Options
'' 12/18/2013   DAJ         Added OEC to Automated Forex check
'' 01/03/2014   DAJ         "Either Feedyard" mode
'' 01/09/2014   DAJ         Fixed bug with FuturePath CQG and FuturePath OEC having same enum
'' 02/25/2014   DAJ         Rations/Ingredients
'' 03/03/2014   DAJ         Got rid of DajLog function; Commented out WaitListCommand stuff
'' 03/07/2014   DAJ         Moved Cattle stuff into NavCattle.DLL; Moved some broker stuff
''                          into NavBroker.DLL; Moved some flex grid stuff into mFlexGrid
'' 04/24/2014   DAJ         Confirm flatten order when account associated to lot
'' 05/01/2014   DAJ         Queue up fills going to automated trading items
'' 08/08/2014   DAJ         Fix initial simulated account if wrong type
'' 08/11/2014   DAJ         New flag for how to calculate open equity on options
'' 08/13/2014   DAJ         Interactive Brokers Symbol Availability
'' 08/22/2014   DAJ         Added E-Trade
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 10/09/2014   DAJ         Clear GenesisOrderID for IB orders when move to history
'' 10/24/2014   DAJ         Core Application functions for DLL's; Moved TypeOfAccount and FillMatchMode
''                          enums to NavBroker.DLL; Commented out more PFG/Xpress stuff; Fill Display
'' 10/28/2014   DAJ         Pass order into the automated trading item for a FillCheck
'' 10/29/2014   DAJ         Remove old synthetic order/MIT code
'' 10/31/2014   DAJ         Consider 'Held' an open order status
'' 11/03/2014   DAJ         Added the SetDllBridgeDatabases function
'' 11/21/2014   DAJ         Modified TypeOfAccount function to determine CQG/IB demo accounts
'' 11/28/2014   DAJ         Don't handle Trigger Order in OrderCallback if it can't be loaded
'' 12/02/2014   DAJ         Load the trigger order with the absolute value of ID in OrderCallback
'' 12/04/2014   DAJ         Remove enabled symbol check for trading; Removed old code
'' 01/29/2015   DAJ         Added ContractInfo calls for Rithmic
'' 04/10/2015   DAJ         Fix for parked trailing stop getting automatically submitted
'' 04/28/2015   DAJ         Allow user to cancel a market order on live account if 'Suspended' or 'PreSubmitted'
'' 05/20/2015   DAJ         Allow multiple accounts for the trade report filter
'' 07/15/2015   DAJ         Added LimitPriceForMarketOrder and NeedToChangeMarketToLimit functions
'' 07/16/2015   DAJ         Allow automated trading on spreads with enablement
'' 07/16/2015   DAJ         Fixed the NeedToChangeMarketToLimit function
'' 07/17/2015   DAJ         Have the NeedToChangeMarketToLimit function take a Symbol or Symbol ID instead of symbol
'' 07/20/2015   DAJ         Tweaked the NeedToChangeMarketToLimit function
'' 08/06/2015   DAJ         Default the account for new order to the default account if order bar not on active chart
'' 09/14/2015   DAJ         Added Tradier
'' 10/06/2015   DAJ         Don't include Tradier in auto broker connect when starting streaming
'' 11/06/2015   DAJ         Utilize new move with trigger flag on order; Tell conditional orders when order ID changes
'' 11/23/2015   DAJ         Fix in CreateOrder when active chart is detached
'' 12/09/2015   DAJ         Fix the duplicate Group ID's in the orders table
'' 01/15/2016   DAJ         Added optional ShowMessage argument to CanActivateAutomatedItem
'' 03/18/2016   DAJ         Added TD Ameritrade
'' 03/31/2016   DAJ         Fix for the duplicate Group ID fix
'' 04/12/2016   DAJ         Added Market type to DebitCredit enumeration
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Global Const kgntGetInfoQueueEmpty = &H301000D
Global Const kTransActSimUserAccount = "3803"
Global Const kTransActOldSimUserAccount = "20394"
Global Const kTransActRegistryKey = "Software\YesTrader\Trader"

Global Const kEntryTradeRuleID = 1000000
Global Const kExitTradeRuleID = 2000000
Global Const kStartingCustomTradeRuleID = 100000

Global Const kBuyColor = &HFFA0A0
Global Const kSellColor = &HA0A0FF

'Public Enum eTT_OrderType
'    eTT_OrderType_Market = 0
'    eTT_OrderType_Stop = 1
'    eTT_OrderType_Limit = 2
'    eTT_OrderType_StopWithLimit = 3
'
'    eTT_OrderType_MarketOnClose = 4
'    eTT_OrderType_StopCloseOnly = 5
'    eTT_OrderType_LimitCloseOnly = 6
'    eTT_OrderType_StopWithLimitCloseOnly = 7
'
'    eTT_OrderType_MIT = 8
'
'    eTT_OrderType_Adjustment = 10
'End Enum

'Public Enum eTT_OrderStatus
'    eTT_OrderStatus_Open = 0
'    eTT_OrderStatus_Partial = 1
'    eTT_OrderStatus_Filled = 2
'    eTT_OrderStatus_Cancelled = 3
'    eTT_OrderStatus_Queued = 4
'    eTT_OrderStatus_Sent = 5
'    eTT_OrderStatus_Working = 6
'    eTT_OrderStatus_Rejected = 7
'    eTT_OrderStatus_BalCancelled = 8
'    eTT_OrderStatus_CancelPending = 9
'    eTT_OrderStatus_AmendPending = 10
'    eTT_OrderStatus_UnconfirmedFilled = 11
'    eTT_OrderStatus_UnconfirmedPartial = 12
'    eTT_OrderStatus_Held = 13
'    eTT_OrderStatus_CancelHeld = 14
'    eTT_OrderStatus_Error = 15
'    eTT_OrderStatus_Amended = 16
'    eTT_OrderStatus_Expired = 17
'    eTT_OrderStatus_Frozen = 18
'    eTT_OrderStatus_ParkPending = 19
'    eTT_OrderStatus_TriggerPending = 20
'    eTT_OrderStatus_Approved = 21
'    eTT_OrderStatus_BrokerParked = 22
'    eTT_OrderStatus_OverFilled = 23
'    eTT_OrderStatus_Parked = 24
'    eTT_OrderStatus_PreSubmitted = 25
'    eTT_OrderStatus_Inactive = 26
'    eTT_OrderStatus_Suspended = 27
'    eTT_OrderStatus_DataPending = 28
'End Enum

'Public Enum eTT_TimeInForce
'    eTT_TimeInForce_Day = 0
'    eTT_TimeInForce_GTC
'    eTT_TimeInForce_GTD
'End Enum

'Public Enum eTT_FillMatchMode
'    eTT_FillMatchMode_Fifo = 0
'    eTT_FillMatchMode_Lifo
'End Enum

'Public Enum eTT_AccountType
'    eTT_AccountType_Standard = 0
'    eTT_AccountType_PATS = 1
'    eTT_AccountType_SimTrade = 2
'    'eTT_AccountType_Photon = 3
'    eTT_AccountType_IntBrokers = 4
'    'eTT_AccountType_LindWaldock = 5
'    'eTT_AccountType_ManLondon = 6
'    'eTT_AccountType_ManChicago = 7
'    'eTT_AccountType_Alaron = 8
'    eTT_AccountType_TransAct = 9
'    'eTT_AccountType_PFG = 10
'    'eTT_AccountType_FXCM = 11
'    eTT_AccountType_TT = 12
'    eTT_AccountType_AdvFut = 13
'    'eTT_AccountType_Gain = 14
'    'eTT_AccountType_ManExpress = 15
'    'eTT_AccountType_Rosenthal = 16
'    'eTT_AccountType_Robbins = 17
'    eTT_AccountType_Ideal = 18
'    'eTT_AccountType_Cadent = 19
'    'eTT_AccountType_Lotus = 20
'    eTT_AccountType_Rithmic = 21
'    eTT_AccountType_Vision = 22
'    eTT_AccountType_Optimus = 23
'    eTT_AccountType_OpVest = 24
'    eTT_AccountType_ZenFire = 25
'    eTT_AccountType_Oec = 26
'    'eTT_AccountType_OptionsXpress = 27
'    eTT_AccountType_SimBroker = 28
'    eTT_AccountType_SimStream = 29
'    eTT_AccountType_SimReplay = 30
'    eTT_AccountType_CQG = 31
'    eTT_AccountType_AmpCqg = 32
'    eTT_AccountType_RjoCqg = 33
'    eTT_AccountType_RjoPats = 34
'    'eTT_AccountType_OptionsHouse = 35
'    eTT_AccountType_Gft = 36
'    eTT_AccountType_CtgCqg = 37
'    eTT_AccountType_CtgPats = 38
'    'eTT_AccountType_CtgPfg = 39
'    'eTT_AccountType_FintecPfg = 40
'    eTT_AccountType_AlpariCurrenex = 41
'    eTT_AccountType_AlpariPats = 42
'    eTT_AccountType_KnightCurrenex = 43
'    eTT_AccountType_KnightCqg = 44
'    eTT_AccountType_ZanerCqg = 45
'    eTT_AccountType_ZanerPats = 46
'    eTT_AccountType_ZanerRithmic = 47
'    eTT_AccountType_ZanerZenFire = 48
'    eTT_AccountType_AlpariZenFire = 49
'    eTT_AccountType_RobbinsCqg = 50
'    eTT_AccountType_RcgPats = 51
'    eTT_AccountType_DemoPats = 52
'    eTT_AccountType_RjoHkPats = 53
'    eTT_AccountType_BornPats = 54
'    eTT_AccountType_ZanerCurrenex = 55
'    eTT_AccountType_Currenex = 56
'    eTT_AccountType_FxddCurrenex = 57
'    eTT_AccountType_VanKarCurrenex = 58
'    eTT_AccountType_VisionCqg = 59
'    eTT_AccountType_FptOec = 60
'    eTT_AccountType_FptCqg = 61
'    eTT_AccountType_Etrade = 62
'    eTT_AccountType_Tradier = 63
'    eTT_AccountType_Ameritrade = 64
'End Enum
Public Const kNumBrokers = eTT_AccountType_Ameritrade + 1

Public Enum eGDEditOrderReturn
    eGDEditOrderReturn_Cancel = 0
    eGDEditOrderReturn_Submit
    eGDEditOrderReturn_Park
End Enum

Public Enum eGDEditOrderMode
    eGDEditOrderMode_FromPosition
    eGDEditOrderMode_OrderOnly
    eGDEditOrderMode_QuickOrder
End Enum

Public Enum eGDSimTradeMessageTypes
    eGDSimTradeMessageType_Account = 100
    eGDSimTradeMessageType_Order
    eGDSimTradeMessageType_Fill
    eGDSimTradeMessageType_Position
    eGDSimTradeMessageType_RefreshOrder
    eGDSimTradeMessageType_RefreshFill
    eGDSimTradeMessageType_RefreshPosition
    eGDSimTradeMessageType_SpreadFill
End Enum

Public Enum eGDTransActMessageTypes
    eGDTransActMessageType_Connect = 1
    eGDTransActMessageType_Disconnect
    eGDTransActMessageType_AddOrder
    eGDTransActMessageType_AmendOrder
    eGDTransActMessageType_CancelOrder
    eGDTransactMessageType_UnloadApp
    eGDTransActMessageType_GetOrders
    eGDTransActMessageType_GetTrades
    eGDTransActMessageType_Subscribe
    eGDTransActMessageType_Unsubscribe
    eGDTransActMessageType_ChangeAccount
    eGDTransActMessageType_GetPositions
    eGDTransActMessageType_UpdatePositions
    eGDTransActMessageType_GetAccounts
    eGDTransActMessageType_LogonForAccountList
    eGDTransActMessageType_Logon
    
    eGDTransActMessageType_Connected = 100
    eGDTransActMessageType_Disconnected
    eGDTransActMessageType_Subscribed
    eGDTransActMessageType_Unsubscribed
    eGDTransActMessageType_CancelledOrder
    eGDTransActMessageType_ChangedOrder
    eGDTransActMessageType_FilledOrder
    eGDTransActMessageType_ExpiredOrder
    eGDTransActMessageType_SentOrder
    eGDTransActMessageType_GetOrder
    eGDTransActMessageType_GetTrade
    eGDTransActMessageType_AppLoaded
    eGDTransActMessageType_AppUnloaded
    eGDTransActMessageType_Heartbeat
    eGDTransActMessageType_PriceUpdate
    eGDTransActMessageType_Error
    eGDTransActMessageType_Account
    eGDTransActMessageType_Position
    eGDTransActMessageType_ConnectionInfo
    eGDTransActMessageType_GetPosition
    eGDTransActMessageType_AccountList
End Enum

Public Enum eGDTransActLoginModes
    eGDTransActLoginMode_Live = 0
    eGDTransActLoginMode_Demo
    eGDTransActLoginMode_SimLive
End Enum

'Public Enum eGDLindXpressMessageTypes
'    eGDLindXpressMessageType_Connect = 1
'    eGDLindXpressMessageType_Disconnect
'    eGDLindXpressMessageType_AddOrder
'    eGDLindXpressMessageType_AmendOrder
'    eGDLindXpressMessageType_CancelOrder
'    eGDLindXpressMessageType_UnloadApp
'    eGDLindXpressMessageType_GetAccounts
'    eGDLindXpressMessageType_GetOrders
'    eGDLindXpressMessageType_GetFills
'    eGDLindXpressMessageType_GetPositions
'    eGDLindXpressMessageType_GetAllOrders
'    eGDLindXpressMessageType_GetAllFills
'    eGDLindXpressMessageType_GetAllPositions
'    eGDLindXpressMessageType_GetAccountInfo
'    eGDLindXpressMessageType_ChangePassword
'
'    eGDLindXpressMessageType_ConnectionInfo = 100
'    eGDLindXpressMessageType_AppLoaded
'    eGDLindXpressMessageType_AppUnloaded
'    eGDLindXpressMessageType_Heartbeat
'    eGDLindXpressMessageType_Account
'    eGDLindXpressMessageType_Order
'    eGDLindXpressMessageType_Fill
'    eGDLindXpressMessageType_Position
'    eGDLindXpressMessageType_Alert
'    eGDLindXpressMessageType_AccountInfo
'End Enum
'
'Public Enum eGDPfgMessageTypes
'    eGDPfgMessageType_Connect = 1
'    eGDPfgMessageType_Disconnect
'    eGDPfgMessageType_AddOrder
'    eGDPfgMessageType_AmendOrder
'    eGDPfgMessageType_CancelOrder
'    eGDPfgMessageType_UnloadApp
'    eGDPfgMessageType_GetAccounts
'    eGDPfgMessageType_GetOrders
'    eGDPfgMessageType_GetFills
'    eGDPfgMessageType_GetPositions
'    eGDPfgMessageType_GetContracts
'    eGDPfgMessageType_SubscribeBook
'    eGDPfgMessageType_UnsubscribeBook
'    eGDPfgMessageType_CloseTrade
'    eGDPfgMessageType_ChangeTradeStopLimit
'    eGDPfgMessageType_DeleteTradeStopLimit
'    eGDPfgMessageType_GetExchangeAvailable
'    eGDPfgMessageType_GetSingleOrder
'
'    eGDPfgMessageType_ConnectionInfo = 100
'    eGDPfgMessageType_AppLoaded
'    eGDPfgMessageType_AppUnloaded
'    eGDPfgMessageType_Heartbeat
'    eGDPfgMessageType_Order
'    eGDPfgMessageType_AccountR
'    eGDPfgMessageType_OrderR
'    eGDPfgMessageType_FillR
'    eGDPfgMessageType_PositionR
'    eGDPfgMessageType_ContractR
'
'    eGDPfgMessageType_GetBlocks = 1000
'    eGDPfgMessageType_BlockR
'    eGDPfgMessageType_GetAccountInfo
'    eGDPfgMessageType_AccountInfo
'    eGDPfgMessageType_LinkOrders
'    eGDPfgMessageType_OrdersLinked
'    eGDPfgMessageType_UnlinkOrders
'    eGDPfgMessageType_OrdersUnlinked
'    eGDPfgMessageType_CreateSpread = 1010
'    eGDPfgMessageType_SpreadCreated
'    eGDPfgMessageType_RequestForQuote
'    eGDPfgMessageType_SpreadQuote
'    eGDPfgMessageType_RequestSpreadCodes
'    eGDPfgMessageType_SpreadCodes
'    eGDPfgMessageType_RequestSpreadLegs
'    eGDPfgMessageType_SpreadLegs
'End Enum

Public Enum eGDIbMessageTypes
    eGDIbMessageType_Connect = 1
    eGDIbMessageType_Disconnect
    eGDIbMessageType_AddOrder
    eGDIbMessageType_AmendOrder
    eGDIbMessageType_CancelOrder
    eGDIbMessageType_UnloadApp
    eGDIbMessageType_GetAccounts
    eGDIbMessageType_GetOrders
    eGDIbMessageType_GetFills
    eGDIbMessageType_GetPositions
    eGDIbMessageType_GetContracts
    eGDIbMessageType_Subscribe
    eGDIbMessageType_Unsubscribe
    
    eGDIbMessageType_ConnectionInfo = 100
    eGDIbMessageType_AppLoaded
    eGDIbMessageType_AppUnloaded
    eGDIbMessageType_Heartbeat
    eGDIbMessageType_Order
    eGDIbMessageType_AccountR
    eGDIbMessageType_OrderR
    eGDIbMessageType_FillR
    eGDIbMessageType_PositionR
    eGDIbMessageType_ContractR
    eGDIbMessageType_QuoteR
    
    eGDIbMessageType_GetNextValidID = 1000
    eGDIbMessageType_NextValidID
    eGDIbMessageType_GetExchangeAvailable
    eGDIbMessageType_ExchangeAvailable
    eGDIbMessageType_GetContractDetails
    eGDIbMessageType_ContractDetails
    eGDIbMessageType_Fill = 1007
    eGDIbMessageType_Position = 1009
    eGDIbMessageType_AddComboOrder
    eGDIbMessageType_AmendComboOrder = 1012
    eGDIbMessageType_RequestHistory = 1014
    eGDIbMessageType_History = 1015
    eGDIbMessageType_Message = 1017
    eGDIbMessageType_GetSymbolAvailable = 1018
    eGDIbMessageType_SymbolAvailable = 1019
End Enum

Public Enum eGDRithmicMessageTypes
    eGDRithmicMessageType_Connect = 1
    eGDRithmicMessageType_Disconnect
    eGDRithmicMessageType_AddOrder
    eGDRithmicMessageType_AmendOrder
    eGDRithmicMessageType_CancelOrder
    eGDRithmicMessageType_UnloadApp
    eGDRithmicMessageType_GetAccounts
    eGDRithmicMessageType_GetOrders
    eGDRithmicMessageType_GetFills
    eGDRithmicMessageType_GetPositions
    
    eGDRithmicMessageType_ConnectionInfo = 100
    eGDRithmicMessageType_AppLoaded
    eGDRithmicMessageType_AppUnloaded
    eGDRithmicMessageType_Heartbeat
    eGDRithmicMessageType_Order
    eGDRithmicMessageType_AccountR
    eGDRithmicMessageType_OrderR
    eGDRithmicMessageType_FillR
    eGDRithmicMessageType_PositionR
    
    eGDRithmicMessageType_Fill = 1001
    eGDRithmicMessageType_Position = 1003
    eGDRithmicMessageType_ExchList = 1005
    eGDRithmicMessageType_GetTradeRoutes = 1006
    eGDRithmicMessageType_TradeRoute = 1007
    eGDRithmicMessageType_SubscribeTradeRoute = 1008
    eGDRithmicMessageType_UnsubscribeTradeRoute = 1010
    eGDRithmicMessageType_GetContractInfo = 1012
    eGDRithmicMessageType_ContractInfo = 1013
End Enum

Public Enum eGDBrokerMessageTypes
    eGDBrokerMessageType_AppLoaded = 0
    eGDBrokerMessageType_Connect = 1
    eGDBrokerMessageType_ConnectionInfo = 2
    eGDBrokerMessageType_Disconnect = 3
    eGDBrokerMessageType_AddOrder = 5
    eGDBrokerMessageType_Order = 6
    eGDBrokerMessageType_AmendOrder = 7
    eGDBrokerMessageType_CancelOrder = 9
    eGDBrokerMessageType_UnloadApp = 11
    eGDBrokerMessageType_AppUnloaded = 12
    eGDBrokerMessageType_GetAccounts = 13
    eGDBrokerMessageType_AccountRefresh = 14
    eGDBrokerMessageType_GetOrders = 15
    eGDBrokerMessageType_OrderRefresh = 16
    eGDBrokerMessageType_GetFills = 17
    eGDBrokerMessageType_FillRefresh = 18
    eGDBrokerMessageType_GetPositions = 19
    eGDBrokerMessageType_PositionRefresh = 20
    eGDBrokerMessageType_Heartbeat = 22
    eGDBrokerMessageType_Fill = 24
    eGDBrokerMessageType_Position = 26
    eGDBrokerMessageType_CarriedFillRefresh = 28
    eGDBrokerMessageType_Subscribe = 29
    eGDBrokerMessageType_PriceUpdate = 30
    eGDBrokerMessageType_Unsubscribe = 31
    eGDBrokerMessageType_GetSecurityDefinition = 33
    eGDBrokerMessageType_UserRequest = 35
    eGDBrokerMessageType_GetAccountDetails = 37
    eGDBrokerMessageType_AccountDetails = 38
    eGDBrokerMessageType_GetSides = 39
    eGDBrokerMessageType_Sides = 40
    eGDBrokerMessageType_GetTifs = 41
    eGDBrokerMessageType_Tifs = 42
    eGDBrokerMessageType_GetOrderTypes = 43
    eGDBrokerMessageType_OrderTypes = 44
    eGDBrokerMessageType_GetSymbols = 45
    eGDBrokerMessageType_Symbols = 46
    eGDBrokerMessageType_GetNumberOfAccounts = 47
    eGDBrokerMessageType_NumberOfAccounts = 48
    eGDBrokerMessageType_AddOcoOrders = 49
    eGDBrokerMessageType_SpreadFill = 50
    eGDBrokerMessageType_SpecialFill = 52
    eGDBrokerMessageType_RejectMessage = 54
    eGDBrokerMessageType_PositionFillRefresh = 56
    eGDBrokerMessageType_ConsumerInfo = 57
    eGDBrokerMessageType_LoginUrl = 58
    eGDBrokerMessageType_GetTransactions = 59
    eGDBrokerMessageType_Transactions = 60
    
    eGDBrokerMessageType_OecOrderIdChanged = 1000
    eGDBrokerMessageType_CnxGetWorkingOrders = 1001
    eGDBrokerMessageType_CnxWorkingOrderRefresh = 1002
    eGDBrokerMessageType_CnxGetSingleOrder = 1003
    eGDBrokerMessageType_CnxSingleOrderRefresh = 1004
End Enum

'Public Enum eGDConnectionStatus
'    eGDConnectionStatus_Disconnected = 0
'    eGDConnectionStatus_Disconnecting
'    eGDConnectionStatus_Connecting
'    eGDConnectionStatus_Connected
'End Enum

Public Enum eGDStopLossType
    eGDStopLossType_None = 0
    eGDStopLossType_Fixed
    eGDStopLossType_Trail
    eGDStopLossType_BreakEven
    eGDStopLossType_TradeSense
End Enum

Public Enum eGDProfitTargetType
    eGDProfitTargetType_Standard = 0
    eGDProfitTargetType_TradeSense
End Enum

Public Enum eGDOrderRejectOption
    eGDOrderRejectOption_Flatten = 0
    eGDOrderRejectOption_Ask
    eGDOrderRejectOption_Nothing
End Enum

Public Enum eGDFlattenItemType
    eGDFlattenItem_Manual = 0
    eGDFlattenItem_AutoExit
    eGDFlattenItem_AutoStrategy
End Enum

Public Enum eGDTradeTrackerTabs
    eGDTradeTrackerTab_Account = 0
    eGDTradeTrackerTab_Orders
    eGDTradeTrackerTab_Transactions
    eGDTradeTrackerTab_Trades
    eGDTradeTrackerTab_Positions
    eGDTradeTrackerTab_ActivityLog
End Enum

'Public Enum eGDTypeOfAccount
'    eGDTypeOfAccount_Simulated = 0
'    eGDTypeOfAccount_BrokerLive
'    eGDTypeOfAccount_BrokerDemo
'End Enum

Public Enum eGDTradeRuleTypes
    eGDTradeRuleType_Entry = 0
    eGDTradeRuleType_Exit
End Enum

'Public Enum eGDOrderLinkStatus
'    eGDOrderLinkStatus_New = 0
'    eGDOrderLinkStatus_LinkSent
'    eGDOrderLinkStatus_Confirmed
'    eGDOrderLinkStatus_UnlinkSent
'    eGDOrderLinkStatus_Parked
'End Enum

'Public Enum eGDWaitListCommands
'    eGDWaitListCommand_Link = 0
'    eGDWaitListCommand_Cancel
'    eGDWaitListCommand_ParkOne
'    eGDWaitListCommand_ParkBoth
'    eGDWaitListCommand_AmendOne
'    eGDWaitListCommand_AmendBoth
'    eGDWaitListCommand_SubmitOne
'    eGDWaitListCommand_SubmitBoth
'End Enum

Public Enum eGDOptionFill
    eGDOptionFill_BidOrAsk = 0
    eGDOptionFill_Midpoint
    eGDOptionFill_TwoThirds
End Enum

Public Enum eGDOptionOpenEquity
    eGDOptionOpenEquity_UseBidAsk = 0
    eGDOptionOpenEquity_UseLast = 1
End Enum

Public Enum eGDOrderAction
    eGDOrderAction_LongEntry = 0
    eGDOrderAction_LongExit
    eGDOrderAction_ShortEntry
    eGDOrderAction_ShortExit
End Enum

Public Enum eGDTradingMenu
    eGDTradingMenu_Connect = 0
    eGDTradingMenu_Disconnect
    eGDTradingMenu_SwitchAccounts
    eGDTradingMenu_SwitchAccountsMode
    eGDTradingMenu_ConnectInfo
    eGDTradingMenu_ChangePassword
    eGDTradingMenu_Refresh
    eGDTradingMenu_ViewActivity
    eGDTradingMenu_BrokerView
    eGDTradingMenu_ViewOnline
    eGDTradingMenu_VerifyPositions
    eGDTradingMenu_AccountDetails
    eGDTradingMenu_NumItems
End Enum

Public Enum eGDDebitCredit
    eGDDebitCredit_Credit = 0
    eGDDebitCredit_Debit
    eGDDebitCredit_Even
    eGDDebitCredit_Market
End Enum

'Public Enum eGDJournalImageTypes
'    eGDJournalImageType_Chart = 0
'    eGDJournalImageType_SummaryReport
'    eGDJournalImageType_OptionNavOrder
'End Enum
'
'Public Enum eGDJournalCategoryTypes
'    eGDJournalCategoryType_Note = -1
'    eGDJournalCategoryType_MoneyCode = 0
'    eGDJournalCategoryType_CustomChecklist = 1
'End Enum

'Public Enum eGDTurnkeyMessage
'    eGDTurnkeyMessage_Heartbeat = 0
'    eGDTurnkeyMessage_Connect = 1
'    eGDTurnkeyMessage_ConnectionStatus = 2
'    eGDTurnkeyMessage_Disconnect = 3
'    eGDTurnkeyMessage_GetFeedYards = 5
'    eGDTurnkeyMessage_FeedYard = 6
'    eGDTurnkeyMessage_GetCustomers = 7
'    eGDTurnkeyMessage_Customer = 8
'    eGDTurnkeyMessage_GetLots = 9
'    eGDTurnkeyMessage_Lot = 10
'    eGDTurnkeyMessage_GetOrders = 11
'    eGDTurnkeyMessage_Order = 12
'    eGDTurnkeyMessage_GetAssociatedFills = 13
'    eGDTurnkeyMessage_AssociatedFill = 14
'    eGDTurnkeyMessage_GetPositions = 15
'    eGDTurnkeyMessage_Position = 16
'    eGDTurnkeyMessage_GetAccounts = 17
'    eGDTurnkeyMessage_Account = 18
'    eGDTurnkeyMessage_UpdateAccount = 19
'    eGDTurnkeyMessage_UpdateOrder = 21
'    eGDTurnkeyMessage_UpdateFill = 23
'    eGDTurnkeyMessage_UpdatePosition = 25
'    eGDTurnkeyMessage_GetTrades = 27
'    eGDTurnkeyMessage_Trade = 28
'    eGDTurnkeyMessage_DeleteTrade = 30
'    eGDTurnkeyMessage_DeleteOrder = 32
'    eGDTurnkeyMessage_DeleteAssociatedFill = 34
'    eGDTurnkeyMessage_AppLoaded = 36
'    eGDTurnkeyMessage_UnloadApp = 37
'    eGDTurnkeyMessage_AppUnloaded = 38
'    eGDTurnkeyMessage_GetAllLotColumns = 39
'    eGDTurnkeyMessage_LotColumns = 40
'    eGDTurnkeyMessage_GetVisibleLotColumns = 41
'    eGDTurnkeyMessage_VisibleLotColumns = 42
'    eGDTurnkeyMessage_RemoveAccount = 43
'    eGDTurnkeyMessage_GetAllFills = 45
'    eGDTurnkeyMessage_Fill = 46
'    eGDTurnkeyMessage_AssociateFill = 47
'    eGDTurnkeyMessage_GetVisibleLots = 49
'    eGDTurnkeyMessage_VisibleLots = 50
'    eGDTurnkeyMessage_GetLotColumnCategories = 51
'    eGDTurnkeyMessage_LotColumnCategories = 52
'    eGDTurnkeyMessage_GetLotContentDetails = 53
'    eGDTurnkeyMessage_LotContentDetails = 54
'    eGDTurnkeyMessage_GenesisCustomerInfo = 56
'    eGDTurnkeyMessage_UpdateLotContentDetails = 57
'    eGDTurnkeyMessage_AddFeedYards = 59
'    eGDTurnkeyMessage_AddLots = 61
'    eGDTurnkeyMessage_AddCustomers = 63
'    eGDTurnkeyMessage_GetLotColumnSubCategories = 65
'    eGDTurnkeyMessage_LotColumnSubCategories = 66
'    eGDTurnkeyMessage_GetDetailOptions = 67
'    eGDTurnkeyMessage_DetailOptions = 68
'    eGDTurnkeyMessage_GetRations = 69
'    eGDTurnkeyMessage_Ration = 70
'    eGDTurnkeyMessage_GetIngredients = 71
'    eGDTurnkeyMessage_Ingredient = 72
'    eGDTurnkeyMessage_UpdateRation = 73
'    eGDTurnkeyMessage_UpdateIngredient = 75
'
'    eGDTurnkeyMessage_UpdateVisibleFeedYards = 1007
'    eGDTurnkeyMessage_UpdateVisibleCustomers = 1009
'    eGDTurnkeyMessage_GetGenesisCustomers = 1011
'    eGDTurnkeyMessage_GenesisCustomers = 1012
'    eGDTurnkeyMessage_GetAllFeedYards = 1013
'    eGDTurnkeyMessage_AllFeedYards = 1014
'    eGDTurnkeyMessage_GetAllCustomers = 1015
'    eGDTurnkeyMessage_AllCustomers = 1016
'    eGDTurnkeyMessage_GetVisibleFeedYards = 1017
'    eGDTurnkeyMessage_VisibleFeedYards = 1018
'    eGDTurnkeyMessage_GetVisibleCustomers = 1019
'    eGDTurnkeyMessage_VisibleCustomers = 1020
'    eGDTurnkeyMessage_UpdateGenesisCustomer = 1021
'    eGDTurnkeyMessage_GetAllLotColumnsAdmin = 1023
'    eGDTurnkeyMessage_LotColumnsAdmin = 1024
'    eGDTurnkeyMessage_GetVisibleLotColumnsAdmin = 1025
'    eGDTurnkeyMessage_VisibleLotColumnsAdmin = 1026
'    eGDTurnkeyMessage_UpdateVisibleLotColumns = 1027
'    eGDTurnkeyMessage_GetDefaultVisibleLotColumnsAdmin = 1029
'    eGDTurnkeyMessage_DefaultVisibleLotColumnsAdmin = 1030
'    eGDTurnkeyMessage_UpdateDefaultVisibleLotColumns = 1031
'End Enum

'Public Enum eGDTurnkeyCustomerType
'    eGDTurnkeyCustomerType_TurnkeyCustomer = 0
'    eGDTurnkeyCustomerType_Broker = 1
'    eGDTurnkeyCustomerType_Admin = 2
'    eGDTurnkeyCustomerType_TurnkeyFeedYard = 3
'    eGDTurnkeyCustomerType_CattleNavCustomer = 4
'    eGDTurnkeyCustomerType_CattleNavFeedYard = 5
'    eGDTurnkeyCustomerType_EitherFeedYard = 6
'End Enum

Public Enum eGDFlattenQueueOperations
    eGDFlattenQueueOperation_Flatten = -1
    eGDFlattenQueueOperation_CancelAll = 0
    eGDFlattenQueueOperation_Reverse = 1
End Enum

Public Enum eGDFilterDirection
    eGDFilterDirection_All = 0
    eGDFilterDirection_Longs
    eGDFilterDirection_Shorts
End Enum

Public Enum eGDFilterTradeType
    eGDFilterTradeType_All = 0
    eGDFilterTradeType_Real
    eGDFilterTradeType_Sim
End Enum

Declare Function gntConfigure Lib "NavTrade.DLL" (ByVal hParms As Long, ByVal hDataIn As Long, ByVal hDataOut As Long, ByVal hReturn As Long) As Long
Declare Function gntGetInfo Lib "NavTrade.DLL" (ByVal hParms As Long, ByVal hDesc As Long, ByVal hDataOut As Long, ByVal hResult As Long) As Long
Declare Function gntTrade Lib "NavTrade.DLL" (ByVal hParms As Long, ByVal hDataIn As Long, ByVal hDataOut As Long, ByVal hReturn As Long) As Long

Declare Function StartSalmonTS Lib "SalmonClient.dll" Alias "SalmonTradeSrv_Start" (ByVal hParms As Long) As Long
Declare Function StopSalmonTS Lib "SalmonClient.dll" Alias "SalmonTradeSrv_Stop" () As Long
Declare Function SendSalmonTS Lib "SalmonClient.dll" Alias "SalmonTradeSrv_SendOrders" (ByVal hParms As Long) As Long

Public Function TransactMessageType(ByVal nType As eGDTransActMessageTypes) As Long
    TransactMessageType = nType
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderRejectOptionString
'' Description: Return the string description of the given reject option
'' Inputs:      Order Reject Option
'' Returns:     String description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderRejectOptionString(ByVal nOrderRejectOption As eGDOrderRejectOption) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    strReturn = ""
    
    Select Case nOrderRejectOption
        Case eGDOrderRejectOption_Flatten
            strReturn = "Flatten"
        Case eGDOrderRejectOption_Ask
            strReturn = "Ask"
        Case eGDOrderRejectOption_Nothing
            strReturn = "Nothing"
    End Select
    
    OrderRejectOptionString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.OrderRejectOptionString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderType
'' Description: Determine the string description of the given order type
'' Inputs:      Order Type
'' Returns:     String description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderType(ByVal eOrderType As eTT_OrderType) As String
On Error GoTo ErrSection:

    Select Case eOrderType
        Case eTT_OrderType_Market
            OrderType = "Market"
        Case eTT_OrderType_Stop
            OrderType = "Stop"
        Case eTT_OrderType_Limit
            OrderType = "Limit"
        Case eTT_OrderType_StopWithLimit
            OrderType = "Stop With Limit"
        Case eTT_OrderType_MarketOnClose
            OrderType = "Market On Close"
        Case eTT_OrderType_StopCloseOnly
            OrderType = "Stop Close Only"
        Case eTT_OrderType_LimitCloseOnly
            OrderType = "Limit Close Only"
        Case eTT_OrderType_StopWithLimitCloseOnly
            OrderType = "Stop With Limit Close Only"
        Case eTT_OrderType_MIT
            OrderType = "MIT"
        Case eTT_OrderType_Adjustment
            OrderType = "Other"
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.OrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTypeFromString
'' Description: Determine the order type given the string description
'' Inputs:      String description
'' Returns:     Order Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderTypeFromString(ByVal strOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Select Case UCase(strOrderType)
        Case "MARKET"
            OrderTypeFromString = eTT_OrderType_Market
        Case "STOP"
            OrderTypeFromString = eTT_OrderType_Stop
        Case "LIMIT"
            OrderTypeFromString = eTT_OrderType_Limit
        Case "STOP WITH LIMIT"
            OrderTypeFromString = eTT_OrderType_StopWithLimit
        Case "MIT"
            OrderTypeFromString = eTT_OrderType_MIT
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.OrderTypeFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderStatus
'' Description: Determine the string description of the given order status
'' Inputs:      Order Status
'' Returns:     String description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderStatus(ByVal nStatus As eTT_OrderStatus) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case nStatus
        Case eTT_OrderStatus_Open
            strReturn = "New"
        Case eTT_OrderStatus_Partial
            strReturn = "Partial Fill"
        Case eTT_OrderStatus_Filled
            strReturn = "Filled"
        Case eTT_OrderStatus_Cancelled
            strReturn = "Cancelled"
        Case eTT_OrderStatus_Queued
            strReturn = "Queued"
        Case eTT_OrderStatus_Sent
            strReturn = "Sent"
        Case eTT_OrderStatus_Working
            strReturn = "Working"
        Case eTT_OrderStatus_Rejected
            strReturn = "Rejected"
        Case eTT_OrderStatus_BalCancelled
            strReturn = "Balance Cancelled"
        Case eTT_OrderStatus_CancelPending
            strReturn = "Cancel Pending"
        Case eTT_OrderStatus_AmendPending
            strReturn = "Amend Pending"
        Case eTT_OrderStatus_UnconfirmedFilled
            strReturn = "Unconfirmed Filled"
        Case eTT_OrderStatus_UnconfirmedPartial
            strReturn = "Unconfirmed Partial Fill"
        Case eTT_OrderStatus_Held
            strReturn = "Held"
        Case eTT_OrderStatus_CancelHeld
            strReturn = "Cancel Held"
        Case eTT_OrderStatus_Error
            strReturn = "Error"
        Case eTT_OrderStatus_Amended
            strReturn = "Amended"
        Case eTT_OrderStatus_Expired
            strReturn = "Expired"
        Case eTT_OrderStatus_Frozen
            strReturn = "Frozen"
        Case eTT_OrderStatus_ParkPending
            strReturn = "Park Pending"
        Case eTT_OrderStatus_TriggerPending
            strReturn = "Pending Trigger"
        Case eTT_OrderStatus_Approved
            strReturn = "Approved"
        Case eTT_OrderStatus_BrokerParked
            strReturn = "Parked at Broker"
        Case eTT_OrderStatus_OverFilled
            strReturn = "Over Filled"
        Case eTT_OrderStatus_Parked
            strReturn = "Parked"
        Case eTT_OrderStatus_PreSubmitted
            strReturn = "PreSubmitted"
        Case eTT_OrderStatus_Inactive
            strReturn = "Inactive"
        Case eTT_OrderStatus_Suspended
            strReturn = "Suspended"
        Case eTT_OrderStatus_DataPending
            strReturn = "Pending Data"
    End Select
    
    OrderStatus = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.OrderStatus"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderStatusFromString
'' Description: Determine the order status given the string description
'' Inputs:      String description
'' Returns:     Order Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderStatusFromString(ByVal strOrderStatus As String) As eTT_OrderStatus
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderStatus      ' Return value for the function

    Select Case UCase(strOrderStatus)
        Case "OPEN", "NEW"
            nReturn = eTT_OrderStatus_Open
        Case "PARTIAL FILL"
            nReturn = eTT_OrderStatus_Partial
        Case "FILLED"
            nReturn = eTT_OrderStatus_Filled
        Case "CANCELLED"
            nReturn = eTT_OrderStatus_Cancelled
        Case "QUEUED"
            nReturn = eTT_OrderStatus_Queued
        Case "SENT"
            nReturn = eTT_OrderStatus_Sent
        Case "WORKING"
            nReturn = eTT_OrderStatus_Working
        Case "REJECTED"
            nReturn = eTT_OrderStatus_Rejected
        Case "BALANCE CANCELLED"
            nReturn = eTT_OrderStatus_BalCancelled
        Case "CANCEL PENDING"
            nReturn = eTT_OrderStatus_CancelPending
        Case "AMEND PENDING"
            nReturn = eTT_OrderStatus_AmendPending
        Case "UNCONFIRMED FILLED"
            nReturn = eTT_OrderStatus_UnconfirmedFilled
        Case "UNCONFIRMED PARTIAL FILL"
            nReturn = eTT_OrderStatus_UnconfirmedPartial
        Case "HELD"
            nReturn = eTT_OrderStatus_Held
        Case "CANCEL HELD"
            nReturn = eTT_OrderStatus_CancelHeld
        Case "ERROR"
            nReturn = eTT_OrderStatus_Error
        Case "AMENDED"
            nReturn = eTT_OrderStatus_Amended
        Case "EXPIRED"
            nReturn = eTT_OrderStatus_Expired
        Case "FROZEN"
            nReturn = eTT_OrderStatus_Frozen
        Case "PARK PENDING"
            nReturn = eTT_OrderStatus_ParkPending
        Case "PENDING TRIGGER"
            nReturn = eTT_OrderStatus_TriggerPending
        Case "APPROVED"
            nReturn = eTT_OrderStatus_Approved
        Case "PARKED AT BROKER"
            nReturn = eTT_OrderStatus_BrokerParked
        Case "OVER FILLED"
            nReturn = eTT_OrderStatus_OverFilled
        Case "PARKED"
            nReturn = eTT_OrderStatus_Parked
        Case "PRESUBMITTED"
            nReturn = eTT_OrderStatus_PreSubmitted
        Case "INACTIVE"
            nReturn = eTT_OrderStatus_Inactive
        Case "SUSPENDED"
            nReturn = eTT_OrderStatus_Suspended
        Case "PENDING DATA"
            nReturn = eTT_OrderStatus_DataPending
    End Select
    
    OrderStatusFromString = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.OrderStatusFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimeInForce
'' Description: Determine the string description of the given time in force
'' Inputs:      Time In Force
'' Returns:     String description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TimeInForce(ByVal nTimeInForce As eTT_TimeInForce) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nTimeInForce
        Case eTT_TimeInForce_Day
            strReturn = "Day"
        Case eTT_TimeInForce_GTC
            strReturn = "GTC"
        Case eTT_TimeInForce_GTD
            strReturn = "GTD"
    End Select
    
    TimeInForce = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.TimeInForce"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimeInForceFromString
'' Description: Determine the time in force for the given description
'' Inputs:      String description
'' Returns:     Time in Force
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TimeInForceFromString(ByVal strTimeInForce As String) As eTT_TimeInForce
On Error GoTo ErrSection:

    Dim nReturn As eTT_TimeInForce      ' Return value for the function
    
    Select Case UCase(strTimeInForce)
        Case "DAY"
            nReturn = eTT_TimeInForce_Day
        Case "GTC"
            nReturn = eTT_TimeInForce_GTC
        Case "GTD"
            nReturn = eTT_TimeInForce_GTD
    End Select
    
    TimeInForceFromString = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.TimeInForceFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DebitCreditString
'' Description: Determine the string description of the given debit/credit flag
'' Inputs:      Debit/Credit flag
'' Returns:     String description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DebitCreditString(ByVal nDebitCredit As eGDDebitCredit) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    strReturn = ""
    Select Case nDebitCredit
        Case eGDDebitCredit_Credit
            strReturn = "Credit"
        Case eGDDebitCredit_Debit
            strReturn = "Debit"
        Case eGDDebitCredit_Even
            strReturn = "Even"
        Case eGDDebitCredit_Market
            strReturn = "Market"
    End Select
    
    DebitCreditString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.DebitCreditString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DebitCreditFromString
'' Description: Determine the debit/credit flag for the given description
'' Inputs:      String description
'' Returns:     Debit/Credit flag
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DebitCreditFromString(ByVal strDebitCredit As String) As eGDDebitCredit
On Error GoTo ErrSection:

    Dim nDebitCredit As eGDDebitCredit  ' Return value for the function
    
    nDebitCredit = -1
    Select Case UCase(strDebitCredit)
        Case "CREDIT"
            nDebitCredit = eGDDebitCredit_Credit
        Case "DEBIT"
            nDebitCredit = eGDDebitCredit_Debit
        Case "EVEN"
            nDebitCredit = eGDDebitCredit_Even
        Case "MARKET"
            nDebitCredit = eGDDebitCredit_Market
    End Select
    
    DebitCreditFromString = nDebitCredit

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.DebitCreditFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PlaceOrder
'' Description: Place an order through the edit order form
'' Inputs:      Order
'' Returns:     Return from the edit order dialog
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PlaceOrder(Order As cPtOrder) As eGDEditOrderReturn
On Error GoTo ErrSection:

    Dim nReturn As eGDEditOrderReturn   ' Return from the edit order form
    
    nReturn = frmTTEditOrder.ShowMe(Order)
    Select Case nReturn
        Case eGDEditOrderReturn_Submit
            If Order.HasTrigger = False Then
                SubmitOrder Order
            Else
                SetTriggerOrderStatus Order
            End If
        
        Case eGDEditOrderReturn_Park
            ParkOrder Order
    
    End Select
    
    PlaceOrder = nReturn
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.PlaceOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOpenOrder
'' Description: Figure out based on status if this is an "open" order
'' Inputs:      Order Status
'' Returns:     True if Open, False if Closed
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsOpenOrder(ByVal nStatus As eTT_OrderStatus, Optional ByVal bAllowPending As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    Select Case nStatus
        Case eTT_OrderStatus_Open, eTT_OrderStatus_Partial, eTT_OrderStatus_Working, eTT_OrderStatus_Frozen, _
             eTT_OrderStatus_Approved, eTT_OrderStatus_BrokerParked, eTT_OrderStatus_Parked, eTT_OrderStatus_PreSubmitted, _
             eTT_OrderStatus_Suspended, eTT_OrderStatus_Held
            bReturn = True
             
        Case eTT_OrderStatus_AmendPending, eTT_OrderStatus_CancelPending, eTT_OrderStatus_Queued, _
             eTT_OrderStatus_Sent, eTT_OrderStatus_ParkPending, eTT_OrderStatus_TriggerPending, _
             eTT_OrderStatus_DataPending
            bReturn = bAllowPending
             
        Case Else
            bReturn = False
    End Select
    
    IsOpenOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.IsOpenOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateOrder
'' Description: Allow the user to create an order
'' Inputs:      None
'' Returns:     none
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateOrder(Optional ByVal strSymbol As String = "", _
    Optional ByVal lAccountID As Long = 0&, Optional nBuy As Byte = 255, _
    Optional ChartOrder As cPtOrder, Optional ByVal lAutoTradeItemID As Long = 0&, _
    Optional ByVal strSource As String = "", Optional ByVal strFeedYardLotID As String = "") As eGDEditOrderReturn
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order to create
    Dim nReturn As eGDEditOrderReturn   ' Return from the create order call
    Dim frm As Form                     ' Chart form

    If ChartOrder Is Nothing Then
        If Len(strSymbol) = 0 Then
            If Not ActiveChart Is Nothing Then strSymbol = RollSymbolForDate(GetSymbol(ActiveChart.SymbolID), Date)
        End If
        
        If lAccountID = 0& Then
            lAccountID = DefaultAccount
            If Not ActiveChart Is Nothing Then
                Set frm = ActiveChart
                If (frm.OrderBarMode = eOrdBarMode_Order) Or (frm.OrderBarMode = eOrdBarMode_BrokerDisconnect) Then
                    lAccountID = ActiveChart.TradeAccountID
                End If
            End If
        End If
    
        Order.SymbolOrSymbolID = strSymbol
        Order.AccountID = lAccountID
        Order.AutoTradeItemID = lAutoTradeItemID
        Order.OrderID = 0&
    Else
        Order.SymbolOrSymbolID = ChartOrder.Symbol
        Order.AccountID = ChartOrder.AccountID
        Order.OrderID = ChartOrder.OrderID
        Order.OrderType = ChartOrder.OrderType
        Order.Buy = ChartOrder.Buy
        Order.OrderPrice(False) = ChartOrder.OrderPrice(False)
        If ChartOrder.Quantity > 0 Then Order.Quantity = ChartOrder.Quantity
    End If
    
    nReturn = frmTTEditOrder.ShowMe(Order, nBuy, , , , , , strFeedYardLotID)
    Select Case nReturn
        Case eGDEditOrderReturn_Submit
            If Len(strSource) > 0 Then
                g.Broker.BrokerDebug g.Broker.AccountTypeForID(Order.AccountID), "Creating Order from " & strSource & ": " & Order.OrderText, True
            End If
            
            If Order.HasTrigger = False Then
                If (Order.TrailAmount <> 0) And (Order.OrderPrice(True) = kNullData) Then
                    Order.IsSnapshot = True
                    If Order.OrderDate = 0# Then
                        Order.OrderDate = Order.BrokerDate(CurrentTime("", Order.Symbol))
                    End If
                    Order.ChangeOrderStatus eTT_OrderStatus_Open
                Else
                    SubmitOrder Order
                End If
            Else
                Order.IsSnapshot = True
                If Order.OrderDate = 0# Then
                    Order.OrderDate = Order.BrokerDate(CurrentTime("", Order.Symbol))
                End If
                
                SetTriggerOrderStatus Order
            End If
            
        Case eGDEditOrderReturn_Park
            If Len(strSource) > 0 Then
                g.Broker.BrokerDebug g.Broker.AccountTypeForID(Order.AccountID), "Creating Order from " & strSource & ": " & Order.OrderText, True
            End If
            ParkOrder Order
            
    End Select

    CreateOrder = nReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mTradeTracker.CreateOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshOrder
'' Description: Refresh the given order in whatever form(s) are loaded
'' Inputs:      Order to Refresh, Received Fill?, Order was Parked?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshOrder(Order As cPtOrder, Optional ByVal bReceivedFill As Boolean = False, Optional ByVal bRefreshCharts As Boolean = True, Optional ByVal bWasParked As Boolean = False)
On Error GoTo ErrSection:

    ' If the edit order form is up with this order, tell it to refresh itself...
    If FormIsLoaded("frmTTEditOrder") Then
        If frmTTEditOrder.OrderID = Order.OrderID Then
            frmTTEditOrder.RefreshOrder Order
        End If
    End If

    If Not g.OrderStrategies Is Nothing Then
        g.OrderStrategies.RefreshOrder Order
    End If
    
    If Not g.CondOrders Is Nothing Then
        'g.CondOrders.RefreshOrder Order
        
        ' DAJ 10/10/2013: Moved this block of code over from cWorkingOrders.OrderToGrid.  The order over
        ' there was a copy and the Bars Handle was different than the one that got added to the
        ' g.RealTime.TickBuffers so the data wasn't updating...
        If ((Order.IsConditional(True) = True) And ((Order.Status = eTT_OrderStatus_TriggerPending) Or (Order.Status = eTT_OrderStatus_DataPending))) Or _
           (((Order.TrailAmount <> 0) Or (Order.ExpireTime <> 0)) And (IsOpenOrder(Order.Status) = True) And (Order.Status <> eTT_OrderStatus_Parked)) Then
            g.CondOrders(Str(Order.OrderID)) = Order
        ElseIf g.CondOrders.Exists(Str(Order.OrderID)) And (Order.Status <> eTT_OrderStatus_TriggerPending) And (Order.TrailAmount = 0) Then
            g.CondOrders.Remove Str(Order.OrderID)
        End If
    End If
    
    If Not g.TsoGroups Is Nothing Then
        g.TsoGroups.RefreshOrder Order
    End If
    
    If (Order.ExitPos > 0) And (Not g.ExitAllOrders Is Nothing) Then
        g.ExitAllOrders.UpdateOrder Order
    End If
    
    If FormIsLoaded("frmOnlineBroker") Then
        If bRefreshCharts = True Then
            If Order.AutoTradeItemID = 0 Then
                If InStr(frmOnlineBroker.tmrUpdateCharts.Tag, "," & Str(Order.SymbolID) & ",") = 0 Then
                    frmOnlineBroker.tmrUpdateCharts.Tag = frmOnlineBroker.tmrUpdateCharts.Tag & "," & Str(Order.SymbolID) & ","
                    frmOnlineBroker.tmrUpdateCharts.Enabled = True
                End If
            End If
        End If
    End If
    
    ' Send the order to the order links object in case it is part of a link...
    'If (Not g.OrderLinks Is Nothing) And (Order.Broker = eTT_AccountType_PFG) Then
    '    g.OrderLinks.RefreshOrder Order, bWasParked
    'End If
    
    g.CattleBridge.Broker_Order g.TnCattle.GenesisOrderToTurnkey(Order)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.RefreshOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditOrder
'' Description: Allow the user to edit the given order
'' Inputs:      Order, Show Buy?, Show Edit Form?, Return Code, Show Message?
'' Returns:     True if submitted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EditOrder(Order As cPtOrder, Optional ByVal bShowBuy As Boolean = False, _
    Optional ByVal bShowEditForm As Boolean = True, _
    Optional ByVal eReturnCode As eGDEditOrderReturn = eGDEditOrderReturn_Cancel, _
    Optional ByVal bShowCanModifyMsg As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Was the order submitted or parked successfully?
    Dim nReturn As eGDEditOrderReturn   ' Return from the edit order call
    Dim bNewOrder As Boolean            ' Is this a new order?
    Dim OriginalOrder As New cPtOrder   ' Original version of the order
    Dim NewOrder As New cPtOrder        ' Order to edit
    
    bReturn = False
    
    If OriginalOrder.Load(Order.OrderID) = False Then
        Set OriginalOrder = Order.MakeCopy
    End If
    Set NewOrder = SetupOrderToEdit(OriginalOrder, bNewOrder)
    If Not NewOrder Is Nothing Then
        If Order.Quantity <> NewOrder.Quantity Then
            NewOrder.Quantity = Order.Quantity
        End If
        
        If Order.StopPrice <> NewOrder.StopPrice Then
            NewOrder.StopPrice = Order.StopPrice
        End If
        If Order.LimitPrice <> NewOrder.LimitPrice Then
            NewOrder.LimitPrice = Order.LimitPrice
        End If
    End If
        
    If CanMoveOrder(NewOrder, bShowCanModifyMsg, Not bShowEditForm) = False Then
        nReturn = eGDEditOrderReturn_Cancel
    ElseIf bShowEditForm Then
        If bShowBuy Then
            nReturn = frmTTEditOrder.ShowMe(NewOrder, Abs(Order.Buy), , , , , False)
        Else
            nReturn = frmTTEditOrder.ShowMe(NewOrder, , , , , , False)
        End If
    Else
        nReturn = eReturnCode
    End If
    
    Select Case nReturn
        Case eGDEditOrderReturn_Park
            ParkOrder NewOrder
            AdjustTriggers OriginalOrder, NewOrder
            bReturn = True
            
        Case eGDEditOrderReturn_Submit
            SubmitAmend OriginalOrder, NewOrder
            bReturn = True
            
        Case eGDEditOrderReturn_Cancel
            If bNewOrder Then
                NewOrder.Delete
            End If
            bReturn = False
            
    End Select
    
    EditOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.EditOrder"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Allow the user to cancel the given order
'' Inputs:      Order to Cancel, Confirm Cancel?, Called from ID, User Cancel?,
''              Cancelling All?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelOrder(Order As cPtOrder, Optional ByVal bConfirm As Boolean = True, Optional ByVal lCalledFromID As Long = 0&, Optional ByVal bUserCancel As Boolean = False, Optional ByVal bCancellingAll As Boolean = False)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strReturn As String             ' Return from infbox
    Dim nStatus As eTT_OrderStatus      ' Current order status
    Dim OtherOrder As cPtOrder          ' Other order from an OCO situation
    Dim BrokerObj As cBroker            ' Broker object

    strReturn = "Y"
    If (bConfirm = True) And (g.Broker.ConfirmManual = True) Then
        strReturn = InfBox("Are you sure that you want to cancel|" & Order.OrderText & "?|", "?", "+Yes|-No", "Order Cancel Confirmation")
        g.Broker.BrokerDebug Order.Broker, vbTab & "Cancel Confirmation: " & strReturn
        
        ' Reload the order in case something happened to it while the confirmation dialog was up
        DoEvents
        Order.Reload
        
        ' If the order status is now closed, don't proceed with the Cancel.  Warn the user if the order
        ' is in any closed state except for Cancelled.
        If IsOpenOrder(Order.Status) = False Then
            If Order.Status <> eTT_OrderStatus_Cancelled Then
                InfBox "The order was not Cancelled because the status changed to '" & OrderStatus(Order.Status) & "'.", "i", , "Order Cancel Information"
            End If
            
            strReturn = "N"
        End If
    End If

    If strReturn = "Y" Then
        g.OrderStrategies.CancelRequested Order
    
        ' Don't allow a user to cancel a working market order with a live brokerage.  This can lead
        ' to some serious issues with the broker server software.  Do allow the cancel, however, if
        ' it is a Parked or Trigger Pending order...
        ' 04/28/2015 DAJ: We will allow a market order to get cancelled on a live account if it is
        ' in 'Suspended' or 'PreSubmitted' status...
        If (Order.OrderType = eTT_OrderType_Market) And (g.Broker.IsLiveAccount(Order.Broker) = True) And (HasBeenSent(Order.Status) = True) And (Order.Status <> eTT_OrderStatus_PreSubmitted) And (Order.Status <> eTT_OrderStatus_Suspended) Then
            g.Broker.BrokerDebug Order.Broker, "Cannot cancel '" & Order.OrderText(True, True, True) & "' because it is a market order"
            frmOnlineBroker.AddDialogMessage "You cannot cancel a market order", "!", , "Order Cancel Error"
            
        ElseIf NotSent(Order.Status) Then
            nStatus = Order.Status
            Order.ChangeOrderStatus eTT_OrderStatus_Cancelled
        
        ElseIf OrderIsPending(Order) Then
'            If (Order.Broker = eTT_AccountType_PFG) And (Len(Order.BrokerID) > 0) Then
'                g.Broker.BrokerDebug Order.Broker, "Calling for single order refresh for " & Order.BrokerID & " because cancel called on pending order"
'                g.PFG.GetSingleOrder Order.BrokerID
'            ElseIf (Order.Broker = eTT_AccountType_CtgPfg) And (Len(Order.BrokerID) > 0) Then
'                g.Broker.BrokerDebug Order.Broker, "Calling for single order refresh for " & Order.BrokerID & " because cancel called on pending order"
'                g.CtgPfg.GetSingleOrder Order.BrokerID
'            ElseIf (Order.Broker = eTT_AccountType_FintecPfg) And (Len(Order.BrokerID) > 0) Then
'                g.Broker.BrokerDebug Order.Broker, "Calling for single order refresh for " & Order.BrokerID & " because cancel called on pending order"
'                g.FintecPfg.GetSingleOrder Order.BrokerID
'            Else
                nStatus = Order.Status
                Order.ChangeOrderStatus eTT_OrderStatus_Cancelled
                g.Broker.GetOrders Order.Broker, Order.AccountID
'            End If
            
        ElseIf g.Broker.ConnectionStatusForAccount(Order.AccountID) <> eGDConnectionStatus_Connected Then
            Order.Message = "Not currently connected to " & g.Broker.BrokerName(Order.Broker) & " account " & g.Broker.AccountNameForID(Order.AccountID)
            g.Broker.ShowNotConnectedError Order.AccountID, Order.Broker, "CancelOrder", True
                
        Else
            If bCancellingAll = False Then
                ' DAJ 08/20/2009: If the user decides to cancel one side of an Order-Cancel-Order,
                ' ask them if they wisk to cancel the other side as well...
                If (bUserCancel = True) And (Order.AutoTradeItemID = 0&) And (Order.IsAutoExit = False) Then
                    If (Order.CancelOrderID <> 0) And (Order.CancelOrderID <> lCalledFromID) Then
                        Set OtherOrder = New cPtOrder
                        If OtherOrder.Load(Order.CancelOrderID) Then
                            If IsOpenOrder(OtherOrder.Status) Then
                                If InfBox("You have chosen to cancel one side of an Order-Cancel-Order.  Would you like to cancel the other one as well?", "?", "+Yes|-No", "Order Cancel Order") = "Y" Then
                                    CancelOrder OtherOrder, False, Order.OrderID
                                End If
                            End If
                        End If
                        
                    ' DAJ 09/14/2009: If the user decides to cancel one side of an Order-Cancel-Order
                    ' that is held at the broker, ask them if they wish to cancel the other side as
                    ' well.  If they do not, we need to unlink before cancelling the order...
'                    ElseIf ((Order.BrokerCancelOrderID <> 0) And (Order.BrokerCancelOrderID <> lCalledFromID)) And (Order.Broker = eTT_AccountType_PFG) Then
'                        Set OtherOrder = New cPtOrder
'                        If OtherOrder.Load(Order.BrokerCancelOrderID) Then
'                            If IsOpenOrder(OtherOrder.Status) Then
'                                If InfBox("You have chosen to cancel one side of an Order-Cancel-Order.  Would you like to cancel the other one as well?", "?", "+Yes|-No", "Order Cancel Order") = "N" Then
'                                    g.OrderLinks.UnlinkAndCancelOrder Order
'                                    Exit Sub
'                                End If
'                            End If
'                        End If
                    End If
'                ElseIf ((Order.IsAutoExit = True) And (Order.BrokerCancelOrderID <> 0)) And (Order.Broker = eTT_AccountType_PFG) Then
'                    Set OtherOrder = New cPtOrder
'                    If OtherOrder.Load(Order.BrokerCancelOrderID) Then
'                        If IsOpenOrder(OtherOrder.Status) Then
'                            'If InfBox("You have chosen to cancel one side of an Order-Cancel-Order.  Would you like to cancel the other one as well?", "?", "+Yes|-No", "Order Cancel Order") = "N" Then
'                                g.OrderLinks.UnlinkAndCancelOrder Order
'                                Exit Sub
'                            'End If
'                        End If
'                    End If
                End If
            End If
            
            nStatus = Order.Status
            Order.ChangeOrderStatus eTT_OrderStatus_CancelPending
            
            Select Case Order.Broker
'                Case eTT_AccountType_CtgPfg
'                    If Not g.CtgPfg Is Nothing Then
'                        If g.CtgPfg.ConnectionStatus = eGDConnectionStatus_Connected Then
'                            g.CtgPfg.CancelOrder Order
'                        Else
'                            Order.ChangeOrderStatus nStatus
'                            InfBox "You cannot cancel this order because you are not currently connected to " & g.CtgPfg.BrokerName, "!", , g.CtgPfg.BrokerName & " Order"
'                        End If
'                    Else
'                        Order.ChangeOrderStatus nStatus
'                        Err.Raise vbObjectError + 1000, , "CTG3 not Initialized"
'                    End If
'
'                Case eTT_AccountType_FintecPfg
'                    If Not g.FintecPfg Is Nothing Then
'                        If g.FintecPfg.ConnectionStatus = eGDConnectionStatus_Connected Then
'                            g.FintecPfg.CancelOrder Order
'                        Else
'                            Order.ChangeOrderStatus nStatus
'                            InfBox "You cannot cancel this order because you are not currently connected to " & g.FintecPfg.BrokerName, "!", , g.FintecPfg.BrokerName & " Order"
'                        End If
'                    Else
'                        Order.ChangeOrderStatus nStatus
'                        Err.Raise vbObjectError + 1000, , "CTG3 not Initialized"
'                    End If
'
'                Case eTT_AccountType_LindWaldock
'                    If Not g.LindWaldock Is Nothing Then
'                        g.LindWaldock.CancelOrder Order
'                    Else
'                        Order.ChangeOrderStatus nStatus
'                        Err.Raise vbObjectError + 1000, , g.Broker.BrokerName(eTT_AccountType_LindWaldock) & " not Initialized"
'                    End If
'
'                Case eTT_AccountType_ManExpress
'                    If Not g.ManExpress Is Nothing Then
'                        g.ManExpress.CancelOrder Order
'                    Else
'                        Order.ChangeOrderStatus nStatus
'                        Err.Raise vbObjectError + 1000, , "Man Express not Initialized"
'                    End If
'
'                Case eTT_AccountType_PFG
'                    If Not g.PFG Is Nothing Then
'                        If g.PFG.ConnectionStatus = eGDConnectionStatus_Connected Then
'                            g.PFG.CancelOrder Order
'                        Else
'                            Order.ChangeOrderStatus nStatus
'                            InfBox "You cannot cancel this order because you are not currently connected to " & g.PFG.BrokerName, "!", , g.PFG.BrokerName & " Order"
'                        End If
'                    Else
'                        Order.ChangeOrderStatus nStatus
'                        Err.Raise vbObjectError + 1000, , "PFG not Initialized"
'                    End If
                    
                Case eTT_AccountType_TransAct
                    If Not g.Transact Is Nothing Then
                        If g.Broker.ConnectionStatusForAccount(Order.AccountID) = eGDConnectionStatus_Connected Then
                            g.Transact.CancelOrder Order
                        Else
                            Order.ChangeOrderStatus nStatus
                            InfBox "You cannot cancel this order because you are not currently connected to TransAct.", "!", , "TransAct Order"
                        End If
                    Else
                        Order.ChangeOrderStatus nStatus
                        Err.Raise vbObjectError + 1000, , "TransAct not Initialized"
                    End If
                    
                Case Else
                    Set BrokerObj = g.Broker.Broker(Order.Broker)
                    If Not BrokerObj Is Nothing Then
                        BrokerObj.CancelOrder Order
                    End If
                
            End Select
            
            ' 10/14/2011 DAJ:  As per Pete, only bring up the order journal if the user chose
            ' to cancel the order...
            If bUserCancel Then
                g.Broker.AutoJournal Order
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.CancelOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitMultipleOrders
'' Description: Allow the user to submit multiple orders
'' Inputs:      Orders to Submit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SubmitMultipleOrders(Orders As cGdTree, Optional ByVal strPrevGenesisOrderID As String = "")
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value from function call
    Dim SimTradeOrders As New cGdTree   ' Collection of orders to send to SimTrade
    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To Orders.Count
        If Orders(lIndex).Broker = eTT_AccountType_SimBroker Then
            SimTradeOrders.Add Orders(lIndex)
        Else
            SubmitOrder Orders(lIndex)
        End If
    Next lIndex
    
    If SimTradeOrders.Count > 0 Then
        If Not g.SimTradeTs Is Nothing Then
            g.SimTradeTs.AddMultipleOrders SimTradeOrders
        Else
            Err.Raise vbObjectError + 1000, , "SimTradeTs not Initialized"
        End If
    End If
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SubmitMultipleOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DefaultAccount
'' Description: Return the ID for the default Account
'' Inputs:      None
'' Returns:     Account ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DefaultAccount() As Long
On Error GoTo ErrSection:
    
    DefaultAccount = GetIniFileProperty("LastAccount", 0&, "TTSummary", g.strIniFile)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.DefaultAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextGenesisOrderID
'' Description: Determine the next Genesis Order ID based on account type
'' Inputs:      Account Number, Broker
'' Returns:     Genesis Order ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextGenesisOrderID(ByVal strAccountNumber As String, Optional ByVal nBroker As eTT_AccountType = -1&) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim BrokerObj As cBroker            ' Broker object
    
    strReturn = ""
    If nBroker = -1& Then
        nBroker = g.Broker.AccountTypeForNumber(strAccountNumber)
    End If
    If nBroker <> -1& Then
        Select Case nBroker
'            Case eTT_AccountType_CtgPfg
'                If Not g.CtgPfg Is Nothing Then
'                    strReturn = g.CtgPfg.NextGenesisID(strAccountNumber)
'                End If
'
'            Case eTT_AccountType_FintecPfg
'                If Not g.FintecPfg Is Nothing Then
'                    strReturn = g.FintecPfg.NextGenesisID(strAccountNumber)
'                End If
'
'            Case eTT_AccountType_LindWaldock
'                If Not g.LindWaldock Is Nothing Then
'                    strReturn = g.LindWaldock.NextGenesisID(strAccountNumber)
'                End If
'
'            Case eTT_AccountType_ManExpress
'                If Not g.ManExpress Is Nothing Then
'                    strReturn = g.ManExpress.NextGenesisID(strAccountNumber)
'                End If
'
'            Case eTT_AccountType_PFG
'                If Not g.PFG Is Nothing Then
'                    strReturn = g.PFG.NextGenesisID(strAccountNumber)
'                End If
                
            Case eTT_AccountType_TransAct
                If Not g.Transact Is Nothing Then
                    strReturn = g.Transact.NextGenesisID(strAccountNumber)
                End If
                
            Case Else
                Set BrokerObj = g.Broker.Broker(nBroker)
                If Not BrokerObj Is Nothing Then
                    strReturn = BrokerObj.NextGenesisID(strAccountNumber)
                End If
                            
        End Select
    End If
    
    NextGenesisOrderID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.NextGenesisOrderID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateAccountFromNumber
'' Description: Create an account from an account number (if not already exist)
'' Inputs:      Account Number, Account Type, Account Name
'' Returns:     Account ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateAccountFromNumber(ByVal strAccountNumber As String, ByVal nAccountType As eTT_AccountType, Optional ByVal strAccountName As String = "") As Long
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lAccountID As Long              ' Account ID to return
    Dim Account As cPtAccount           ' Account object
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                "WHERE [AccountNumber]='" & strAccountNumber & "';", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        lAccountID = rs!AccountID
    Else
        rs.AddNew
        rs!AccountNumber = strAccountNumber
        If Len(Trim(strAccountName)) = 0 Then
            rs!Name = strAccountNumber
        Else
            rs!Name = strAccountName
        End If
        rs!StartingBalance = 25000#
        rs!CurrentBalance = 25000#
        rs!StartingDate = Date
        rs!AccountType = nAccountType
        If Not g.Broker.IsLiveAccount(nAccountType) Then
            rs!SecTypeMask = 31
        Else
            rs!SecTypeMask = 1
        End If
        lAccountID = rs!AccountID
        rs.Update
        
        Set Account = New cPtAccount
        Account.Load lAccountID
        
        g.Broker.UpdateAccount Account
        RefreshAccountCombos
        SendAccountToOptionNav Account, False
    End If
    
    CreateAccountFromNumber = lAccountID

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.CreateAccountFromNumber"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Cancel the order and mark it as parked
'' Inputs:      Order, Called From ID, Ask user if OCO?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ParkOrder(Order As cPtOrder, Optional ByVal lCalledFromID As Long = 0&, Optional ByVal bAskUserOnOCO As Boolean = True)
On Error GoTo ErrSection:

    Dim nStatus As eTT_OrderStatus      ' Current order status
    Dim OtherOrder As cPtOrder          ' Other side of an OCO
    Dim strReturn As String             ' Return from an InfBox
    Dim BrokerObj As cBroker            ' Broker object
    Dim nBroker As eTT_AccountType      ' Broker type for the account number

    If (Order.OrderType = eTT_OrderType_Market) And (g.Broker.IsLiveAccount(Order.Broker) = True) And (Len(Order.BrokerID) > 0) Then
        InfBox "You cannot park a market order", "!", , "Order Park Error"
        
    Else
        ' DAJ 08/14/2009: If this is one side of an order-cancel-order situation, we need to see
        ' if the user wants to park the other side of the OCO as well, break the OCO link, or
        ' cancel the park altogether...
        If (Order.AutoTradeItemID = 0&) And (Order.IsAutoExit = False) Then
            If (Order.CancelOrderID <> 0&) And (Order.CancelOrderID <> lCalledFromID) Then
                Set OtherOrder = New cPtOrder
                If OtherOrder.Load(Order.CancelOrderID) Then
                    If IsOpenOrder(OtherOrder.Status) And (OtherOrder.Status <> eTT_OrderStatus_Parked) Then
                        If bAskUserOnOCO Then
                            strReturn = InfBox("You have chosen to park one side of an Order-Cancel-Order.  What would you like to do with the other one?", "?", "+Park|Break OCO|-Cancel", "Order Cancel Order")
                            Select Case strReturn
                                Case "P"
                                    ParkOrder OtherOrder, Order.OrderID
                                    
                                Case "B"
                                    Order.CancelOrderID = 0
                                
                                Case "C"
                                    Exit Sub
                            End Select
                        Else
                            ParkOrder OtherOrder, Order.OrderID
                        End If
                    End If
                End If
'            ElseIf ((Order.BrokerCancelOrderID <> 0&) And (Order.BrokerCancelOrderID <> lCalledFromID)) And (Order.Broker = eTT_AccountType_PFG) Then
'                Set OtherOrder = New cPtOrder
'                If OtherOrder.Load(Abs(Order.BrokerCancelOrderID)) Then
'                    If IsOpenOrder(OtherOrder.Status) And (OtherOrder.Status <> eTT_OrderStatus_Parked) Then
'                        If bAskUserOnOCO Then
'                            strReturn = InfBox("You have chosen to park one side of an Order-Cancel-Order.  What would you like to do with the other one?", "?", "+Park|Break OCO|-Cancel", "Order Cancel Order")
'                            Select Case strReturn
'                                Case "P"
'                                    If g.OrderLinks.UnlinkAndParkBoth(Order, OtherOrder) = False Then
'                                        ParkOrder OtherOrder, Order.OrderID
'                                    Else
'                                        Exit Sub
'                                    End If
'
'                                Case "B"
'                                    If g.OrderLinks.UnlinkAndParkOne(Order) = False Then
'                                        Order.BrokerCancelOrderID = 0&
'                                    Else
'                                        Exit Sub
'                                    End If
'
'                                Case "C"
'                                    Exit Sub
'                            End Select
'                        Else
'                            If g.OrderLinks.UnlinkAndParkBoth(Order, OtherOrder) = False Then
'                                ParkOrder OtherOrder, Order.OrderID
'                            End If
'                        End If
'                    End If
'                End If
            End If
        End If
        
        ' 05/06/2010 DAJ: If any orders are triggered by this order, set the triggered by
        ' order ID on the triggered orders negative (so that things work right when the order
        ' gets resubmitted) and also park the triggered by orders since they don't do any
        ' good unless the triggering order is working (Issue #5715)...
        SetOtosNegative Order, True
        
        If NotSent(Order.Status) Then
            If Order.OrderDate = 0 Then
                Order.OrderDate = Order.BrokerDate(CurrentTime("", Order.Symbol))
            End If
            Order.ChangeOrderStatus eTT_OrderStatus_Parked
        Else
            ' Change the order status to Park Pending and refresh...
            nStatus = Order.Status
            Order.ChangeOrderStatus eTT_OrderStatus_ParkPending
            nBroker = Order.Broker
            
            Select Case nBroker
'                Case eTT_AccountType_CtgPfg
'                    If Not g.CtgPfg Is Nothing Then
'                        g.CtgPfg.ParkOrder Order
'                    Else
'                        Order.ChangeOrderStatus nStatus
'                        Err.Raise vbObjectError + 1000, , "CTG3 not Initialized"
'                    End If
'
'                Case eTT_AccountType_FintecPfg
'                    If Not g.FintecPfg Is Nothing Then
'                        g.FintecPfg.ParkOrder Order
'                    Else
'                        Order.ChangeOrderStatus nStatus
'                        Err.Raise vbObjectError + 1000, , "Fintec not Initialized"
'                    End If
'
'                Case eTT_AccountType_LindWaldock
'                    If Not g.LindWaldock Is Nothing Then
'                        g.LindWaldock.ParkOrder Order
'                    Else
'                        Order.ChangeOrderStatus nStatus
'                        Err.Raise vbObjectError + 1000, , g.Broker.BrokerName(eTT_AccountType_LindWaldock) & " not Initialized"
'                    End If
'
'                Case eTT_AccountType_ManExpress
'                    If Not g.ManExpress Is Nothing Then
'                        g.ManExpress.ParkOrder Order
'                    Else
'                        Order.ChangeOrderStatus nStatus
'                        Err.Raise vbObjectError + 1000, , "ManExpress not Initialized"
'                    End If
'
'                Case eTT_AccountType_PFG
'                    If Not g.PFG Is Nothing Then
'                        g.PFG.ParkOrder Order
'                    Else
'                        Order.ChangeOrderStatus nStatus
'                        Err.Raise vbObjectError + 1000, , "PFG not Initialized"
'                    End If
            
                Case eTT_AccountType_TransAct
                    If Not g.Transact Is Nothing Then
                        g.Transact.ParkOrder Order
                    Else
                        Order.ChangeOrderStatus nStatus
                        Err.Raise vbObjectError + 1000, , "TransAct not Initialized"
                    End If
                    
                Case Else
                    Set BrokerObj = g.Broker.Broker(nBroker)
                    If Not BrokerObj Is Nothing Then
                        BrokerObj.ParkOrder Order
                    Else
                        Order.ChangeOrderStatus nStatus
                        Err.Raise vbObjectError + 1000, , g.Broker.BrokerName(nBroker) & " not Initialized"
                    End If
                            
            End Select
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ParkOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAccount
'' Description: Refresh the account
'' Inputs:      Account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshAccount(ByVal lAccountID As Long, Optional ByVal bRefreshCombos As Boolean = True, Optional ByVal bReloadAccount As Boolean = True)
On Error GoTo ErrSection:

    If bRefreshCombos Then
        RefreshAccountCombos
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.RefreshAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConvertBrokerDate
'' Description: Convert an order or fill date from an online broker to local
'' Inputs:      Date, Broker, Symbol, ToLocal?, Time Zone Info
'' Returns:     Local Date/Time for Date passed in
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConvertBrokerDate(ByVal dDateTime As Double, ByVal nBroker As eTT_AccountType, ByVal strSymbol As String, Optional ByVal bToLocal As Boolean = True, Optional ByVal strExchTime As String = "") As Double
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary bars object
    Dim strChangeTo As String           ' Time zone to change to
    Dim BrokerObj As cBroker            ' Broker object
    Dim dReturn As Double               ' Return value for the function
    
    If Len(strExchTime) = 0 Then
        SetBarProperties Bars, strSymbol
        strExchTime = Bars.Prop(eBARS_ExchangeTimeZoneInf)
    End If
    
    If bToLocal = True Then
        strChangeTo = ""
    Else
        strChangeTo = strExchTime
    End If
    
    Select Case nBroker
        Case eTT_AccountType_AdvFut
            dReturn = ConvertTimeZone(dDateTime, "CHI", strChangeTo)
        
'        Case eTT_AccountType_LindWaldock, eTT_AccountType_ManExpress
'            dReturn = ConvertTimeZone(dDateTime, "CHI", strChangeTo)

'        Case eTT_AccountType_CtgPfg, eTT_AccountType_FintecPfg, eTT_AccountType_PFG
'            dReturn = ConvertTimeZone(dDateTime, "CHI", strChangeTo)
        
        Case eTT_AccountType_TransAct
            dReturn = ConvertTimeZone(dDateTime, "GMT", strChangeTo)
            
        Case eTT_AccountType_TT
            dReturn = ConvertTimeZone(dDateTime, "GMT", strChangeTo)
        
        Case Else
            Set BrokerObj = g.Broker.Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                dReturn = ConvertTimeZone(dDateTime, BrokerObj.TimeZone(strSymbol), strChangeTo)
            Else
                dReturn = dDateTime
            End If
    
    End Select
    
    ConvertBrokerDate = dReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ConvertBrokerDate"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConvertToBrokerDate
'' Description: Convert a local date to an online broker date
'' Inputs:      Date, Broker, Symbol, From Local?
'' Returns:     Broker Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConvertToBrokerDate(ByVal dDateTime As Double, ByVal nBroker As eTT_AccountType, ByVal strSymbol As String, Optional ByVal bFromLocal As Boolean = True) As Double
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary bars object
    Dim strChangeFrom As String         ' Time zone to change to
    Dim strExchTime As String           ' Exchange time zone information
    Dim BrokerObj As cBroker            ' Broker object
    Dim dReturn As Double               ' Return value for the function
    
    SetBarProperties Bars, strSymbol
    
    strExchTime = Bars.Prop(eBARS_ExchangeTimeZoneInf)
    If bFromLocal = True Then
        strChangeFrom = ""
    Else
        strChangeFrom = strExchTime
    End If
    
    Select Case nBroker
        Case eTT_AccountType_AdvFut
            dReturn = ConvertTimeZone(dDateTime, strChangeFrom, "CHI")
        
'        Case eTT_AccountType_LindWaldock, eTT_AccountType_ManExpress
'            dReturn = ConvertTimeZone(dDateTime, strChangeFrom, "CHI")
        
'        Case eTT_AccountType_CtgPfg, eTT_AccountType_FintecPfg, eTT_AccountType_PFG
'            dReturn = ConvertTimeZone(dDateTime, strChangeFrom, "CHI")
            
        Case eTT_AccountType_TransAct
            dReturn = ConvertTimeZone(dDateTime, strChangeFrom, "GMT")
        
        Case eTT_AccountType_TT
            dReturn = ConvertTimeZone(dDateTime, strChangeFrom, "GMT")
            
        Case Else
            Set BrokerObj = g.Broker.Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                dReturn = ConvertTimeZone(dDateTime, strChangeFrom, BrokerObj.TimeZone(strSymbol))
            Else
                dReturn = dDateTime
            End If
    
    End Select
    
    ConvertToBrokerDate = dReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ConvertToBrokerDate"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDefaultEntryForSymbol
'' Description: Get the default entry quantity for the symbol
'' Inputs:      Symbol ID, Symbol
'' Returns:     Entry Quantity if Found, Zero if not
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDefaultEntryForSymbol(ByVal lSymbolID As Long, ByVal strSymbol As String) As Long
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' Default Entry lookup file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value from the function
    Dim lPos As Long                    ' Position in the array
    Dim strKey As String                ' Key to lookup in the string
    
    lReturn = 0&
    astrFile.FromFile AddSlash(App.Path) & "EntryQty.CFG"
    If astrFile.Size > 0 Then
        astrFile.Sort
        If lSymbolID = 0 Then strKey = strSymbol Else strKey = Str(lSymbolID)
        If astrFile.BinarySearch(strKey & "|", lPos, eGdSort_MatchUsingSearchStringLength) Then
            lReturn = CLng(Val(Parse(astrFile(lPos), "|", 2)))
        End If
    End If
    
    GetDefaultEntryForSymbol = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.GetDefaultEntryForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetDefaultEntryForSymbol
'' Description: Set the default entry quantity for the symbol
'' Inputs:      Symbol ID, Symbol, Quantity
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetDefaultEntryForSymbol(ByVal lSymbolID As Long, ByVal strSymbol As String, ByVal lQuantity As Long)
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' Default Entry lookup file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position in the array
    Dim strKey As String                ' Key to lookup in the string

    astrFile.FromFile AddSlash(App.Path) & "EntryQty.CFG"
    If lSymbolID = 0 Then strKey = strSymbol Else strKey = Str(lSymbolID)
    
    If astrFile.BinarySearch(strKey & "|", lPos, eGdSort_MatchUsingSearchStringLength) Then
        astrFile(lPos) = strKey & "|" & Str(lQuantity)
    Else
        astrFile.Add strKey & "|" & Str(lQuantity), lPos
    End If
    
    astrFile.ToFile AddSlash(App.Path) & "EntryQty.CFG"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SetDefaultEntryForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderCallback
'' Description: Do things necessary after getting an order status change
'' Inputs:      Order, Refresh Charts?, Order Was Parked?, Check Trigger Order?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderCallback(Order As cPtOrder, Optional ByVal bRefreshCharts As Boolean = True, Optional ByVal bWasParked As Boolean = False, Optional ByVal bCheckTrigger As Boolean = True)
On Error GoTo ErrSection:

    Dim TriggerOrder As cPtOrder        ' Trigger order
    Dim OtherOrder As cPtOrder          ' Other order

gdStartProfile 903

    RefreshOrder Order, , bRefreshCharts, bWasParked
    g.TradingItems.OrderCallback Order.AutoTradeItemID, Order
    g.OrderStrategies.OrderCallback Order
    
gdStopProfile 903
gdStartProfile 904
    g.Alerts.CheckAlerts
gdStopProfile 904
gdStartProfile 905
    If IsOpenOrder(Order.Status) = False Then g.Alerts.RemoveAlertsForManualOrder Order.OrderID
gdStopProfile 905

    If bCheckTrigger Then
        If Order.TriggerOrderID <> 0& Then
            Set TriggerOrder = New cPtOrder
            
            ' DAJ 11/28/2014: Customer is running into an issue with OrderCallback being called
            ' a second time and getting a 'related record is required in table tblAccounts' error.
            ' This is the only way I can see where OrderCallback can be called back-to-back and the
            ' only way that error could occur is if the trigger order didn't get loaded...
            ' DAJ 12/02/2014: Found out what caused the error was when the TriggerOrderID was
            ' negative ( which happens when the triggered order goes parked ).  Given that knowledge,
            ' we can load up the order with the absolute value here...
            If TriggerOrder.Load(Abs(Order.TriggerOrderID)) = True Then
                TriggerOrder.UpdateContingentOrder Order
                OrderCallback TriggerOrder
            Else
                g.Broker.BrokerDebug Order.Broker, "Could not load trigger order " & Str(Order.TriggerOrderID)
            End If
        End If
    End If
    
    If (Order.Status = eTT_OrderStatus_Working) And (Order.CancelOrderID > 0&) Then
        Set OtherOrder = New cPtOrder
        If OtherOrder.Load(Order.CancelOrderID) Then
            If (OtherOrder.Status = eTT_OrderStatus_Filled) Or (OtherOrder.Status = eTT_OrderStatus_Partial) Then
                g.Broker.BrokerDebug Order.Broker, "Cancelling Order '" & Order.OrderText(True, True, True) & "' because of an OCO (" & Str(OtherOrder.OrderID) & ", '" & OtherOrder.GenesisOrderID & "', '" & OtherOrder.BrokerID & "')", True
                CancelOrder Order, False
            End If
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mTradeTracker.OrderCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillCallback
'' Description: Do things necessary  after getting a fill for an order
'' Inputs:      Order, Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FillCallback(Order As cPtOrder, Fill As cPtFill, Optional ByVal bFillExisted As Boolean = False, Optional ByVal bRefreshCharts As Boolean = True, Optional ByVal bFillChanged As Boolean = True)
On Error GoTo ErrSection:

    Dim Order2 As New cPtOrder          ' Order to cancel if applicable
    Dim rs As Recordset                 ' Recordset into the database
    Dim strReturn As String             ' Return from an InfBox
    Dim nAcctType As eTT_AccountType    ' Account type

    nAcctType = g.Broker.AccountTypeForID(Order.AccountID)
    
    ' Need this block to happen before RefreshOrder gets called...
    If bFillExisted = False Then
        If Not g.Alerts Is Nothing Then
            g.Alerts.OrderStatusChange Order
        End If
        If Not g.TsoGroups Is Nothing Then
            g.TsoGroups.FillCallback Fill, Order
        End If
    
        ' If the user has auto journal turned on and we get a new manual or auto exit fill, create
        ' an automatic date journal entry for the user...
        If g.Broker.AutoJournalPopUp Then
            If (Order.IsAutomated = False) Or (Order.IsAutoExit = True) Then
                ' If there is no symbol ID, do the journal now since there won't be a chart...
                If Fill.SymbolID = 0& Then
                    g.TnJournal.AutoJournalForFill Fill
                    
                ' Otherwise, queue it up so that we create the journal entry after the chart
                ' is updated...
                Else
                    frmOnlineBroker.FillsToJournal.Add Fill
                End If
            End If
        End If
    End If
    
    RefreshOrder Order, , bRefreshCharts
    
    If bFillExisted = False Then
        g.TradingItems.OrderCallback Order.AutoTradeItemID, Order
        'g.TradingItems.FillCallback Order.AutoTradeItemID, Fill
        g.TradingItems.AddFillCheck Fill, Order
        
        Select Case nAcctType
'            Case eTT_AccountType_CtgPfg, eTT_AccountType_FintecPfg, eTT_AccountType_PFG, eTT_AccountType_TransAct, eTT_AccountType_LindWaldock, eTT_AccountType_ManExpress
            Case eTT_AccountType_TransAct
                g.OrderStrategies.OrderCallback Order
            
            Case Else
                g.OrderStrategies.FillCallback Fill, Order
        End Select
    End If
    
    ' If this is part of an OCO, Cancel the other order...
    If Order.CancelOrderID <> 0 Then
        If Order2.Load(Order.CancelOrderID) Then
            If IsOpenOrder(Order2.Status, False) And (Order2.ExitPos = 0) Then
                g.Broker.BrokerDebug Order2.Broker, "Cancelling Order '" & Order2.OrderText(True, True, True) & "' because of an OCO (" & Str(Order.OrderID) & ", '" & Order.GenesisOrderID & "', '" & Order.BrokerID & "')", True
                CancelOrder Order2, False
            ElseIf OrderIsPending(Order2) Then
                g.Broker.BrokerDebug Order2.Broker, "Not Cancelling Order '" & Order2.OrderText(True, True, True) & "' because of an OCO (" & Str(Order.OrderID) & ", '" & Order.GenesisOrderID & "', '" & Order.BrokerID & "') because the order is pending", True
            End If
        End If
        
    ' DAJ 07/22/2010: Also need to handle a Broker OCO that hasn't been confirmed yet (such as one or both sides
    ' is a conditional order that hasn't triggered yet)...
    ElseIf Order.BrokerCancelOrderID < 0 Then
        If Order2.Load(Abs(Order.BrokerCancelOrderID)) Then
            If IsOpenOrder(Order2.Status, False) And (Order2.ExitPos = 0) Then
                g.Broker.BrokerDebug Order2.Broker, "Cancelling Order '" & Order2.OrderText(True, True, True) & "' because of an OCO (" & Str(Order.OrderID) & ", '" & Order.GenesisOrderID & "', '" & Order.BrokerID & "')", True
                CancelOrder Order2, False
            ElseIf OrderIsPending(Order2) Then
                g.Broker.BrokerDebug Order2.Broker, "Not Cancelling Order '" & Order2.OrderText(True, True, True) & "' because of an OCO (" & Str(Order.OrderID) & ", '" & Order.GenesisOrderID & "', '" & Order.BrokerID & "') because the order is pending", True
            End If
        End If
    End If
        
    ' If this order is supposed to trigger other orders, do it...
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
            "WHERE [TriggerOrderID]=" & Str(Order.OrderID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        g.ExitAllOrders.CheckOrders
    End If
    Do While Not rs.EOF
        Set Order2 = New cPtOrder
        If Order2.Load(rs!OrderID) Then
            'If ((Order.Status = eTT_OrderStatus_Filled) Or (Order2.TriggerOnPartial = True)) And (IsOpenOrder(Order2.Status) = True) And (OrderIsPending(Order2) = False) Then
            If ((Order.Status = eTT_OrderStatus_Filled) Or (Order2.TriggerOnPartial = True)) And (Order2.Status = eTT_OrderStatus_TriggerPending) Then
                If Order2.IsConditional(False) = False Then
                    strReturn = "S"
                    If g.Broker.ConfirmTriggered Then
                        strReturn = ConfirmOrder(Order2)
                    End If
                    Select Case strReturn
                        Case "S"
                            If Order2.Quantity > 0 Then
                                If Len(Order2.GenesisOrderID) = 0 Then
                                    Order2.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Order2.AccountID))
                                End If
                                If Len(Parse(Order2.TriggerOptions, ",", 3)) > 0 Then
                                    Select Case Order2.OrderType
                                        Case eTT_OrderType_Stop
                                            If Order2.Buy Then
                                                Order2.StopPrice = Order.AvgFillPrice + Val(Parse(Order2.TriggerOptions, ",", 3))
                                            Else
                                                Order2.StopPrice = Order.AvgFillPrice - Val(Parse(Order2.TriggerOptions, ",", 3))
                                            End If
                                            Order2.TriggerOptions = "1,0"
                                            
                                        Case eTT_OrderType_Limit
                                            If Order2.Buy Then
                                                Order2.LimitPrice = Order.AvgFillPrice - Val(Parse(Order2.TriggerOptions, ",", 3))
                                            Else
                                                Order2.LimitPrice = Order.AvgFillPrice + Val(Parse(Order2.TriggerOptions, ",", 3))
                                            End If
                                            Order2.TriggerOptions = "1,0"
                                            
                                        Case eTT_OrderType_MIT
                                            If Order2.Buy Then
                                                Order2.MitPrice = Order.AvgFillPrice - Val(Parse(Order2.TriggerOptions, ",", 3))
                                            Else
                                                Order2.MitPrice = Order.AvgFillPrice + Val(Parse(Order2.TriggerOptions, ",", 3))
                                            End If
                                            Order2.TriggerOptions = "1,0"
                                    
                                    End Select
                                End If
                                
                                ' 05/24/2010 DAJ: Clear the trigger order stuff here so that it is never
                                ' considered a trigger-by order again...
                                g.Broker.BrokerDebug Order2.Broker, "Trigger Information: " & Str(Order2.TriggerOrderID) & " (" & Order2.TriggerOptions & ") being cleared out because order " & Str(Order.OrderID) & " at least partially filled"
                                Order2.TriggerOrderID = 0&
                                Order2.TriggerOptions = ""
                                Order2.Save
                                
                                SubmitOrder Order2
                            Else
                                g.Broker.BrokerDebug g.Broker.AccountTypeForID(Order2.AccountID), "Cancelling Triggered Order " & Order2.OrderText & " (" & Order2.BrokerID & ") because quantity is zero", True
                                CancelOrder Order2, False
                            End If
                        
                        Case "P"
                            g.Broker.BrokerDebug g.Broker.AccountTypeForID(Order2.AccountID), "Parking Triggered Order " & Order2.OrderText & " (" & Order2.BrokerID & ") because user parked order", True
                            Order2.ChangeOrderStatus eTT_OrderStatus_Parked
                        
                        Case "C"
                            g.Broker.BrokerDebug g.Broker.AccountTypeForID(Order2.AccountID), "Cancelling Triggered Order " & Order2.OrderText & " (" & Order2.BrokerID & ") because user cancelled order", True
                            CancelOrder Order2, False
                            
                    End Select
                End If
            End If
        End If
        
        rs.MoveNext
    Loop
    
    g.CattleBridge.Broker_Fill g.TnCattle.GenesisFillToTurnkey(Fill)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.FillCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConfirmOrder
'' Description: Confirm with the user whether to submit the order or not
'' Inputs:      Order, Reason
'' Returns:     (S)ubmit, (P)ark, (C)ancel
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConfirmOrder(ByVal Order As cPtOrder, Optional ByVal strReason As String = "", Optional ByVal strDefault As String = "S") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return from the InfBox
    Dim strButtons As String            ' Button information string
    
    Select Case Left(UCase(strDefault), 1)
        Case "S"
            strButtons = "+Submit Order|Park Order|-Cancel Order"
        Case "P"
            strButtons = "Submit Order|+Park Order|-Cancel Order"
        Case "C"
            strButtons = "Submit Order|Park Order|+-Cancel Order"
    End Select
    
    If Len(strReason) = 0 Then
        strReturn = InfBox("Do you want to submit the following order?|" & Order.OrderText & "|", "?", strButtons, "Order Confirmation", , 20&)
        g.Broker.BrokerDebug Order.Broker, "User Asked: 'Do you want to submit the following order? " & Order.OrderText & "'; Response = " & strReturn
    Else
        strReturn = InfBox(strReason & "|Do you want to submit the following order?|" & Order.OrderText & "|", "?", strButtons, "Order Confirmation", , 20&)
        g.Broker.BrokerDebug Order.Broker, "User Asked: '" & strReason & " Do you want to submit the following order? " & Order.OrderText & "'; Response = " & strReturn
    End If
    
    ConfirmOrder = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ConfirmOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIsEntry
'' Description: Determine whether the given order will be an entry or an exit
'' Inputs:      Order, Use Category?
'' Returns:     True if Entry, False if Exit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderIsEntry(ByVal Order As cPtOrder, Optional ByVal bUseCategory As Boolean = True, Optional bReverse As Boolean, Optional ByVal bDumpToLog As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lCurrentPos As Long             ' Current position
    Dim lTriggerPos As Long             ' Trigger position adjustment
    Dim lTriggerOrderID As Long         ' Triggered-by order ID
    Dim Trigger As New cPtOrder         ' Triggering order if applicable
    
    If bUseCategory Then
        lCurrentPos = g.Broker.CurrentPosition(Order.AccountID, Order.Symbol, Order.AutoTradeItemID)
    Else
        lCurrentPos = g.Broker.CurrentPosition(Order.AccountID, Order.Symbol, -1&)
    End If
    
    ' Adjust for triggering order if applicable...
    lTriggerPos = 0&
    lTriggerOrderID = Order.TriggerOrderID
    Do While lTriggerOrderID <> 0&
        If Trigger.Load(lTriggerOrderID) Then
            If Order.SymbolOrSymbolID = Trigger.SymbolOrSymbolID Then
                If Trigger.Buy Then
                    lTriggerPos = lTriggerPos + Trigger.Quantity
                Else
                    lTriggerPos = lTriggerPos - Trigger.Quantity
                End If
            End If
            
            lTriggerOrderID = Trigger.TriggerOrderID
        Else
            lTriggerOrderID = 0&
        End If
    Loop
    
    If Order.Buy Then
        If ((lCurrentPos + Order.Quantity + lTriggerPos) > 0) Then
            bReturn = True
            bReverse = (lCurrentPos < 0)
        Else
            bReturn = False
            bReverse = False
        End If
    Else
        If ((lCurrentPos - Order.Quantity + lTriggerPos) < 0) Then
            bReturn = True
            bReverse = (lCurrentPos > 0)
        Else
            bReturn = False
            bReverse = False
        End If
    End If
    
    If bDumpToLog Then
        If Order.Buy Then
            g.Broker.BrokerDebug Order.Broker, vbTab & "Order Is Entry = " & Str(bReturn) & " ( Order Quantity = " & Str(Order.Quantity) & "; Current = " & Str(lCurrentPos) & "; Trigger = " & Str(lTriggerPos) & " )"
        Else
            g.Broker.BrokerDebug Order.Broker, vbTab & "Order Is Entry = " & Str(bReturn) & " ( Order Quantity = " & Str(Order.Quantity * -1#) & "; Current = " & Str(lCurrentPos) & "; Trigger = " & Str(lTriggerPos) & " )"
        End If
    End If
    
    OrderIsEntry = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.OrderIsEntry"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderLegIsEntry
'' Description: Determine whether the given order leg will be an entry or an exit
'' Inputs:      Order, Leg, Use Category?, (out)Reverse?
'' Returns:     True if Entry, False if Exit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderLegIsEntry(ByVal Order As cPtOrder, ByVal lOrderLeg As Long, Optional ByVal bUseCategory As Boolean = True, Optional bReverse As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim OrderLeg As cOrderLeg           ' Order leg that we are working with
    Dim lCurrentPos As Long             ' Current position
    Dim lTriggerOrderID As Long         ' Triggered-by order ID
    Dim Trigger As New cPtOrder         ' Triggering order if applicable
    Dim TriggerLeg As cOrderLeg         ' Order leg on the trigerring order
    Dim lIndex As Long                  ' Index into a for loop
    
    Set OrderLeg = Order.OrderLegs(lOrderLeg)
    
    If bUseCategory Then
        lCurrentPos = g.Broker.CurrentPosition(Order.AccountID, OrderLeg.Symbol, Order.AutoTradeItemID)
    Else
        lCurrentPos = g.Broker.CurrentPosition(Order.AccountID, OrderLeg.Symbol, -1&)
    End If
    
    ' Adjust for triggering order if applicable...
    lTriggerOrderID = Order.TriggerOrderID
    Do While lTriggerOrderID <> 0&
        If Trigger.Load(lTriggerOrderID) Then
            For lIndex = 1 To Trigger.NumberOfLegs
                Set TriggerLeg = Trigger.OrderLegs(lIndex)
                If OrderLeg.SymbolOrSymbolID = TriggerLeg.SymbolOrSymbolID Then
                    If TriggerLeg.IsBuy Then
                        lCurrentPos = lCurrentPos + (Trigger.Quantity * TriggerLeg.Multiplier)
                    Else
                        lCurrentPos = lCurrentPos - (Trigger.Quantity * TriggerLeg.Multiplier)
                    End If
                End If
            Next lIndex
            
            lTriggerOrderID = Trigger.TriggerOrderID
        Else
            lTriggerOrderID = 0&
        End If
    Loop
    
    If OrderLeg.IsBuy Then
        'g.Broker.BrokerDebug Order.Broker, Str(lCurrentPos) & " + ( " & Str(Order.Quantity) & " * " & Str(OrderLeg.Multiplier) & " ) = " & Str(lCurrentPos + (Order.Quantity * OrderLeg.Multiplier))
        
        If ((lCurrentPos + (Order.Quantity * OrderLeg.Multiplier)) > 0) Then
            bReturn = True
            bReverse = (lCurrentPos < 0)
        Else
            bReturn = False
            bReverse = False
        End If
    Else
        'g.Broker.BrokerDebug Order.Broker, Str(lCurrentPos) & " - ( " & Str(Order.Quantity) & " * " & Str(OrderLeg.Multiplier) & " ) = " & Str(lCurrentPos - (Order.Quantity * OrderLeg.Multiplier))
        
        If ((lCurrentPos - (Order.Quantity * OrderLeg.Multiplier)) < 0) Then
            bReturn = True
            bReverse = (lCurrentPos > 0)
        Else
            bReturn = False
            bReverse = False
        End If
    End If
    
    OrderLegIsEntry = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.OrderLegIsEntry"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckTriggerByOrders
'' Description: Check to see if we need to park the triggered by orders
'' Inputs:      Triggering Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckTriggerByOrders(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Order2 As New cPtOrder          ' Triggered-by order
    Dim lIndex As Long                  ' Index into a for loop
    
    Select Case Order.Status
        Case eTT_OrderStatus_Cancelled, eTT_OrderStatus_Rejected, eTT_OrderStatus_BalCancelled, eTT_OrderStatus_Error, eTT_OrderStatus_Expired
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [TriggerOrderID]=" & Str(Order.OrderID) & ";", dbOpenDynaset)
            Do While Not rs.EOF
                If Order2.Load(rs!OrderID, rs) Then
                    If (IsOpenOrder(Order2.Status, False) = True) Or (Order2.Status = eTT_OrderStatus_TriggerPending) Then
                        g.Broker.BrokerDebug Order2.Broker, "Cancelling Order '" & Order2.OrderText(True, True, True) & "' because Triggered by Order has been cancelled (" & Str(Order.OrderID) & ", '" & Order.GenesisOrderID & "', '" & Order.BrokerID & "')"
                        CancelOrder Order2, False
                    End If
                End If
                
                rs.MoveNext
            Loop
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.CheckTriggerByOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExpireNonSubmittedOrders
'' Description: Expire any non-submitted orders where the expiration date for
''              the order is less than or equal to the last download date
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExpireNonSubmittedOrders()
On Error GoTo ErrSection:

    Dim lLastDownload As Long           ' Last download date
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As New cPtOrder           ' Temporary Order object
    Dim lExpirationDate As Long         ' Expiration date to check
    
    lLastDownload = LastDailyDownload
    With frmTTSummary.fgOrders
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If TypeOf .RowData(lIndex) Is cPtOrder Then
                Set Order = .RowData(lIndex)
                If NotSent(Order.Status) Then
                    If Order.Expiration = -1& Then
                        lExpirationDate = Order.SessionDate
                    Else
                        lExpirationDate = Abs(Order.Expiration)
                    End If

                    If (lExpirationDate <= lLastDownload) And (lExpirationDate <> 0) Then
                        Order.ChangeOrderStatus eTT_OrderStatus_Expired
                    End If
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ExpireNonSubmittedOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AdjustTriggeredOrders
'' Description: Adjust triggered orders with the same symbol by the same amount
''              as the trigerring order if the user so desires
'' Inputs:      Triggering Order, Adjust Amount
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AdjustTriggeredOrders(ByVal Order As cPtOrder, ByVal dAdjustAmount As Double, Optional ByVal bAsk As Boolean = True)
On Error GoTo ErrSection:

    Dim TrigOrders As cGdTree           ' Collection of orders triggered by the passed in order
    Dim TrigOrder As New cPtOrder       ' Triggering order
    Dim lIndex As Long                  ' Index into a for loop
    Dim strReturn As String             ' Return value from an InfBox
    Dim bAllMarket As Boolean           ' All triggered orders are market orders
    
    Set TrigOrders = g.Broker.TriggeredOrdersForOrder(Order)
    If TrigOrders.Count > 0 Then
        ' Remove the contingency orders because they will be adjusted elsewhere...
        For lIndex = TrigOrders.Count To 1 Step -1
            Set TrigOrder = TrigOrders(lIndex)
            If (TrigOrder.OrderID = Order.Contingency.ProfitOrderId) Or (TrigOrder.OrderID = Order.Contingency.StopOrderId) Then
                TrigOrders.Remove lIndex
            End If
        Next lIndex
        
        bAllMarket = True
        For lIndex = 1 To TrigOrders.Count
            If TrigOrders(lIndex).OrderType <> eTT_OrderType_Market Then
                bAllMarket = False
            End If
        Next lIndex
    
        If bAllMarket = False Then
            For lIndex = 1 To TrigOrders.Count
                Set TrigOrder = TrigOrders(lIndex)
                    
                If TrigOrder.MoveWithTrigger = True Then
                    Select Case TrigOrder.OrderType
                        Case eTT_OrderType_Stop, eTT_OrderType_StopCloseOnly
                            g.Broker.BrokerDebug TrigOrder.Broker, vbTab & "Changing Stop Price on Order " & Str(TrigOrder.OrderID) & " from " & Str(TrigOrder.StopPrice) & " to " & Str(TrigOrder.StopPrice + dAdjustAmount)
                            TrigOrder.StopPrice = TrigOrder.StopPrice + dAdjustAmount
                            TrigOrder.Save
                            g.Broker.AddOrder TrigOrder
                            OrderCallback TrigOrder
                            
                        Case eTT_OrderType_Limit, eTT_OrderType_LimitCloseOnly
                            g.Broker.BrokerDebug TrigOrder.Broker, vbTab & "Changing Limit Price on Order " & Str(TrigOrder.OrderID) & " from " & Str(TrigOrder.LimitPrice) & " to " & Str(TrigOrder.LimitPrice + dAdjustAmount)
                            TrigOrder.LimitPrice = TrigOrder.LimitPrice + dAdjustAmount
                            TrigOrder.Save
                            g.Broker.AddOrder TrigOrder
                            OrderCallback TrigOrder
                            
                        Case eTT_OrderType_StopWithLimit, eTT_OrderType_StopWithLimitCloseOnly
                            g.Broker.BrokerDebug TrigOrder.Broker, vbTab & "Changing Stop Price on Order " & Str(TrigOrder.OrderID) & " from " & Str(TrigOrder.StopPrice) & " to " & Str(TrigOrder.StopPrice + dAdjustAmount)
                            g.Broker.BrokerDebug TrigOrder.Broker, vbTab & "Changing Limit Price on Order " & Str(TrigOrder.OrderID) & " from " & Str(TrigOrder.LimitPrice) & " to " & Str(TrigOrder.LimitPrice + dAdjustAmount)
                            TrigOrder.StopPrice = TrigOrder.StopPrice + dAdjustAmount
                            TrigOrder.LimitPrice = TrigOrder.LimitPrice + dAdjustAmount
                            TrigOrder.Save
                            g.Broker.AddOrder TrigOrder
                            OrderCallback TrigOrder
                            
                        Case eTT_OrderType_MIT
                            g.Broker.BrokerDebug TrigOrder.Broker, vbTab & "Changing MIT Price on Order " & Str(TrigOrder.OrderID) & " from " & Str(TrigOrder.MitPrice) & " to " & Str(TrigOrder.MitPrice + dAdjustAmount)
                            TrigOrder.MitPrice = TrigOrder.MitPrice + dAdjustAmount
                            TrigOrder.Save
                            g.Broker.AddOrder TrigOrder
                            OrderCallback TrigOrder
                                                        
                    End Select
                    
                    AdjustTriggeredOrders TrigOrder, dAdjustAmount, False
                End If
            Next lIndex
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.AdjustTriggeredOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetDirtyChartTrades
'' Description: Set the dirty trades flag on any chart with the given symbol
''              and account so that the chart knows to reload the trades
'' Inputs:      Symbol or Symbol ID, Account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetDirtyChartTrades(ByVal vSymbolOrSymbolID As Variant, ByVal lAccountID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim frm As Form                     ' Temporary Chart form object
    Dim lSymbolID As Long

    lSymbolID = GetSymbolID(ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True))))
    If lSymbolID <> 0 And lAccountID <> 0 Then
        For lIndex = 0 To Forms.Count - 1
            If g.bUnloading Then
                Exit For
            Else
                If IsFrmChart(Forms(lIndex)) Then
                    Set frm = Forms(lIndex)
                    If (frm.Chart.TradeAccountID = lAccountID) Then
                        If ConvertToTradeSymbol(frm.Chart.SymbolID, Int(CurrentTime("", "", True))) = lSymbolID Then
                            frm.Chart.SetTrackerTradesReload
                        End If
                    End If
                End If
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SetDirtyChartTrades"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ModifyExitPosOrders
'' Description: Modify any exit position orders that exist for the given fill
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ModifyExitPosOrders(ByVal Fill As cPtFill)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Order As New cPtOrder           ' Temporary order object
    Dim lAdjust As Long                 ' Adjustment to the order quantity
    
    If Fill.SymbolID = 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                    "WHERE [Symbol]='" & Fill.Symbol & "' AND [AccountID]=" & Str(Fill.AccountID) & " AND [AutoTradeItemID]=" & Str(Fill.AutoTradingItemID) & " " & _
                    "AND [ExitPos]>0;", dbOpenDynaset)
    Else
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                    "WHERE [SymbolID]=" & Str(Fill.SymbolID) & " AND [AccountID]=" & Str(Fill.AccountID) & " AND [AutoTradeItemID]=" & Str(Fill.AutoTradingItemID) & " " & _
                    "AND [ExitPos]>0;", dbOpenDynaset)
    End If
    
    Do While Not rs.EOF
        If IsOpenOrder(rs!Status) Then
            Set Order = New cPtOrder
            If Order.Load(rs!OrderID) Then
                lAdjust = 0&
                
                If (Order.Buy <> Fill.Buy) Then
                    lAdjust = Fill.Quantity
                ElseIf Order.OrderID <> Fill.OrderID Then
                    lAdjust = -Fill.Quantity
                End If
                
                If lAdjust <> 0& Then
                    If NotSent(Order.Status) Then
                        Order.Quantity = Order.Quantity + lAdjust
                        Order.Save
                        OrderCallback Order
                    ElseIf Order.Quantity + lAdjust > 0 Then
                        '*** PRICE OR QUANTITY ON ORDER CHANGED...***
                        Order.Quantity = Order.Quantity + lAdjust
                        Order.Save
                        SubmitOrder Order, Order.GenesisOrderID
                    Else
                        CancelOrder Order, False
                    End If
                End If
            End If
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ModifyExitPosOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillOrdersRT
'' Description: Do we fill orders in demo real-time mode for the given account?
'' Inputs:      Account ID
'' Returns:     True if Fill RT, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillOrdersRT(ByVal lAccountID As Long) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Acct As cPtAccount              ' Account object
    Dim bReturn As Boolean              ' Return value for the function
    
    Set Acct = Nothing
    bReturn = False
    
    Set Acct = g.Broker.Account(lAccountID)
    If Not Acct Is Nothing Then
        bReturn = Acct.FillRT
    Else
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            If Not g.Broker.IsLiveAccount(rs!AccountType) Then
                bReturn = rs!FillRT
            End If
        End If
    End If
    
    FillOrdersRT = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.FillOrdersRT"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleDemoOrders
'' Description: If there are currently demo orders, then ask if the user wants
''              to send them to the Trade Server or cancel them
'' Inputs:      Message Timeout
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleDemoOrders(Optional ByVal lMsgTimeout As Long = 0&)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lTimeOut As Long                ' Timeout variable
    
    ' Need to cancel any working "quick fill" orders because streaming is going away...
    If g.SimTradeStream.Broker.BrokerInfo.HasWorkingOrders(False, True, False) Then
        If InfBox("Would you like to Park or Cancel all of your working simulated orders?", "?", "+Park|-Cancel", "Simluated Orders", , lMsgTimeout) = "C" Then
            Do While g.SimTradeStream.Broker.BrokerInfo.HasWorkingOrders(False, True, False) And (lTimeOut < 30&)
                g.SimTradeStream.CancelAllWorkingOrders False, False
            
                Sleep 1, False, True
                lTimeOut = lTimeOut + 1&
            Loop
        Else
            Do While g.SimTradeStream.Broker.BrokerInfo.HasWorkingOrders(False, True, False) And (lTimeOut < 30&)
                g.SimTradeStream.ParkAllWorkingOrders False
            
                Sleep 1, False, True
                lTimeOut = lTimeOut + 1&
            Loop
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.HandleDemoOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetupInitialAccounts
'' Description: Set up initial SimTrade and broker accounts if not already done
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetupInitialAccounts()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim lAccountID As Long              ' Account ID of the newly created account
    Dim bChanged As Boolean             ' Did we change the account?
    Dim Account As cPtAccount           ' Account object
    
    ' If no SimTrade accounts exist yet, create one...
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                "WHERE [AccountType]=" & Str(eTT_AccountType_SimStream) & " OR [AccountType]=" & Str(eTT_AccountType_SimBroker) & ";", dbOpenDynaset)
    If rs.BOF And rs.EOF Then
        If HasModule("RTG") = True Then
            CreateAccountFromNumber "SIM0001", eTT_AccountType_SimStream
        Else
            CreateAccountFromNumber "GEN0001", eTT_AccountType_SimBroker
        End If
        
    ' The default database going out to customers already has one account in it.  If the user has
    ' a brand new database with one account in it, make sure it is the right style of simulated
    ' ( SimStream if they have streaming, SimBroker otherwise )...
    ElseIf (rs.RecordCount = 1) And (rs!AccountID = 1) Then
        bChanged = False
        Set rs2 = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] WHERE [AccountID]=" & rs!AccountID & ";", dbOpenDynaset)
        If rs2.BOF And rs2.EOF Then
            If (rs!AccountType = eTT_AccountType_SimBroker) And (HasModule("RTG") = True) Then
                rs.Edit
                rs!AccountType = eTT_AccountType_SimStream
                rs!Name = "Simulated"
                rs.Update
                
                bChanged = True
            ElseIf (rs!AccountType = eTT_AccountType_SimStream) And (HasModule("RTG") = False) Then
                rs.Edit
                rs!AccountType = eTT_AccountType_SimBroker
                rs!Name = "Sim Broker"
                rs.Update
                
                bChanged = True
            End If
            
            If bChanged Then
                g.Broker.BrokerInfo(eTT_AccountType_SimBroker).LoadCollections
                g.Broker.Refresh eTT_AccountType_SimBroker
                
                g.Broker.BrokerInfo(eTT_AccountType_SimStream).LoadCollections
                g.Broker.Refresh eTT_AccountType_SimStream
                
                Set Account = New cPtAccount
                If Account.Load(rs!AccountID) Then
                    g.Broker.UpdateAccountCache Account
                End If
                
                If FormIsLoaded("frmTTSummary") Then
                    frmTTSummary.DoBrokerTimer
                End If
                If FormIsLoaded("frmAccounts") Then
                    frmAccounts.DoBrokerTimer
                End If
            End If
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_TransAct) Then
        If GetIniFileProperty("IPMsg", 0&, "User", AddSlash(App.Path) & "TransAct.INI") = 1& Then
            SetIniFileProperty "IPMsg", 0&, "User", AddSlash(App.Path) & "TransAct.INI"
            InfBox "For live trading, please call your broker|to get setup on TransAct's new|server for Trade Navigator.", "i", "+-OK", "TransAct"
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SetupInitialAccounts"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoBrokerConnect
'' Description: Automatically connect to any brokerage accounts that were
''              connected last time Trade Navigator was running
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AutoBrokerConnect()
On Error GoTo ErrSection:

    Dim strAccountTypes As String       ' Account types to connect to
    Dim bIsBroker As Boolean            ' Is this user a broker?
    
    strAccountTypes = GetIniFileProperty("AutoConnect", "", "Brokers", g.strIniFile)
    
    If g.Broker.IsBrokerUser(eTT_AccountType_AdvFut) Then
        'If GetIniFileProperty("AdvFutAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.AdvFut Is Nothing Then
                Set g.AdvFut = New cBrokerTt
                g.AdvFut.Init eTT_AccountType_AdvFut
            End If
            If Not g.AdvFut Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_AdvFut) = True Then
                    If g.AdvFut.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.AdvFut.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_AlpariCurrenex) Then
        'If GetIniFileProperty("AlpCnxAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.AlpariCnx Is Nothing Then
                Set g.AlpariCnx = New cBrokerCurrenex
                g.AlpariCnx.Init eTT_AccountType_AlpariCurrenex
            End If
            If Not g.AlpariCnx Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_AlpariCurrenex) = True Then
                    g.AlpariCnx.Broker.Connect
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_AlpariPats) Then
        'If GetIniFileProperty("AlpPatsAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.AlpariPats Is Nothing Then
                Set g.AlpariPats = New cBrokerPats
                g.AlpariPats.Init eTT_AccountType_AlpariPats
            End If
            If Not g.AlpariPats Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_AlpariPats) = True Then
                    g.AlpariPats.Broker.Connect
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_AlpariZenFire) Then
        If g.AlpariZenFire Is Nothing Then
            g.AlpariZenFire = New cRithmic
            g.AlpariZenFire.Init eTT_AccountType_AlpariZenFire
        End If
        
        If Not g.AlpariZenFire Is Nothing Then
            If LiveTradingAllowed(eTT_AccountType_AlpariZenFire) = True Then
                If g.AlpariZenFire.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                    g.AlpariZenFire.Broker.Connect
                End If
            End If
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_Ameritrade) Then
        'If GetIniFileProperty("AmeritradeAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.Ameritrade Is Nothing Then
                Set g.Ameritrade = New cBrokerAmeritrade
                g.Ameritrade.Init eTT_AccountType_Ameritrade
            End If
            If Not g.Ameritrade Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_Ameritrade) = True Then
                    If g.Ameritrade.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.Ameritrade.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_AmpCqg) Then
        'If GetIniFileProperty("AmpAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.AmpCqg Is Nothing Then
                Set g.AmpCqg = New cBrokerCqg
                g.AmpCqg.Init eTT_AccountType_AmpCqg
            End If
            If Not g.AmpCqg Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_AmpCqg) = True Then
                    If g.AmpCqg.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.AmpCqg.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_BornPats) Then
        'If GetIniFileProperty("BornAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.BornPats Is Nothing Then
                Set g.BornPats = New cBrokerPats
                g.BornPats.Init eTT_AccountType_BornPats
            End If
            If Not g.BornPats Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_BornPats) = True Then
                    If g.BornPats.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.BornPats.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_CQG) Then
        'If GetIniFileProperty("CqgAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.CQG Is Nothing Then
                Set g.CQG = New cBrokerCqg
                g.CQG.Init eTT_AccountType_CQG
            End If
            If Not g.CQG Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_CQG) = True Then
                    If g.CQG.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.CQG.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_CtgCqg) Then
        'If GetIniFileProperty("CtgCqgAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.CtgCqg Is Nothing Then
                Set g.CtgCqg = New cBrokerCqg
                g.CtgCqg.Init eTT_AccountType_CtgCqg
            End If
            If Not g.CtgCqg Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_CtgCqg) = True Then
                    If g.CtgCqg.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.CtgCqg.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_CtgPats) Then
        'If GetIniFileProperty("CtgPatsAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.CtgPats Is Nothing Then
                Set g.CtgPats = New cBrokerPats
                g.CtgPats.Init eTT_AccountType_CtgPats
            End If
            If Not g.CtgPats Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_CtgPats) = True Then
                    If g.CtgPats.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.CtgPats.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

'    If g.Broker.IsBrokerUser(eTT_AccountType_CtgPfg) Then
'        'If GetIniFileProperty("CtgPfgAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
'            If g.CtgPfg Is Nothing Then
'                Set g.CtgPfg = New cPFG
'                g.CtgPfg.Init eTT_AccountType_CtgPfg
'            End If
'            If Not g.CtgPfg Is Nothing Then
'                If LiveTradingAllowed(eTT_AccountType_CtgPfg) = True Then
'                    If g.CtgPfg.ConnectionStatus = eGDConnectionStatus_Disconnected Then
'                        g.CtgPfg.Connect
'                    End If
'                End If
'            End If
'        'End If
'    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_Currenex) Then
        'If GetIniFileProperty("CurrenexAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.Currenex Is Nothing Then
                Set g.Currenex = New cBrokerCurrenex
                g.Currenex.Init eTT_AccountType_Currenex
            End If
            If Not g.Currenex Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_Currenex) = True Then
                    g.Currenex.Broker.Connect
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_DemoPats) Then
        'If GetIniFileProperty("DemoPatsAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.DemoPats Is Nothing Then
                Set g.DemoPats = New cBrokerPats
                g.DemoPats.Init eTT_AccountType_DemoPats
            End If
            If Not g.DemoPats Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_DemoPats) = True Then
                    If g.DemoPats.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.DemoPats.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_Etrade) Then
        'If GetIniFileProperty("EtradeAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.Etrade Is Nothing Then
                Set g.Etrade = New cBrokerEtrade
                g.Etrade.Init eTT_AccountType_Etrade
            End If
            If Not g.Etrade Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_Etrade) = True Then
                    If g.Etrade.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.Etrade.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

'    If g.Broker.IsBrokerUser(eTT_AccountType_FintecPfg) Then
'        'If GetIniFileProperty("FintecPfgAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
'            If g.FintecPfg Is Nothing Then
'                Set g.FintecPfg = New cPFG
'                g.FintecPfg.Init eTT_AccountType_FintecPfg
'            End If
'            If Not g.FintecPfg Is Nothing Then
'                If LiveTradingAllowed(eTT_AccountType_FintecPfg) = True Then
'                    If g.FintecPfg.ConnectionStatus = eGDConnectionStatus_Disconnected Then
'                        g.FintecPfg.Connect
'                    End If
'                End If
'            End If
'        'End If
'    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_FptCqg) Then
        If g.FptCqg Is Nothing Then
            g.FptCqg = New cBrokerCqg
            g.FptCqg.Init eTT_AccountType_FptCqg
        End If
        
        If Not g.FptCqg Is Nothing Then
            If LiveTradingAllowed(eTT_AccountType_FptCqg) = True Then
                If g.FptCqg.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                    g.FptCqg.Broker.Connect
                End If
            End If
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_FptOec) Then
        If g.FptOec Is Nothing Then
            g.FptOec = New cBrokerOec
            g.FptOec.Init eTT_AccountType_FptOec
        End If
        
        If Not g.FptOec Is Nothing Then
            If LiveTradingAllowed(eTT_AccountType_FptOec) = True Then
                If g.FptOec.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                    g.FptOec.Broker.Connect
                End If
            End If
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_FxddCurrenex) Then
        'If GetIniFileProperty("FxddCnxAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.FxddCnx Is Nothing Then
                Set g.FxddCnx = New cBrokerCurrenex
                g.FxddCnx.Init eTT_AccountType_FxddCurrenex
            End If
            If Not g.FxddCnx Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_FxddCurrenex) = True Then
                    g.FxddCnx.Broker.Connect
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_Gft) Then
        'If GetIniFileProperty("GftAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.Gft Is Nothing Then
                Set g.Gft = New cBrokerGft
                g.Gft.Init eTT_AccountType_Gft
            End If
            If Not g.Gft Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_Gft) = True Then
                    If g.Gft.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.Gft.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_Ideal) Then
        'If GetIniFileProperty("IdealAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.Ideal Is Nothing Then
                Set g.Ideal = New cIntBrokers
                g.Ideal.Init eTT_AccountType_Ideal
            End If
            If Not g.Ideal Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_Ideal) = True Then
                    If g.Ideal.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.Ideal.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_IntBrokers) Then
        'If GetIniFileProperty("IbAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.IntBroker Is Nothing Then
                Set g.IntBroker = New cIntBrokers
                g.IntBroker.Init eTT_AccountType_IntBrokers
            End If
            If Not g.IntBroker Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_IntBrokers) = True Then
                    If g.IntBroker.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.IntBroker.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_KnightCqg) Then
        'If GetIniFileProperty("KnightCqgAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.KnightCqg Is Nothing Then
                Set g.KnightCqg = New cBrokerCqg
                g.KnightCqg.Init eTT_AccountType_KnightCqg
            End If
            If Not g.KnightCqg Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_KnightCqg) = True Then
                    g.KnightCqg.Broker.Connect
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_KnightCurrenex) Then
        'If GetIniFileProperty("KnightCnxAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.KnightCnx Is Nothing Then
                Set g.KnightCnx = New cBrokerCurrenex
                g.KnightCnx.Init eTT_AccountType_KnightCurrenex
            End If
            If Not g.KnightCnx Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_KnightCurrenex) = True Then
                    g.KnightCnx.Broker.Connect
                End If
            End If
        'End If
    End If

'    If g.Broker.IsBrokerUser(eTT_AccountType_LindWaldock) Then
'        'If GetIniFileProperty("LindWaldockAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
'            If g.LindWaldock Is Nothing Then
'                Set g.LindWaldock = New cXpress
'                g.LindWaldock.Init eTT_AccountType_LindWaldock
'            End If
'            If Not g.LindWaldock Is Nothing Then
'                If LiveTradingAllowed(eTT_AccountType_LindWaldock) = True Then
'                    g.LindWaldock.Connect
'                End If
'            End If
'        'End If
'    End If
    
'    If g.Broker.IsBrokerUser(eTT_AccountType_ManExpress) Then
'        'If GetIniFileProperty("ManExpressAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
'            If g.ManExpress Is Nothing Then
'                Set g.ManExpress = New cXpress
'                g.ManExpress.Init eTT_AccountType_ManExpress
'            End If
'            If Not g.ManExpress Is Nothing Then
'                If LiveTradingAllowed(eTT_AccountType_ManExpress) = True Then
'                    g.ManExpress.Connect
'                End If
'            End If
'        'End If
'    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_Oec) Then
        If g.Oec Is Nothing Then
            g.Oec = New cBrokerOec
            g.Oec.Init eTT_AccountType_Oec
        End If
        
        If Not g.Oec Is Nothing Then
            If LiveTradingAllowed(eTT_AccountType_Oec) = True Then
                If g.Oec.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                    g.Oec.Broker.Connect
                End If
            End If
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_Optimus) Then
        If g.Optimus Is Nothing Then
            g.Optimus = New cRithmic
            g.Optimus.Init eTT_AccountType_Optimus
        End If
        
        If Not g.Optimus Is Nothing Then
            If LiveTradingAllowed(eTT_AccountType_Optimus) = True Then
                If g.Optimus.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                    g.Optimus.Broker.Connect
                End If
            End If
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_OpVest) Then
        If g.OpVest Is Nothing Then
            g.OpVest = New cRithmic
            g.OpVest.Init eTT_AccountType_OpVest
        End If
        
        If Not g.OpVest Is Nothing Then
            If LiveTradingAllowed(eTT_AccountType_OpVest) = True Then
                If g.OpVest.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                    g.OpVest.Broker.Connect
                End If
            End If
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_PATS) Then
        'If GetIniFileProperty("PatsAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.Pats Is Nothing Then
                Set g.Pats = New cBrokerPats
                g.Pats.Init eTT_AccountType_PATS
            End If
            If Not g.Pats Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_PATS) = True Then
                    If g.Pats.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.Pats.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If
    
'    If g.Broker.IsBrokerUser(eTT_AccountType_PFG) Then
'        'If GetIniFileProperty("PfgAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
'            If g.PFG Is Nothing Then
'                Set g.PFG = New cPFG
'                g.PFG.Init eTT_AccountType_PFG
'            End If
'            If Not g.PFG Is Nothing Then
'                If LiveTradingAllowed(eTT_AccountType_PFG) = True Then
'                    If g.PFG.ConnectionStatus = eGDConnectionStatus_Disconnected Then
'                        g.PFG.Connect
'                    End If
'                End If
'            End If
'        'End If
'    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_RcgPats) Then
        'If GetIniFileProperty("RcgPatsAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.RcgPats Is Nothing Then
                Set g.RcgPats = New cBrokerPats
                g.RcgPats.Init eTT_AccountType_RcgPats
            End If
            If Not g.RcgPats Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_RcgPats) = True Then
                    If g.RcgPats.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.RcgPats.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_Rithmic) Then
        If g.Rithmic Is Nothing Then
            g.Rithmic = New cRithmic
            g.Rithmic.Init eTT_AccountType_Rithmic
        End If
        
        If Not g.Rithmic Is Nothing Then
            If LiveTradingAllowed(eTT_AccountType_Rithmic) = True Then
                If g.Rithmic.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                    g.Rithmic.Broker.Connect
                End If
            End If
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_RjoCqg, bIsBroker) Then
        'If GetIniFileProperty("RjoAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.RjoCqg Is Nothing Then
                Set g.RjoCqg = New cBrokerCqg
                g.RjoCqg.Init eTT_AccountType_RjoCqg, bIsBroker
            End If
            If Not g.RjoCqg Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_RjoCqg) = True Then
                    If g.RjoCqg.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.RjoCqg.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_RjoPats) Then
        'If GetIniFileProperty("RjoAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.RjoPats Is Nothing Then
                Set g.RjoPats = New cBrokerPats
                g.RjoPats.Init eTT_AccountType_RjoPats
            End If
            If Not g.RjoPats Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_RjoPats) = True Then
                    If g.RjoPats.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.RjoPats.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_RjoHkPats) Then
        'If GetIniFileProperty("RjoHkAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.RjoHkPats Is Nothing Then
                Set g.RjoHkPats = New cBrokerPats
                g.RjoHkPats.Init eTT_AccountType_RjoHkPats
            End If
            If Not g.RjoHkPats Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_RjoHkPats) = True Then
                    If g.RjoHkPats.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.RjoHkPats.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_RobbinsCqg) Then
        'If GetIniFileProperty("RobbinsCqgAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.RobbinsCqg Is Nothing Then
                Set g.RobbinsCqg = New cBrokerCqg
                g.RobbinsCqg.Init eTT_AccountType_RobbinsCqg
            End If
            If Not g.RobbinsCqg Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_RobbinsCqg) = True Then
                    g.RobbinsCqg.Broker.Connect
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_SimBroker) Then
        If g.SimTradeTs Is Nothing Then
            Set g.SimTradeTs = New cSimTradeTs
            g.SimTradeTs.Init eTT_AccountType_SimBroker, frmOnlineBroker.txtSalmonCallbackTs, frmOnlineBroker.tmrTradeServer
        End If
        If Not g.SimTradeTs Is Nothing Then
            If LiveTradingAllowed(eTT_AccountType_SimBroker) = True Then
                If g.SimTradeTs.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                    g.SimTradeTs.Broker.Connect
                End If
            End If
        End If
    End If
    
' DAJ 10/06/2015: With Tradier, we are bringing up a dialog with a web page instead of a true dialog, so you cannot
' cancel out of the dialog like the others which makes it a bit of a pain.  I am taking this out of the auto connect
' for now...
'    If g.Broker.IsBrokerUser(eTT_AccountType_Tradier) Then
'        'If GetIniFileProperty("EtradeAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
'            If g.Tradier Is Nothing Then
'                Set g.Tradier = New cBrokerTradier
'                g.Tradier.Init eTT_AccountType_Tradier
'            End If
'            If Not g.Tradier Is Nothing Then
'                If LiveTradingAllowed(eTT_AccountType_Tradier) = True Then
'                    If g.Tradier.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
'                        g.Tradier.Broker.Connect
'                    End If
'                End If
'            End If
'        'End If
'    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_TransAct) Then
        'If GetIniFileProperty("TransactAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.Transact Is Nothing Then Set g.Transact = New cTransact
            If Not g.Transact Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_TransAct) = True Then
                    If g.Transact.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.Transact.Connect
                    End If
                End If
            End If
        'End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_TT) Then
        'If GetIniFileProperty("TtAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.TT Is Nothing Then
                Set g.TT = New cBrokerTt
                g.TT.Init eTT_AccountType_TT
            End If
            If Not g.TT Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_TT) = True Then
                    If g.TT.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.TT.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_VanKarCurrenex) Then
        'If GetIniFileProperty("VanKarCnxAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.VanKarCnx Is Nothing Then
                Set g.VanKarCnx = New cBrokerCurrenex
                g.VanKarCnx.Init eTT_AccountType_VanKarCurrenex
            End If
            If Not g.VanKarCnx Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_VanKarCurrenex) = True Then
                    g.VanKarCnx.Broker.Connect
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_Vision) Then
        If g.Vision Is Nothing Then
            g.Vision = New cRithmic
            g.Vision.Init eTT_AccountType_Vision
        End If
        
        If Not g.Vision Is Nothing Then
            If LiveTradingAllowed(eTT_AccountType_Vision) = True Then
                If g.Vision.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                    g.Vision.Broker.Connect
                End If
            End If
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_VisionCqg) Then
        'If GetIniFileProperty("VisionCqgAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.VisionCqg Is Nothing Then
                Set g.VisionCqg = New cBrokerCqg
                g.VisionCqg.Init eTT_AccountType_VisionCqg
            End If
            If Not g.VisionCqg Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_VisionCqg) = True Then
                    g.VisionCqg.Broker.Connect
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_ZanerCqg) Then
        'If GetIniFileProperty("ZanerCqgAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.ZanerCqg Is Nothing Then
                Set g.ZanerCqg = New cBrokerCqg
                g.ZanerCqg.Init eTT_AccountType_ZanerCqg
            End If
            If Not g.ZanerCqg Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_ZanerCqg) = True Then
                    g.ZanerCqg.Broker.Connect
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_ZanerCurrenex) Then
        'If GetIniFileProperty("ZanerCnxAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.ZanerCnx Is Nothing Then
                Set g.ZanerCnx = New cBrokerCurrenex
                g.ZanerCnx.Init eTT_AccountType_ZanerCurrenex
            End If
            If Not g.ZanerCnx Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_ZanerCurrenex) = True Then
                    g.ZanerCnx.Broker.Connect
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_ZanerPats) Then
        'If GetIniFileProperty("ZanerAutoConnect", "0", "Brokers", g.strIniFile) = "1" Then
            If g.ZanerPats Is Nothing Then
                Set g.ZanerPats = New cBrokerPats
                g.ZanerPats.Init eTT_AccountType_ZanerPats
            End If
            If Not g.ZanerPats Is Nothing Then
                If LiveTradingAllowed(eTT_AccountType_ZanerPats) = True Then
                    If g.ZanerPats.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.ZanerPats.Broker.Connect
                    End If
                End If
            End If
        'End If
    End If

    If g.Broker.IsBrokerUser(eTT_AccountType_ZanerRithmic) Then
        If g.ZanerRithmic Is Nothing Then
            g.ZanerRithmic = New cRithmic
            g.ZanerRithmic.Init eTT_AccountType_ZanerRithmic
        End If
        
        If Not g.ZanerRithmic Is Nothing Then
            If LiveTradingAllowed(eTT_AccountType_ZanerRithmic) = True Then
                If g.ZanerRithmic.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                    g.ZanerRithmic.Broker.Connect
                End If
            End If
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_ZanerZenFire) Then
        If g.ZanerZenFire Is Nothing Then
            g.ZanerZenFire = New cRithmic
            g.ZanerZenFire.Init eTT_AccountType_ZanerZenFire
        End If
        
        If Not g.ZanerZenFire Is Nothing Then
            If LiveTradingAllowed(eTT_AccountType_ZanerZenFire) = True Then
                If g.ZanerZenFire.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                    g.ZanerZenFire.Broker.Connect
                End If
            End If
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_ZenFire) Then
        If g.ZenFire Is Nothing Then
            g.ZenFire = New cRithmic
            g.ZenFire.Init eTT_AccountType_ZenFire
        End If
        
        If Not g.ZenFire Is Nothing Then
            If LiveTradingAllowed(eTT_AccountType_ZenFire) = True Then
                If g.ZenFire.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                    g.ZenFire.Broker.Connect
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.AutoBrokerConnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAccountCombos
'' Description: Refresh the account combo boxes on charts and price ladders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshAccountCombos()
On Error GoTo ErrSection:

    Dim i&
    Dim frm As Form
    
    DebugLog "Begin mTradeTracker.RefreshAccountCombos (Count = " & Str(Forms.Count) & ")"
    For i = 0 To Forms.Count - 1
        If g.bUnloading Then
            Exit For
        Else
            Set frm = Forms(i)
            
            DebugLog vbTab & "Name = '" & frm.Name & "'; Caption = '" & frm.Caption & "' (Count = " & Str(Forms.Count) & ")"
            If IsFrmChart(frm) Then
                DebugLog vbTab & vbTab & "IsFrmChart = True"
                If frm.Chart.ShowTrades = 2 Then
                    DebugLog vbTab & vbTab & "ShowTrades = 2"
                    PopulateAccountsCbo frm.cboAccounts, frm.Chart.TradeAccountID, True
                End If
            ElseIf TypeOf frm Is frmTickDistribution Then
                DebugLog vbTab & vbTab & "TypeOf frm is frmTickDistribution"
                If frm.ShowOrderBar And frm.DisplayStyle = 0 Then
                    DebugLog vbTab & vbTab & "frm.ShowOrderBar = True And frm.DisplayStyle = 0"
                    PopulateAccountsCbo frm.cboAccounts, frm.TradeAccountID, True
                End If
            End If
        End If
    Next
    DebugLog "End mTradeTracker.RefreshAccountCombos (Count = " & Str(Forms.Count) & ")"
    
ErrExit:
    Exit Sub
    
ErrSection:
    If frm Is Nothing Then
        DebugLog vbTab & vbTab & "Error: '" & Err.Description & "'; Form = Nothing (Count = " & Str(Forms.Count) & ")"
    Else
        DebugLog vbTab & vbTab & "Error: '" & Err.Description & "'; Name = '" & frm.Name & "'; Caption = '" & frm.Caption & "' (Count = " & Str(Forms.Count) & ")"
    End If
    RaiseError "mTradeTracker.RefreshAccountCombos"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllowLiveAutoTrading
'' Description: Determine whether or not to allow live automated trading
'' Inputs:      None
'' Returns:     True if allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllowLiveAutoTrading(ByVal nAcctType As eTT_AccountType, Optional ByVal bShowMsg As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function

    If Not g.Broker.IsLiveAccount(nAcctType) Then
        bReturn = True
    Else
        bReturn = (HasModule("BRKRAUTO") Or FileExist(AddSlash(App.Path) & "AdvBroker.FLG")) And (LiveTradingAllowed(nAcctType) = True)
    End If
    
    If (bReturn = False) And (bShowMsg = True) Then
        InfBox "You are currently not allowed to do automated trading in a live account.", "!", , "Live Automated Trading"
    End If

    AllowLiveAutoTrading = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.AllowLiveAutoTrading"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIsPending
'' Description: Determine whether or not an order is in a pending status
'' Inputs:      Order
'' Returns:     True if pending, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderIsPending(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Select Case Order.Status
        Case eTT_OrderStatus_Sent, eTT_OrderStatus_AmendPending, eTT_OrderStatus_CancelPending, eTT_OrderStatus_Queued, eTT_OrderStatus_ParkPending
            OrderIsPending = True
        
        Case Else
            OrderIsPending = False
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.OrderIsPending"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PopulateAccountsCbo
'' Description: Populate the accounts combo with the appropriate accounts
'' Inputs:      Combo Control, Account ID, Refresh Accounts?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateAccountsCbo(cboCtrl As ctlUniComboImageXP, ByVal nAccountID As Long, Optional ByVal bRefreshAccounts As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, j&, iListIndex&, bAdd As Boolean
    Dim rs As Recordset                 ' Recordset into the database

    If bRefreshAccounts Then cboCtrl.Clear
    
    If g.nReplaySession > 0 Then
        If cboCtrl.ListCount = 0 Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountID]=" & Str(g.nReplayAccountID) & ";", dbOpenDynaset)
            If Not (rs.BOF And rs.EOF) Then
                With cboCtrl
                    .AddItem rs!Name
                    .ItemData(.NewIndex) = rs!AccountID
                    If j = nAccountID Then
                        iListIndex = .NewIndex
                    End If
                End With
            End If
        Else
            j = nAccountID
            With cboCtrl
                If j = .ItemData(.ListIndex) Then
                    iListIndex = .ListIndex
                Else
                    For i = 0 To .ListCount - 1
                        If j = .ItemData(i) Then
                            iListIndex = i
                            Exit For
                        End If
                    Next
                End If
            End With
        End If
    Else
        If cboCtrl.ListCount = 0 Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)
            If (rs.EOF And rs.BOF) Then
                rs.Requery
                ''If FormIsLoaded("frmTTSummary") Then frmTTSummary.LoadAccountsGrid
            End If
            Do While Not rs.EOF
                j = rs!AccountID
                With cboCtrl
                    bAdd = True
                    
                    If rs!AccountType = eTT_AccountType_TransAct Then
                        If g.Transact Is Nothing Then
                            bAdd = False
                        Else
                            If (g.Broker.IsBrokerSimUser(eTT_AccountType_TransAct) = True) And (g.Transact.UserName = g.Transact.SimUserUserName) Then
                                bAdd = False
                            Else
                                bAdd = g.Broker.IsBrokerUser(eTT_AccountType_TransAct) And (Not TransActSimulatedAccount(rs!AccountNumber))
                            End If
                        End If
                        
                    ElseIf g.Broker.IsLiveAccount(rs!AccountType) = False Then
                        bAdd = (rs!AccountType <> eTT_AccountType_SimReplay)
                    Else
                        bAdd = g.Broker.IsBrokerUser(rs!AccountType)
                    End If
                    
                    If bAdd Then
                        .AddItem rs!Name 'AccountNumber
                        .ItemData(.NewIndex) = rs!AccountID
                        If j = nAccountID Then
                            iListIndex = .NewIndex
                        End If
                    End If
                End With
                rs.MoveNext
            Loop
        Else
            j = nAccountID
            With cboCtrl
                If j = .ItemData(.ListIndex) Then
                    iListIndex = .ListIndex
                Else
                    For i = 0 To .ListCount - 1
                        If j = .ItemData(i) Then
                            iListIndex = i
                            Exit For
                        End If
                    Next
                End If
            End With
        End If
    End If

    With cboCtrl
        If .ListCount <= 0 Then ' 1 Then
            .Enabled = False
        Else
            .Enabled = True
        End If
        If iListIndex < .ListCount Then
            .ListIndex = iListIndex
        End If
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mTradeTracker.PopulateAccountsCbo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeAccountCombo
'' Description: Determine whether to allow a change to this account
'' Inputs:      New Account Name
'' Returns:     True if allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ChangeAccountCombo(ByVal strNewAccount As String) As Boolean
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID
    Dim Account As cPtAccount           ' Temporary account object
    Dim BrokerObj As cBroker            ' Broker object
    Dim bReturn As Boolean              ' Return value for the function
    
    lAccountID = g.Broker.AccountIDForName(strNewAccount)
    Set Account = g.Broker.Account(lAccountID)
    If Not Account Is Nothing Then
        Select Case Account.AccountType
'            Case eTT_AccountType_CtgPfg
'                If Not g.CtgPfg Is Nothing Then
'                    Select Case g.CtgPfg.ConnectionStatus
'                        Case eGDConnectionStatus_Connected
'                            bReturn = True
'
'                        Case eGDConnectionStatus_Connecting
'                            bReturn = True
'
'                        Case eGDConnectionStatus_Disconnecting
'                            bReturn = False
'                            InfBox "The " & g.CtgPfg.BrokerName & " connection is in the process of disconnecting.  Please try again later.", "!", "+-OK", g.CtgPfg.BrokerName & " Connection"
'
'                        Case eGDConnectionStatus_Disconnected
'                            If InfBox("You are not currently connected to " & g.CtgPfg.BrokerName & ".  Would you like to connect to the " & g.CtgPfg.BrokerName & " servers?", "?", "+Yes|-No", g.CtgPfg.BrokerName & " Connection") = "Y" Then
'                                g.CtgPfg.Connect
'                            End If
'                            bReturn = False
'
'                    End Select
'                Else
'                    bReturn = False
'                End If
'
'            Case eTT_AccountType_FintecPfg
'                If Not g.FintecPfg Is Nothing Then
'                    Select Case g.FintecPfg.ConnectionStatus
'                        Case eGDConnectionStatus_Connected
'                            bReturn = True
'
'                        Case eGDConnectionStatus_Connecting
'                            bReturn = True
'
'                        Case eGDConnectionStatus_Disconnecting
'                            bReturn = False
'                            InfBox "The " & g.FintecPfg.BrokerName & " connection is in the process of disconnecting.  Please try again later.", "!", "+-OK", g.FintecPfg.BrokerName & " Connection"
'
'                        Case eGDConnectionStatus_Disconnected
'                            If InfBox("You are not currently connected to " & g.FintecPfg.BrokerName & ".  Would you like to connect to the " & g.FintecPfg.BrokerName & " servers?", "?", "+Yes|-No", g.FintecPfg.BrokerName & " Connection") = "Y" Then
'                                g.FintecPfg.Connect
'                            End If
'                            bReturn = False
'
'                    End Select
'                Else
'                    bReturn = False
'                End If
'
'            Case eTT_AccountType_LindWaldock
'                If Not g.LindWaldock Is Nothing Then
'                    Select Case g.LindWaldock.ConnectionStatus
'                        Case eGDConnectionStatus_Connected
'                            bReturn = True
'
'                        Case eGDConnectionStatus_Connecting
'                            bReturn = True
'
'                        Case eGDConnectionStatus_Disconnecting
'                            bReturn = False
'                            InfBox "The " & g.Broker.BrokerName(eTT_AccountType_LindWaldock) & " connection is in the process of disconnecting.  Please try again later.", "!", "+-OK", g.Broker.BrokerName(eTT_AccountType_LindWaldock) & " Connection"
'
'                        Case eGDConnectionStatus_Disconnected
'                            If InfBox("You are not currently connected to " & g.Broker.BrokerName(eTT_AccountType_LindWaldock) & ".  Would you like to connect to the " & g.Broker.BrokerName(eTT_AccountType_LindWaldock) & " servers?", "?", "+Yes|-No", g.Broker.BrokerName(eTT_AccountType_LindWaldock) & " Connection") = "Y" Then
'                                g.LindWaldock.Connect
'                            End If
'                            bReturn = False
'
'                    End Select
'                Else
'                    bReturn = False
'                End If
'
'            Case eTT_AccountType_ManExpress
'                If Not g.ManExpress Is Nothing Then
'                    Select Case g.ManExpress.ConnectionStatus
'                        Case eGDConnectionStatus_Connected
'                            bReturn = True
'
'                        Case eGDConnectionStatus_Connecting
'                            bReturn = True
'
'                        Case eGDConnectionStatus_Disconnecting
'                            bReturn = False
'                            InfBox "The ManExpress connection is in the process of disconnecting.  Please try again later.", "!", "+-OK", "ManExpress Connection"
'
'                        Case eGDConnectionStatus_Disconnected
'                            If InfBox("You are not currently connected to ManExpress.  Would you like to connect to the ManExpress servers?", "?", "+Yes|-No", "ManExpress Connection") = "Y" Then
'                                g.ManExpress.Connect
'                            End If
'                            bReturn = False
'
'                    End Select
'                Else
'                    bReturn = False
'                End If
'
'            Case eTT_AccountType_PFG
'                If Not g.PFG Is Nothing Then
'                    Select Case g.PFG.ConnectionStatus
'                        Case eGDConnectionStatus_Connected
'                            bReturn = True
'
'                        Case eGDConnectionStatus_Connecting
'                            bReturn = True
'
'                        Case eGDConnectionStatus_Disconnecting
'                            bReturn = False
'                            InfBox "The " & g.PFG.BrokerName & " connection is in the process of disconnecting.  Please try again later.", "!", "+-OK", g.PFG.BrokerName & " Connection"
'
'                        Case eGDConnectionStatus_Disconnected
'                            If InfBox("You are not currently connected to " & g.PFG.BrokerName & ".  Would you like to connect to the " & g.PFG.BrokerName & " servers?", "?", "+Yes|-No", g.PFG.BrokerName & " Connection") = "Y" Then
'                                g.PFG.Connect
'                            End If
'                            bReturn = False
'
'                    End Select
'                Else
'                    bReturn = False
'                End If
                
            Case eTT_AccountType_TransAct
                If Not g.Transact Is Nothing Then
                    Select Case g.Transact.ConnectionStatus
                        Case eGDConnectionStatus_Connected
                            If g.Transact.Account = Account.AccountNumber Then
                                bReturn = True
                            Else
                                bReturn = g.Transact.Connect(Account.UserName, Account.AccountNumber)
                            End If
                            
                        Case eGDConnectionStatus_Connecting
                            bReturn = True
                        
                        Case eGDConnectionStatus_Disconnecting
                            bReturn = False
                            InfBox "The TransAct connection is in the process of disconnecting.  Please try again later.", "!", "+-OK", "TransAct Connection"
                        
                        Case eGDConnectionStatus_Disconnected
                            If InfBox("You are not currently connected to TransAct.  Would you like to connect to the TransAct servers?", "?", "+Yes|-No", "TransAct Connection") = "Y" Then
                                bReturn = g.Transact.Connect(Account.UserName, Account.AccountNumber)
                            Else
                                bReturn = False
                            End If
                            
                    End Select
                Else
                    bReturn = False
                End If
                
            Case Else
                Set BrokerObj = g.Broker.Broker(Account.AccountType)
                If Not BrokerObj Is Nothing Then
                    Select Case BrokerObj.ConnectionStatus
                        Case eGDConnectionStatus_Connected
                            bReturn = True
                            
                        Case eGDConnectionStatus_Connecting
                            bReturn = True
                        
                        Case eGDConnectionStatus_Disconnecting
                            bReturn = False
                            InfBox "The " & BrokerObj.BrokerName & " connection is in the process of disconnecting.  Please try again later.", "!", "+-OK", BrokerObj.BrokerName & " Connection"
                        
                        Case eGDConnectionStatus_Disconnected
                            If InfBox("You are not currently connected to " & BrokerObj.BrokerName & ".  Would you like to connect to the " & BrokerObj.BrokerName & " servers?", "?", "+Yes|-No", BrokerObj.BrokerName & " Connection") = "Y" Then
                                BrokerObj.Connect
                            End If
                            bReturn = False
                            
                    End Select
                Else
                    bReturn = False
                End If
            
        End Select
    End If
    
    ChangeAccountCombo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ChangeAccountCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResyncSubscriptionList
'' Description: Synchronize the subscription list with the broker object
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ResyncSubscriptionList(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Select Case nBroker
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then g.Transact.ResyncSubscriptionList
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ResyncSubscriptionList"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenForSymbol
'' Description: Flatten the user's position for a symbol and cancel open orders
'' Inputs:      Account, Symbol or SymbolID, Auto Trade Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FlattenForSymbol(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant, ByVal lAutoTradeItemID As Long)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol
    Dim strAccount As String            ' Account number for Account ID passed in
    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim bDataAvailable As Boolean       ' Is the data available for the symbol?
        
    vSymbolOrSymbolID = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    strAccount = g.Broker.AccountNumberForID(lAccountID)
    nBroker = g.Broker.AccountTypeForID(lAccountID)
    
    ' If user is using NextGen, make sure that the data is available before allowing them
    ' to flatten the position (#5768)...
    bDataAvailable = g.RealTime.RtDataAvailable(vSymbolOrSymbolID, ePRD_Days)
    
    ' Only allow a Flatten for a symbol that does not currently have a position mismatch...
    If (g.Broker.PositionMatch(lAccountID, vSymbolOrSymbolID) = False) Then
        g.Broker.BrokerDebug nBroker, "FlattenForSymbol(" & strAccount & ", " & strSymbol & ", " & Str(lAutoTradeItemID) & "): Aborted due to Position Mismatch"
        InfBox "Trade Navigator has received inconsistent data from the broker for " & strSymbol & " in account '" & strAccount & "' and therefore cannot perform a Flatten on the position.||PLEASE CALL YOUR BROKER IMMEDIATELY TO VERIFY YOUR POSITION IN THIS ACCOUNT.||", "!", , "Inconsistent Broker Data"
    ElseIf (bDataAvailable = False) And (nBroker = eTT_AccountType_SimStream) Then
        g.Broker.BrokerDebug nBroker, "FlattenForSymbol(" & strAccount & ", " & strSymbol & ", " & Str(lAutoTradeItemID) & "): Aborted due to data not available"
        InfBox "Trade Navigator does not have all of the data for " & strSymbol & " yet.  Please wait to Flatten until all data is available", "!", , "Data not Available"
    Else
        g.Broker.BrokerDebug nBroker, "FlattenForSymbol(" & strAccount & ", " & strSymbol & ", " & Str(lAutoTradeItemID) & ")"
        g.TsoGroups.CancelForSymbol vSymbolOrSymbolID, lAccountID, "Flatten for Symbol"
        g.FlattenQueue.AddToFlattenQueue strAccount, strSymbol, lAutoTradeItemID, eGDFlattenQueueOperation_Flatten
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.FlattenForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelAllForSymbol
'' Description: Cancel all orders for a symbol
'' Inputs:      Account, Symbol or SymbolID, Auto Trade Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelAllForSymbol(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant, ByVal lAutoTradeItemID As Long)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol
    Dim strAccount As String            ' Account number for Account ID passed in
    Dim nBroker As eTT_AccountType      ' Broker for the account
        
    vSymbolOrSymbolID = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    strAccount = g.Broker.AccountNumberForID(lAccountID)
    nBroker = g.Broker.AccountTypeForID(lAccountID)
    
    g.Broker.BrokerDebug nBroker, "CancelAllForSymbol(" & strAccount & ", " & strSymbol & ", " & Str(lAutoTradeItemID) & ")"
    g.TsoGroups.CancelForSymbol vSymbolOrSymbolID, lAccountID, "Cancel All for Symbol"
    g.FlattenQueue.AddToFlattenQueue strAccount, strSymbol, lAutoTradeItemID, eGDFlattenQueueOperation_CancelAll
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.CancelAllForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReverseForSymbol
'' Description: Reverse the user's position for a symbol and cancel open orders
'' Inputs:      Account, Symbol or SymbolID, Auto Trade Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReverseForSymbol(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant, ByVal lAutoTradeItemID As Long)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol
    Dim strAccount As String            ' Account number for Account ID passed in
    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim bDataAvailable As Boolean       ' Is the data available for the symbol?
        
    vSymbolOrSymbolID = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    strAccount = g.Broker.AccountNumberForID(lAccountID)
    nBroker = g.Broker.AccountTypeForID(lAccountID)
    
    ' If user is using NextGen, make sure that the data is available before allowing them
    ' to flatten the position (#5768)...
    bDataAvailable = g.RealTime.RtDataAvailable(vSymbolOrSymbolID, ePRD_Days)
    
    ' Only allow a Reverse for a symbol that does not currently have a position mismatch...
    If (g.Broker.PositionMatch(lAccountID, vSymbolOrSymbolID) = False) Then
        g.Broker.BrokerDebug nBroker, "ReverseForSymbol(" & strAccount & ", " & strSymbol & ", " & Str(lAutoTradeItemID) & "): Aborted due to Position Mismatch"
        InfBox "Trade Navigator has received inconsistent data from the broker for " & strSymbol & " in account '" & strAccount & "' and therefore cannot perform a Reverse on the position.||PLEASE CALL YOUR BROKER IMMEDIATELY TO VERIFY YOUR POSITION IN THIS ACCOUNT.||", "!", , "Inconsistent Broker Data"
    ElseIf (bDataAvailable = False) And (nBroker = eTT_AccountType_SimStream) Then
        g.Broker.BrokerDebug nBroker, "ReverseForSymbol(" & strAccount & ", " & strSymbol & ", " & Str(lAutoTradeItemID) & "): Aborted due to data not available"
        InfBox "Trade Navigator does not have all of the data for " & strSymbol & " yet.  Please wait to Reverse until all data is available", "!", , "Data not Available"
    Else
        g.Broker.BrokerDebug nBroker, "ReverseForSymbol(" & strAccount & ", " & strSymbol & ", " & Str(lAutoTradeItemID) & ")"
        g.TsoGroups.CancelForSymbol vSymbolOrSymbolID, lAccountID, "Reverse for Symbol"
        g.FlattenQueue.AddToFlattenQueue strAccount, strSymbol, lAutoTradeItemID, eGDFlattenQueueOperation_Reverse
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ReverseForSymbol"
    
End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenForSymbol
'' Description: Flatten the user's position for a symbol and cancel open orders
'' Inputs:      Account, Symbol or SymbolID, Auto Trade Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FlattenForSymbolOld(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant, ByVal lAutoTradeItemID As Long)
On Error GoTo ErrSection:

    Static strInProcess As String       ' String of in process items
    Dim strKey As String                ' Key into the in-process list
    Dim lSymbolID As Long               ' Symbol ID
    Dim strSymbol As String             ' Symbol
    Dim lPosition As Long               ' Current position
    Dim lCount As Long                  ' Loop counter
    Dim lTimeOut As Long                ' Timeout counter
    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim strMessage As String            ' Message to display to the user
    Dim bContinue As Boolean            ' Do we want to continue?
    Dim TradeItem As cAutoTradeItem     ' Automated Trading Item
        
    vSymbolOrSymbolID = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    nBroker = g.Broker.AccountTypeForID(lAccountID)
    
    g.Broker.BrokerDebug nBroker, "FlattenForSymbol(" & g.Broker.AccountNumberForID(lAccountID) & ", " & strSymbol & ", " & Str(lAutoTradeItemID) & ")"
    
    If lSymbolID = 0& Then
        strKey = "," & Str(lAccountID) & vbTab & strSymbol & vbTab & Str(lAutoTradeItemID) & ","
    Else
        strKey = "," & Str(lAccountID) & vbTab & Str(lSymbolID) & vbTab & Str(lAutoTradeItemID) & ","
    End If
    
    If InStr(strInProcess, strKey) = 0 Then
        strInProcess = strInProcess & strKey
        
        If (lAutoTradeItemID > 0&) Then
            Set TradeItem = g.TradingItems(Str(lAutoTradeItemID))
            If Not TradeItem Is Nothing Then
                TradeItem.ClosePosition g.Broker.ConfirmManual
            End If
        ElseIf Len(g.OrderStrategies.ExitForAccountAndSymbol(lAccountID, vSymbolOrSymbolID)) > 0 Then
            ' Have auto exit handle it's own order cancellation and position flattening...
            g.OrderStrategies.Flatten lAccountID, vSymbolOrSymbolID
            
            ' Cancel any remaining, non auto-exit orders for the account and symbol...
            g.Broker.CancelWorkingOrders lAccountID, vSymbolOrSymbolID, lAutoTradeItemID
        Else
            bContinue = True
            
            ' Attempt to cancel all working orders for the account/symbol...
            g.Broker.CancelWorkingOrders lAccountID, vSymbolOrSymbolID, lAutoTradeItemID
            
            ' Wait until no working orders left or 10 seconds have elapsed...
            lTimeOut = 0&
            Do
                Sleep 1&
                lTimeOut = lTimeOut + 1&
            Loop While (g.Broker.HasWorkingOrders(lAccountID, vSymbolOrSymbolID, lAutoTradeItemID)) And (lTimeOut < 10&)
            
            ' If we still have working orders, call for a refresh and wait for it to finish...
            If g.Broker.HasWorkingOrders(lAccountID, vSymbolOrSymbolID, lAutoTradeItemID) Then
                g.Broker.BrokerDebug nBroker, "Flatten: Requesting a refresh because there are still working orders"
                g.Broker.Refresh nBroker
                
                lTimeOut = 0&
                Do
                    Sleep 1&
                    lTimeOut = lTimeOut + 1&
                Loop While g.Broker.Refreshing(nBroker) And (lTimeOut < 10&)
            
                ' If we still have working orders, warn the user and get out...
                If g.Broker.HasWorkingOrders(lAccountID, vSymbolOrSymbolID, lAutoTradeItemID) Then
                    g.Broker.BrokerDebug nBroker, "Flatten: Aborting because there are still working orders after a refresh"
                    InfBox "Timed out while attempting to Cancel Orders for " & strSymbol & " in account '" & g.Broker.AccountNameForID(lAccountID) & "'", "!", , "Flatten for " & strSymbol
                    bContinue = False
                End If
            End If
            
            If bContinue Then
                ' If in a position, confirm and submit market order to get out of position...
                If g.Broker.IsPitSymbol(lAccountID, vSymbolOrSymbolID) Then
                    g.Broker.BrokerDebug nBroker, "Flatten: Aborting because this is a pit session symbol"
                    InfBox "Trade Navigator has cancelled your working orders, but cannot flatten a position on a pit-session symbol.", , "!", "Flatten for " & strSymbol
                ElseIf g.Broker.PositionMatch(lAccountID, vSymbolOrSymbolID) Then
                    lPosition = g.Broker.CurrentPosition(lAccountID, vSymbolOrSymbolID, lAutoTradeItemID)
                    If lPosition <> 0 Then
                        If g.Broker.FlattenPosition(lAccountID, vSymbolOrSymbolID, lAutoTradeItemID, True) Then
                            lTimeOut = 0&
                            Do
                                Sleep 1&
                                lTimeOut = lTimeOut + 1&
                                lPosition = g.Broker.CurrentPosition(lAccountID, vSymbolOrSymbolID, lAutoTradeItemID)
                            Loop While (lPosition <> 0&) And (lTimeOut < 10&)
                            
                            If lPosition <> 0& Then
                                g.Broker.BrokerDebug nBroker, "Flatten: Requesting a refresh because still not flat"
                                g.Broker.Refresh nBroker
                                
                                lTimeOut = 0&
                                Do
                                    Sleep 1&
                                    lTimeOut = lTimeOut + 1&
                                Loop While g.Broker.Refreshing(nBroker) And (lTimeOut < 10&)
                            
                                ' If we are still in a position, warn the user and get out...
                                lPosition = g.Broker.CurrentPosition(lAccountID, vSymbolOrSymbolID, lAutoTradeItemID)
                                If lPosition <> 0& Then
                                    g.Broker.BrokerDebug nBroker, "Flatten: Aborting because still not flat after a refresh"
                                    InfBox "Timed out while attempting to Flatten position for " & strSymbol & " in account '" & g.Broker.AccountNameForID(lAccountID) & "'", "!", , "Flatten for " & strSymbol
                                End If
                            End If
                        End If
                    End If
                Else
                    g.Broker.BrokerDebug nBroker, "Flatten: Aborting because of a position mismatch"
                    InfBox "Because of inconsistent position data from the broker for " & strSymbol & " in account '" & g.Broker.AccountNameForID(lAccountID) & "', we cannot flatten your position at this time.||PLEASE CALL YOUR BROKER AND VERIFY YOUR POSITION IN THIS ACCOUNT.|", "!", , "Flatten for " & strSymbol
                End If
            End If
        End If
        
        strInProcess = Replace(strInProcess, strKey, "")
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.FlattenForSymbol"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExitPositionForSymbol
'' Description: Exit the position for a symbol in an account
'' Inputs:      Account, Symbol, Auto Trade Item ID, Position, Genesis Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExitPositionForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAutoTradeItemID As Long, ByVal lPosition As Long, Optional strGenesisOrderID As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim Order As cPtOrder               ' Order object
    Dim bConfirm As Boolean             ' Confirm the order?
    Dim strAccount As String            ' Account number
    Dim nBroker As eTT_AccountType      ' Broker
    
    bReturn = False
    If lPosition <> 0 Then
        bConfirm = False
        If Not g.CattleBridge Is Nothing Then
            strAccount = g.Broker.GetAccountNumber(vAccountNumberOrID)
            nBroker = g.Broker.AccountTypeForNumber(strAccount)
            
            bConfirm = g.CattleBridge.ConfirmOrder(strAccount, nBroker)
        End If
        
        vSymbolOrSymbolID = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
        Set Order = CreateMarketOrder(vAccountNumberOrID, vSymbolOrSymbolID, (lPosition < 0), Abs(lPosition), lAutoTradeItemID, False)
        
        If bConfirm Then
            bReturn = (CreateOrder(, , , Order) = eGDEditOrderReturn_Submit)
        Else
            Order.Save
            
            strGenesisOrderID = Order.GenesisOrderID
            
            SubmitOrder Order
            
            bReturn = True
        End If
    End If
    
    ExitPositionForSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ExitPositionForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReversePositionForSymbol
'' Description: Reverse the position for a symbol in an account
'' Inputs:      Account, Symbol, Auto Trade Item ID, Position, Genesis Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReversePositionForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAutoTradeItemID As Long, ByVal lPosition As Long, Optional strGenesisOrderID As String)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    
    If lPosition <> 0 Then
        vSymbolOrSymbolID = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
        Set Order = CreateMarketOrder(vAccountNumberOrID, vSymbolOrSymbolID, (lPosition < 0), Abs(lPosition * 2), lAutoTradeItemID, False)
        Order.Save
        
        strGenesisOrderID = Order.GenesisOrderID
        
        SubmitOrder Order
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ReversePositionForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateMarketOrder
'' Description: Create a market order given the arguments
'' Inputs:      Account, Symbol, Buy, Quantity, Auto Trade Item ID, Enter Position?
'' Returns:     Created Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateMarketOrder(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal bBuy As Boolean, ByVal lQuantity As Long, Optional ByVal lAutoTradeItemID As Long = 0&, Optional ByVal bEnter As Boolean = False) As cPtOrder
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    
    Set Order = New cPtOrder
    With Order
        .AccountID = g.Broker.GetAccountID(vAccountNumberOrID)
        If lAutoTradeItemID <= 0& Then
            .AutoTradeItemID = 0&
        Else
            .AutoTradeItemID = lAutoTradeItemID
        End If
        .SymbolOrSymbolID = vSymbolOrSymbolID
        .Buy = bBuy
        .Enter = bEnter
        .Expiration = -1&
        .GenesisOrderID = NextGenesisOrderID(g.Broker.GetAccountNumber(vAccountNumberOrID))
        .OrderDate = .BrokerDate(CurrentTime("", Order.Symbol))
        .OrderType = eTT_OrderType_Market
        .Quantity = lQuantity
    End With
    
    Set CreateMarketOrder = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.CreateMarketOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateLimitOrder
'' Description: Create a limit order given the arguments
'' Inputs:      Account, Symbol, Buy, Quantity, Price, Auto Trade Item ID, Enter Position?
'' Returns:     Created Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateLimitOrder(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal bBuy As Boolean, ByVal lQuantity As Long, ByVal dLimitPrice As Double, Optional ByVal lAutoTradeItemID As Long = 0&, Optional ByVal bEnter As Boolean = False) As cPtOrder
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    
    Set Order = New cPtOrder
    With Order
        .AccountID = g.Broker.GetAccountID(vAccountNumberOrID)
        If lAutoTradeItemID <= 0& Then
            .AutoTradeItemID = 0&
        Else
            .AutoTradeItemID = lAutoTradeItemID
        End If
        .SymbolOrSymbolID = vSymbolOrSymbolID
        .Buy = bBuy
        .Enter = bEnter
        .Expiration = -1&
        .GenesisOrderID = NextGenesisOrderID(g.Broker.GetAccountNumber(vAccountNumberOrID))
        .LimitPrice = dLimitPrice
        .OrderDate = .BrokerDate(CurrentTime("", Order.Symbol))
        .OrderType = eTT_OrderType_Limit
        .Quantity = lQuantity
    End With
    
    Set CreateLimitOrder = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.CreateLimitOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateStopOrder
'' Description: Create a stop order given the arguments
'' Inputs:      Account, Symbol, Buy, Quantity, Price, Auto Trade Item ID,
''              Enter Position?, Limit Price
'' Returns:     Created Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateStopOrder(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal bBuy As Boolean, ByVal lQuantity As Long, ByVal dStopPrice As Double, Optional ByVal lAutoTradeItemID As Long = 0&, Optional ByVal bEnter As Boolean = False, Optional ByVal dLimitPrice As Double = kNullData) As cPtOrder
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    
    Set Order = New cPtOrder
    With Order
        .AccountID = g.Broker.GetAccountID(vAccountNumberOrID)
        If lAutoTradeItemID <= 0& Then
            .AutoTradeItemID = 0&
        Else
            .AutoTradeItemID = lAutoTradeItemID
        End If
        .SymbolOrSymbolID = vSymbolOrSymbolID
        .Buy = bBuy
        .Enter = bEnter
        .Expiration = -1&
        .GenesisOrderID = NextGenesisOrderID(g.Broker.GetAccountNumber(vAccountNumberOrID))
        .OrderDate = .BrokerDate(CurrentTime("", Order.Symbol))
        .StopPrice = dStopPrice
        If dLimitPrice = kNullData Then
            .OrderType = eTT_OrderType_Stop
        Else
            .LimitPrice = dLimitPrice
            .OrderType = eTT_OrderType_StopWithLimit
        End If
        .Quantity = lQuantity
    End With
    
    Set CreateStopOrder = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.CreateStopOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnterPositionForSymbol
'' Description: Open the given position in the given account for the symbol
'' Inputs:      Account, Symbol, Auto Trade Item ID, Position, Genesis Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EnterPositionForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAutoTradeItemID As Long, ByVal lPosition As Long, Optional strGenesisOrderID As String)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    
    If lPosition <> 0& Then
        vSymbolOrSymbolID = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
        Set Order = CreateMarketOrder(vAccountNumberOrID, vSymbolOrSymbolID, (lPosition > 0), Abs(lPosition), lAutoTradeItemID, True)
        Order.Save
        
        strGenesisOrderID = Order.GenesisOrderID
        
        SubmitOrder Order
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.EnterPositionForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIDChanged
'' Description: If an ID changes on an order (like in the case of a cancel/
''              replace), update any triggered or cancelled orders that used
''              the old order ID
'' Inputs:      Old Order ID, New Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderIDChanged(ByVal lOldOrderID As Long, ByVal lNewOrderID As Long)
On Error GoTo ErrSection:

    Dim Orders As cGdTree               ' Collection of orders
    Dim lIndex As Long                  ' Index into a for loop
    Dim NewOrder As cPtOrder            ' New Order
    
    ' If the old order was an auto exit order, we will need to update the
    ' auto exit appropriately...
    If Not g.OrderStrategies Is Nothing Then
        g.OrderStrategies.OrderIDChanged lOldOrderID, lNewOrderID
    End If
    
    ' Next, we need to update any triggered orders so that they will now be
    ' triggered by the new order...
    Set Orders = g.Broker.TriggeredOrdersForOrderID(lOldOrderID)
    For lIndex = 1 To Orders.Count
        Orders(lIndex).TriggerOrderID = lNewOrderID
        Orders(lIndex).Save
        
        g.Broker.AddOrder Orders(lIndex)
        OrderCallback Orders(lIndex)
    Next lIndex

    ' We also need to update any OCO's to be linked to the new order instead of
    ' the old order...
    Set Orders = g.Broker.CancelOrdersForOrderID(lOldOrderID)
    For lIndex = 1 To Orders.Count
        Orders(lIndex).CancelOrderID = lNewOrderID
        Orders(lIndex).Save
        
        g.Broker.AddOrder Orders(lIndex)
        OrderCallback Orders(lIndex)
    Next lIndex
    
    ' Notify the conditional orders collection as well...
    If Not g.CondOrders Is Nothing Then
        g.CondOrders.OrderIDChanged lOldOrderID, lNewOrderID
    End If
    
    ' Also notify any Broker OCO links about the change...
'    If Not g.OrderLinks Is Nothing Then
'        g.OrderLinks.OrderIDChanged lOldOrderID, lNewOrderID
'    End If
    
    ' We will also need to copy any order journal entries from the old order
    ' to the new order (if the journal form is up for the old order ID, send
    ' the new order ID to it so that on a save all the journal entries get
    ' saved under the new ID)...
    g.JournalBridge.OrderIDChanged lOldOrderID, lNewOrderID
    
    If Not g.TsoGroups Is Nothing Then
        Set NewOrder = New cPtOrder
        If NewOrder.Load(lNewOrderID) Then
            g.TsoGroups.RefreshOrder NewOrder, lOldOrderID
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.OrderIDChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetExitOrderStrategies
'' Description: Get a list of the current exit order strategies
'' Inputs:      None
'' Returns:     List of Exit Order Strategies
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetExitOrderStrategies() As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As New cGdArray      ' List of exit order strategies to return
    Dim astrFiles As New cGdArray       ' Array of matching files
    Dim lIndex As Long                  ' Index into a for loop
    
    astrFiles.GetMatchingFiles AddSlash(App.Path) & "Provided\*.XOS", False
    For lIndex = 0 To astrFiles.Size - 1
        astrReturn.Add FileBase(astrFiles(lIndex)) & vbTab & "Provided\" & astrFiles(lIndex)
    Next lIndex

    astrFiles.GetMatchingFiles AddSlash(App.Path) & "Custom\*.XOS", False
    For lIndex = 0 To astrFiles.Size - 1
        astrReturn.Add FileBase(astrFiles(lIndex)) & vbTab & "Custom\" & astrFiles(lIndex)
    Next lIndex
    
    Set GetExitOrderStrategies = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.GetExitOrderStrategies"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExitOrderStrategyFileFromName
'' Description: Determine the exit order strategy filename from the name
'' Inputs:      Exit Order Strategy Name
'' Returns:     Exit Order Strategy Filename
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExitOrderStrategyFileFromName(ByVal strExitOrderStrategy As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from the function

    If FileExist(AddSlash(App.Path) & "Provided\" & strExitOrderStrategy & ".XOS") Then
        strReturn = "Provided\" & strExitOrderStrategy & ".XOS"
    ElseIf FileExist(AddSlash(App.Path) & "Custom\" & strExitOrderStrategy & ".XOS") Then
        strReturn = "Custom\" & strExitOrderStrategy & ".XOS"
    End If
    
    ExitOrderStrategyFileFromName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ExitOrderStrategyFileFromName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoTradeItemNameForID
'' Description: Determine the name for the automated trading item with the given ID
'' Inputs:      Auto Trading Item ID
'' Returns:     Auto Trading Item Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AutoTradeItemNameForID(ByVal lAutoTradeItemID As Long) As String
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim AutoTrade As cAutoTradeItem     ' Automated trading item
    Dim strReturn As String             ' Return value from the function

    If lAutoTradeItemID = 0 Then
        strReturn = ""
    Else
        Set AutoTrade = g.TradingItems.Item(Str(lAutoTradeItemID))
        If Not AutoTrade Is Nothing Then
            strReturn = AutoTrade.Name
        Else
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblAutoTradingItem] " & _
                        "WHERE [TradingItemID]=" & Str(lAutoTradeItemID) & ";", dbOpenDynaset)
            If Not (rs.BOF And rs.EOF) Then
                strReturn = rs!Name
            Else
                strReturn = "No Longer Exists"
            End If
        End If
    End If
    
    AutoTradeItemNameForID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.AutoTradeItemNameForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerInfoForm
'' Description: Retrieve the broker info form for the given broker
'' Inputs:      Broker
'' Returns:     Broker Info form if loaded, Nothing otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerInfoForm(ByVal nBroker As eTT_AccountType) As frmBrokerInfo
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the broker info form found?
    
    For lIndex = 0 To Forms.Count - 1
        If TypeOf Forms(lIndex) Is frmBrokerInfo Then
            If Forms(lIndex).Broker = nBroker Then
                Set BrokerInfoForm = Forms(lIndex)
                bFound = True
                Exit For
            End If
        End If
    Next lIndex
    
    If bFound = False Then
        Set BrokerInfoForm = Nothing
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.BrokerInfoForm"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeTrackerForm
'' Description: Retrieve the trade tracker form for the given account
'' Inputs:      Account
'' Returns:     Trade Tracker form if loaded, Nothing otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TradeTrackerForm(ByVal lAccountID As Long) As frmTTPositions
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the broker info form found?
    
    For lIndex = 0 To Forms.Count - 1
        If TypeOf Forms(lIndex) Is frmTTPositions Then
            If Forms(lIndex).AccountID = lAccountID Then
                Set TradeTrackerForm = Forms(lIndex)
                bFound = True
                Exit For
            End If
        End If
    Next lIndex
    
    If bFound = False Then
        Set TradeTrackerForm = Nothing
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.TradeTrackerForm"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CombineOrders
'' Description: Combine two orders that end up at the same price
'' Inputs:      Order being Changed, Order at the Price
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CombineOrders(OrderToChange As cPtOrder, OrderAtPrice As cPtOrder)
On Error GoTo ErrSection:
    
    If OrderToChange.Buy = OrderAtPrice.Buy Then
        CancelOrder OrderToChange, False
        ModifyOrder OrderAtPrice, , OrderAtPrice.RemainingQuantity + OrderToChange.RemainingQuantity, False
    ElseIf OrderAtPrice.RemainingQuantity > OrderToChange.RemainingQuantity Then
        CancelOrder OrderToChange, False
        ModifyOrder OrderAtPrice, , OrderAtPrice.RemainingQuantity - OrderToChange.RemainingQuantity, False
    ElseIf OrderAtPrice.RemainingQuantity < OrderToChange.RemainingQuantity Then
        CancelOrder OrderAtPrice, False
        ModifyOrder OrderToChange, OrderAtPrice.OrderPrice(True), OrderToChange.RemainingQuantity - OrderAtPrice.RemainingQuantity, False
    Else
        CancelOrder OrderAtPrice, False
        CancelOrder OrderToChange, False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.CombineOrders"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ModifyOrder
'' Description: Modify the given order
'' Inputs:      Order, New Price, New Quantity, Confirm?, New Order, Allow merge?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ModifyOrder(Order As cPtOrder, Optional ByVal dNewPrice As Double = 0#, Optional ByVal lNewQuantity As Long = 0&, Optional ByVal bConfirm As Boolean = True, Optional ByVal NewOrder As cPtOrder = Nothing, Optional ByVal bAllowMerge As Boolean = True)
On Error GoTo ErrSection:

    Dim OriginalOrder As New cPtOrder   ' Original order
    Dim OrderToEdit As New cPtOrder     ' Order to edit
    Dim bNewOrder As Boolean            ' Is this a new order?
    Dim nReturn As eGDEditOrderReturn   ' Return from the edit order call
        
    If OriginalOrder.Load(Order.OrderID) = False Then
        Set OriginalOrder = Order.MakeCopy
    End If
        
    bNewOrder = False
    If NewOrder Is Nothing Then
        Set OrderToEdit = SetupOrderToEdit(OriginalOrder, bNewOrder)
        
        If dNewPrice <> 0# Then
            OrderToEdit.OrderPrice(True) = dNewPrice
        End If
        If lNewQuantity <> 0& Then
            OrderToEdit.Quantity = lNewQuantity
        End If
    Else
        Set OrderToEdit = NewOrder
    End If

    ' 4) Show order form if appropriate...
    If CanMoveOrder(OrderToEdit, , Not bConfirm) = False Then
        nReturn = eGDEditOrderReturn_Cancel
    ElseIf bConfirm Then
        nReturn = frmTTEditOrder.ShowMe(OrderToEdit, Order.Buy, eGDTTEditOrderMode_Normal)
    Else
        nReturn = eGDEditOrderReturn_Submit
    End If
    
    ' 5) Do the appropriate thing with the order based on the return code...
    Select Case nReturn
        Case eGDEditOrderReturn_Park
            ParkOrder OrderToEdit
            AdjustTriggers OriginalOrder, OrderToEdit
        
        Case eGDEditOrderReturn_Submit
            SubmitAmend OriginalOrder, OrderToEdit, bAllowMerge
        
        Case eGDEditOrderReturn_Cancel
            If bNewOrder Then
                OrderToEdit.Delete
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ModifyOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrder
'' Description: Allow the user to submit the given order
'' Inputs:      Order to Submit, Previous ID, Called By ID, Submit All mode?,
''              Ask about other side?, Allow merge?
'' Returns:     Order ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SubmitOrder(Order As cPtOrder, Optional ByVal strPrevGenesisOrderID As String = "", Optional ByVal lCalledByID As Long = 0&, Optional ByVal bSubmitAll As Boolean = False, Optional ByVal bAskUserAboutOtherside As Boolean = True, Optional ByVal bAllowMerge As Boolean = True) As Long
On Error GoTo ErrSection:

    Dim OrdersAtPrice As cGdTree        ' Collection of other orders at the given price
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' An order was found to combine with
    Dim OrderAtPrice As cPtOrder        ' Temporary order object
    Dim dDiff As Double                 ' Difference between the stop and limit price
    Dim OtherOrder As cPtOrder          ' Other side of an OCO
    Dim strReturn As String             ' Return value from an InfBox
    Dim lReturn As Long                 ' Return value for the function
    Dim lBrokerCancelOrderID As Long    ' Broker cancel order ID
    Dim OrderToEdit As cPtOrder         ' Order to edit
    Dim frm As frmAlertPopup            ' Alert popup form
    
    lReturn = 0&
    
    strReturn = "S"
    If ((g.Broker.IsLiveAccount(Order.Broker) = True) Or (Order.Broker = eTT_AccountType_SimBroker)) And (g.RealTime.SymbolDelay(Order.Symbol) <> 0) Then
        strReturn = InfBox("You are about to submit an order to your broker based on delayed data in Trade Navigator.||Do you want to continue?", "?", "+Submit Order|Park Order|-Cancel Order", "Warning")
        g.Broker.BrokerDebug Order.Broker, "User answered '" & strReturn & "' to warning about submitting the order despite delayed data"
    End If
    
    If strReturn = "S" Then
        If Order.Quantity = 0 Then
            Order.Message = "Invalid Quantity"
            Order.ChangeOrderStatus eTT_OrderStatus_Rejected
        
            Set frm = New frmAlertPopup
            frm.ShowMessageBox Order.OrderText & " REJECTED||Invalid Quantity", "Order Rejected", vbLeftJustify
            
        ElseIf (Order.OrderID > 0) And (g.OrderStrategies.OrderExistsInStrategy(Order)) Then
            lReturn = g.Broker.SendOrder(Order, strPrevGenesisOrderID)
        ElseIf g.TsoGroups.OrderExistsInGroup(Order) Then
            ' 10/07/2010 DAJ: If the order is part of an active TradeSense order
            ' group, don't try to merge it with anything...
            lReturn = g.Broker.SendOrder(Order, strPrevGenesisOrderID, , bSubmitAll)
        Else
            ' DAJ 08/14/2009: If this is an order that was parked as one side of an order-
            ' cancel-order situation, we need to see if the user wants to submit the other
            ' side of the OCO as well, break the OCO link, or cancel the submission altogether...
            If (Order.AutoTradeItemID = 0&) And (Order.IsAutoExit = False) Then
                If (Order.CancelOrderID <> 0) And (Order.CancelOrderID <> lCalledByID) Then
                    Set OtherOrder = New cPtOrder
                    If OtherOrder.Load(Order.CancelOrderID) Then
                        ''If (Order.Status = eTT_OrderStatus_Parked) Or (OtherOrder.Status = eTT_OrderStatus_Parked) Then
                        If OtherOrder.Status = eTT_OrderStatus_Parked Then
                            ' If we are in a submit all situation, then assume Submit here, otherwise
                            ' ask the user whether to submit or break the OCO...
                            If (bSubmitAll = True) Or (bAskUserAboutOtherside = False) Then
                                strReturn = "S"
                            Else
                                strReturn = InfBox("You have chosen to submit one side of an Order-Cancel-Order.  What would you like to do with the other one?", "?", "+Submit|Break OCO|-Cancel", "Order Cancel Order")
                            End If
                            Select Case strReturn
                                Case "S"
                                    ' DAJ 08/27/2009: First need to check to make sure that both sides of the
                                    ' OCO will be on the correct side of the market...
                                    If CanSubmitOCO(Order, OtherOrder) = True Then
                                        Order.CancelOrderID = SubmitOrder(OtherOrder, strPrevGenesisOrderID, Order.OrderID, bSubmitAll, bAskUserAboutOtherside)
                                    Else
                                        Exit Function
                                    End If
                                    
                                Case "B"
                                    Order.CancelOrderID = 0
                                
                                Case "C"
                                    Exit Function
                            End Select
                        End If
                    End If
                ElseIf (Order.BrokerCancelOrderID <> 0) And (Abs(Order.BrokerCancelOrderID) <> lCalledByID) Then
                    Set OtherOrder = New cPtOrder
                    If OtherOrder.Load(Abs(Order.BrokerCancelOrderID)) Then
                        If OtherOrder.Status = eTT_OrderStatus_Parked Then
                            ' If we are in a submit all situation, then assume Submit here, otherwise
                            ' ask the user whether to submit or break the OCO...
                            If (bSubmitAll = True) Or (bAskUserAboutOtherside = False) Then
                                strReturn = "S"
                            Else
                                strReturn = InfBox("You have chosen to submit one side of an Order-Cancel-Order.  What would you like to do with the other one?", "?", "+Submit|Break OCO|-Cancel", "Order Cancel Order")
                            End If
                            Select Case strReturn
                                Case "S"
                                    ' DAJ 08/27/2009: First need to check to make sure that both sides of the
                                    ' OCO will be on the correct side of the market...
                                    If CanSubmitOCO(Order, OtherOrder) = True Then
                                        lBrokerCancelOrderID = SubmitOrder(OtherOrder, strPrevGenesisOrderID, Order.OrderID, bSubmitAll, bAskUserAboutOtherside)
                                        If Order.BrokerCancelOrderID < 0 Then
                                            If lBrokerCancelOrderID > 0 Then
                                                lBrokerCancelOrderID = lBrokerCancelOrderID * -1&
                                            End If
                                        End If
                                        Order.BrokerCancelOrderID = lBrokerCancelOrderID
                                    Else
                                        Exit Function
                                    End If
                                    
                                Case "B"
'                                    If Order.Broker = eTT_AccountType_PFG Then
'                                        If g.OrderLinks.UnlinkAndSubmitOne(Order) = True Then
'                                            Exit Function
'                                        Else
'                                            Order.BrokerCancelOrderID = 0&
'                                        End If
'                                    Else
                                        Order.BrokerCancelOrderID = 0&
'                                    End If
                                
                                Case "C"
                                    Exit Function
                            End Select
                        End If
                    End If
                End If
            End If
            
            If bAllowMerge = True And Order.OrderPrice(True) <> kNullData Then
                Set OrdersAtPrice = g.Broker.PrimaryOrdersForSymbol(Order.AccountID, Order.SymbolID, Order.AutoTradeItemID, Order.OrderPrice(True))
                If Not OrdersAtPrice Is Nothing Then
                    ' 05/11/2010 DAJ: Remove any orders that are conditional, have an OCO link,
                    ' or order triggers so that they don't get merged...
                    ' 05/20/2010 DAJ: Also don't try to merge with an order that is currently
                    ' in a pending status...
                    ' 10/07/2010 DAJ: Also don't try to merge with an order that is part of an
                    ' active TradeSense order group...
                    For lIndex = OrdersAtPrice.Count To 1 Step -1
                        If OrdersAtPrice(lIndex).HasTriggeredOrders Or OrdersAtPrice(lIndex).HasOcoOrders Or OrdersAtPrice(lIndex).HasTrigger Then
                            OrdersAtPrice.Remove lIndex
                        ElseIf g.TsoGroups.OrderExistsInGroup(OrdersAtPrice(lIndex)) Then
                            OrdersAtPrice.Remove lIndex
                        ElseIf OrderIsPending(OrdersAtPrice(lIndex)) Then
                            OrdersAtPrice.Remove lIndex
                        End If
                    Next lIndex
                End If
            End If
            
            ' 05/06/2010 DAJ: If the order has a trigger (either it is a conditional order or a
            ' "triggered by" order), change the order status to "Trigger Pending" instead of
            ' actually submitting it (Issues #5715, #5721)...
            If Order.HasTrigger Then
                SetTriggerOrderStatus Order
                
                ' 06/25/2010 DAJ: Need to fix any negative OTOs that are coming off of this order
                ' here so that they trigger correctly when this order fills (#5820)...
                ChangeNegativeOtos Order, Order.OrderID, bSubmitAll
                
                lReturn = Order.OrderID
            ElseIf bAllowMerge = False Then
                lReturn = g.Broker.SendOrder(Order, strPrevGenesisOrderID, , bSubmitAll)
            ElseIf Order.HasTriggeredOrders Or Order.HasOcoOrders Or Order.HasTrigger Then
                ' 05/06/2010 DAJ: If the order has orders triggered off of it or an
                ' OCO link, don't merge them together...
                lReturn = g.Broker.SendOrder(Order, strPrevGenesisOrderID, , bSubmitAll)
            ElseIf OrdersAtPrice Is Nothing Then
                lReturn = g.Broker.SendOrder(Order, strPrevGenesisOrderID, , bSubmitAll)
            ElseIf OrdersAtPrice.Count = 0 Then
                lReturn = g.Broker.SendOrder(Order, strPrevGenesisOrderID, , bSubmitAll)
            Else
                bFound = False
                For lIndex = 1 To OrdersAtPrice.Count
                    If g.OrderStrategies.OrderExistsInStrategy(OrdersAtPrice(lIndex)) = False Then
                        Set OrderAtPrice = OrdersAtPrice(lIndex)
                        If (OrderAtPrice.OrderID <> Order.OrderID) And (OrderAtPrice.GenesisOrderID <> strPrevGenesisOrderID) Then
                            If OrderAtPrice.HasTriggeredOrders Or OrderAtPrice.HasOcoOrders Or OrderAtPrice.HasTrigger Then
                                ' 05/06/2010 DAJ: If the order that is sitting at that price has
                                ' orders triggered off of it or an OCO link, don't merge them
                                ' together...
                                lReturn = g.Broker.SendOrder(Order, strPrevGenesisOrderID)
                            ElseIf OrderAtPrice.Buy = Order.Buy Then
                                ' 03/11/2010 DAJ: If the expiration is different (e.g. one is a day
                                ' order and one is a GTC order), then keep them separate (Issue #5644)...
                                If SameTif(OrderAtPrice, Order) = False Then
                                    lReturn = g.Broker.SendOrder(Order, strPrevGenesisOrderID)
                                Else
                                    CancelOrder Order, False
                                    ModifyOrder OrderAtPrice, , OrderAtPrice.RemainingQuantity + Order.RemainingQuantity, False
                                End If
                            Else
                                If OrderAtPrice.RemainingQuantity = Order.RemainingQuantity Then
                                    CancelOrder Order, False
                                    CancelOrder OrderAtPrice, False
                                Else
                                    If OrderAtPrice.RemainingQuantity > Order.RemainingQuantity Then
                                        CancelOrder Order, False
                                        OrderAtPrice.Quantity = OrderAtPrice.RemainingQuantity - Order.RemainingQuantity
                                        lReturn = g.Broker.SendOrder(OrderAtPrice, OrderAtPrice.GenesisOrderID)
                                    Else
                                        CancelOrder OrderAtPrice, False
                                        Order.Quantity = Order.RemainingQuantity - OrderAtPrice.RemainingQuantity
                                        Select Case Order.OrderType
                                            Case eTT_OrderType_Stop
                                                If OrderAtPrice.OrderType = eTT_OrderType_Limit Then
                                                    Order.StopPrice = OrderAtPrice.LimitPrice
                                                ElseIf OrderAtPrice.OrderType = eTT_OrderType_MIT Then
                                                    Order.StopPrice = OrderAtPrice.MitPrice
                                                Else
                                                    Order.StopPrice = OrderAtPrice.StopPrice
                                                End If
                                                
                                            Case eTT_OrderType_Limit
                                                If OrderAtPrice.OrderType = eTT_OrderType_Limit Then
                                                    Order.LimitPrice = OrderAtPrice.LimitPrice
                                                ElseIf OrderAtPrice.OrderType = eTT_OrderType_MIT Then
                                                    Order.LimitPrice = OrderAtPrice.MitPrice
                                                Else
                                                    Order.LimitPrice = OrderAtPrice.StopPrice
                                                End If
                                            
                                            Case eTT_OrderType_StopWithLimit
                                                dDiff = Order.StopPrice - Order.LimitPrice
                                                If OrderAtPrice.OrderType = eTT_OrderType_Limit Then
                                                    Order.StopPrice = OrderAtPrice.LimitPrice
                                                ElseIf OrderAtPrice.OrderType = eTT_OrderType_MIT Then
                                                    Order.StopPrice = OrderAtPrice.MitPrice
                                                Else
                                                    Order.StopPrice = OrderAtPrice.StopPrice
                                                End If
                                                Order.LimitPrice = Order.StopPrice - dDiff
                                                
                                            Case eTT_OrderType_MIT
                                                If OrderAtPrice.OrderType = eTT_OrderType_Limit Then
                                                    Order.MitPrice = OrderAtPrice.LimitPrice
                                                ElseIf OrderAtPrice.OrderType = eTT_OrderType_MIT Then
                                                    Order.MitPrice = OrderAtPrice.MitPrice
                                                Else
                                                    Order.MitPrice = OrderAtPrice.StopPrice
                                                End If
                                                
                                        End Select
                                        lReturn = g.Broker.SendOrder(Order, Order.GenesisOrderID)
                                    End If
                                End If
                            End If
                            
                            bFound = True
                            Exit For
                        End If
                    End If
                Next lIndex
                
                If bFound = False Then
                    lReturn = g.Broker.SendOrder(Order, strPrevGenesisOrderID)
                End If
            End If
        End If
    ElseIf strReturn = "P" Then
        Order.ChangeOrderStatus eTT_OrderStatus_Parked
    Else
        Order.ChangeOrderStatus eTT_OrderStatus_Cancelled
    End If
    
    SubmitOrder = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.SubmitOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StreamReplayAccount
'' Description: Return the account ID for the given account number
'' Inputs:      Account Number, Delete After Date
'' Returns:     Account ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StreamReplayAccount(ByVal strAccountNumber As String, ByVal dDeleteAfter As Double) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value from the function
    Dim rs As Recordset                 ' Recordset into the database
    Dim lLast As Long                   ' Last counter used for account number
    Dim lCounter As Long                ' Counter for the current record
    Dim strAccountName As String        ' Account name

    If Len(Trim(strAccountNumber)) = 0 Then
        strAccountNumber = "GenSr00001"
        strAccountName = "Replay00001"
    ElseIf UCase(strAccountNumber) = "NEW" Then
        lLast = 0&
        
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountType]=" & Str(eTT_AccountType_SimReplay) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            If Len(rs!AccountNumber) = 10 Then
                If UCase(Left(rs!AccountNumber, 5)) = "GENSR" Then
                    lCounter = CLng(Val(Right(rs!AccountNumber, 5)))
                    If lCounter > lLast Then
                        lLast = lCounter
                    End If
                End If
            End If
            rs.MoveNext
        Loop
        
        strAccountNumber = "GenSr" & Format(lLast + 1, "00000")
        strAccountName = "Replay" & Format(lLast + 1, "00000")
    End If
    
    lReturn = g.Broker.AccountIDForNumber(strAccountNumber)
    If lReturn = -1& Then
        lReturn = CreateAccountFromNumber(strAccountNumber, eTT_AccountType_SimReplay, strAccountName)
    End If
    
    g.SimTradeReplay.Broker.BrokerInfo.AddAccount strAccountNumber
    DeleteTransactionsAfterDate strAccountNumber, dDeleteAfter
        
    StreamReplayAccount = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.StreamReplayAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteTransactionsAfterDate
'' Description: Delete all orders and fills in an account after a given date and
''              recalculate the positions and account positions
'' Inputs:      Account ID, Date/Time
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteTransactionsAfterDate(ByVal vAccountNumberOrID As Variant, ByVal dDeleteAfterDate As Double)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lAccountID As Long              ' Account ID
    Dim nAcctType As eTT_AccountType    ' Account type for the account ID
    
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    nAcctType = g.Broker.AccountTypeForID(lAccountID)
    
    ' 1) Delete all of the orders with an order date greater than or equal to the
    '    given date...
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                "WHERE [AccountID]=" & Str(lAccountID) & " AND [StatusDate]>=" & Str(dDeleteAfterDate) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
    
    ' 2) Delete all of the fills with an order date greater than or equal to the
    '    given date...
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                "WHERE [AccountID]=" & Str(lAccountID) & " AND [FillDate]>=" & Str(dDeleteAfterDate) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
    
    ' 2) Rebuild the positions for the account...
    g.Broker.RebuildFillSummaryForAccount vAccountNumberOrID, True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.DeleteTransactionsAfterDate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PostFixVersion28
'' Description: Post fix the database tables for version 28 update
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PostFixVersion28()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim bNotified As Boolean            ' Has the customer been notified?
    Dim Bars As New cGdBars             ' Bars object
    Dim AcctPos As cAccountPosition     ' Account position object
    Dim lPrevSession As Long            ' Session previous to the last session date
    Dim bDoPass2 As Boolean             ' Do we want to perform pass 2 as well?
    Dim astrLastSession As cGdArray     ' Last session array
    Dim strKey As String                ' Key into the array
    Dim lPos As Long                    ' Position into the array
    
    Set astrLastSession = New cGdArray
    astrLastSession.Create eGDARRAY_Strings
    
    Set rs = g.dbPaper.OpenRecordset("SELECT tblOrders.*, tblAccounts.AccountType " & _
                "FROM tblOrders INNER JOIN tblAccounts ON tblOrders.AccountID=tblAccounts.AccountID " & _
                "WHERE tblOrders.SessionDate<0;", dbOpenDynaset)
    Do While Not rs.EOF
        If bNotified = False Then
            frmSplash.Message -1, "Applying update to TradeTracker.MDB ..."
            bNotified = True
        End If
        
        Set rs2 = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderLegs] " & _
                "WHERE [OrderID]=" & Str(rs!OrderID) & " AND [LegNumber]=1;", dbOpenDynaset)
        If Not (rs2.BOF And rs2.EOF) Then
            SetBarProperties Bars, rs2!Symbol
            
            rs.Edit
            rs!SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(rs!OrderDate, rs!AccountType, rs2!Symbol, False))
            strKey = Str(rs!AccountID) & vbTab
            If astrLastSession.BinarySearch(strKey, lPos, eGdSort_MatchUsingSearchStringLength) Then
                If rs!SessionDate > CLng(Val(Parse(astrLastSession(lPos), vbTab, 2))) Then
                    astrLastSession(lPos) = strKey & Str(rs!SessionDate)
                End If
            Else
                astrLastSession.Add strKey & Str(rs!SessionDate), lPos
            End If
            rs.Update
            bDoPass2 = True
        End If
        
        rs.MoveNext
    Loop
    
    If bDoPass2 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
        Do While Not rs.EOF
            If astrLastSession.BinarySearch(Str(rs!AccountID) & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
                lPrevSession = CLng(Val(Parse(astrLastSession(lPos), vbTab, 2))) - 1&
                Do While IsWeekday(lPrevSession) = False
                    lPrevSession = lPrevSession - 1&
                Loop
            
                rs.Edit
                rs!IsSnapshot = (rs!SessionDate >= lPrevSession)
                rs.Update
            End If
            
            rs.MoveNext
        Loop
    End If
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                "WHERE [LastTraded]<0 AND [LastTradedSnapshot]<0;", dbOpenDynaset)
    Do While Not rs.EOF
        If bNotified = False Then
            frmSplash.Message -1, "Applying update to TradeTracker.MDB ..."
            bNotified = True
        End If
        
        Set AcctPos = New cAccountPosition
        If AcctPos.Load(rs!AccountPositionID, rs) Then
            AcctPos.RecalculateHistory
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.PostFixVersion28"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditOrderFromOrder
'' Description: Allow the user to edit the given order
'' Inputs:      Order, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EditOrderFromOrder(Order As cPtOrder, ByVal strSource As String)
On Error GoTo ErrSection:

    Dim bContinue As Boolean            ' Continue?

    bContinue = True
    If Order.AutoTradeItemID <> 0 Then
        If g.TradingItems.IsTradingItemActive(Order.AutoTradeItemID) = True Then
            InfBox "You cannot edit an order generated by an active automated trading item", "!", "+-OK", "Edit Order Error"
            bContinue = False
        End If
    End If
    
    If bContinue Then
        If HasBeenSent(Order.Status) Then
            If g.Broker.ConnectionStatusForAccount(Order.AccountID) <> eGDConnectionStatus_Connected Then
                InfBox "You cannot edit the order because you are not connected to the " & g.Broker.BrokerName(Order.Broker) & " servers", "!", "+-OK", "Edit Order Error"
                bContinue = False
            End If
        End If
        
        If bContinue Then
            g.Broker.BrokerDebug Order.Broker, "Modifying Order from " & strSource & ": " & Order.OrderText, True
            EditOrder Order
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.EditOrderFromOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditOrderFromGrid
'' Description: Allow the user to edit the selected order in the grid
'' Inputs:      Grid, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EditOrderFromGrid(Grid As VSFlexGrid, ByVal strSource As String)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to edit

    If (Grid.Row >= Grid.FixedRows) And (Grid.Row < Grid.Rows) Then
        If TypeOf Grid.RowData(Grid.Row) Is cPtOrder Then
            Set Order = Grid.RowData(Grid.Row)
            If Order.NumberOfLegs = 1 Then
                EditOrderFromOrder Order, strSource
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.EditOrderFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditOrderFromID
'' Description: Allow the user to edit the order with the given order ID
'' Inputs:      Order ID, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EditOrderFromID(ByVal lOrderID As Long, ByVal strSource As String)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to edit
    
    Set Order = New cPtOrder
    If Order.Load(lOrderID) Then
        EditOrderFromOrder Order, strSource
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.EditOrderFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrderFromOrder
'' Description: Cancel the given order
'' Inputs:      Order, Source, User Cancel?, Confirm Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelOrderFromOrder(Order As cPtOrder, ByVal strSource As String, Optional ByVal bUserCancel As Boolean = False, Optional ByVal bConfirm As Boolean = True)
On Error GoTo ErrSection:

    Dim bContinue As Boolean            ' Continue?

    bContinue = True
    If Order.AutoTradeItemID <> 0 Then
        If g.TradingItems.IsTradingItemActive(Order.AutoTradeItemID) = True Then
            InfBox "You cannot cancel an order generated by an active automated trading item", "!", "+-OK", "Cancel Order Error"
            bContinue = False
        End If
    End If
    
    If bContinue Then
        If HasBeenSent(Order.Status) Then
            If g.Broker.ConnectionStatusForAccount(Order.AccountID) <> eGDConnectionStatus_Connected Then
                InfBox "You cannot cancel the order because you are not connected to the" & g.Broker.BrokerName(Order.Broker) & " servers", "!", "+-OK", "Cancel Order Error"
                bContinue = False
            End If
        End If
            
        If bContinue Then
            g.Broker.BrokerDebug Order.Broker, "Cancelling Order '" & Order.OrderText(True, True) & "' (" & Order.GenesisOrderID & ", " & Order.BrokerID & ") from " & strSource, True
            CancelOrder Order, bConfirm, , bUserCancel
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.CancelOrderFromOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrderFromGrid
'' Description: Cancel the selected open order from the grid
'' Inputs:      Grid, Source, User Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelOrderFromGrid(Grid As VSFlexGrid, ByVal strSource As String, Optional ByVal bUserCancel As Boolean = False)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to cancel

    If (Grid.Row >= Grid.FixedRows) And (Grid.Row < Grid.Rows) Then
        If TypeOf Grid.RowData(Grid.Row) Is cPtOrder Then
            Set Order = Grid.RowData(Grid.Row)
            CancelOrderFromOrder Order, strSource, bUserCancel
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.CancelOrderFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrderFromID
'' Description: Cancel the order with the given order ID
'' Inputs:      Order ID, Source, User Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelOrderFromID(ByVal lOrderID As Long, ByVal strSource As String, Optional ByVal bUserCancel As Boolean = False)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to cancel

    Set Order = New cPtOrder
    If Order.Load(lOrderID) Then
        CancelOrderFromOrder Order, strSource, bUserCancel
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.CancelOrderFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrderFromOrder
'' Description: Park the given order
'' Inputs:      Order, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ParkOrderFromOrder(Order As cPtOrder, ByVal strSource As String)
On Error GoTo ErrSection:

    Dim bContinue As Boolean            ' Continue?

    ' DAJ 04/02/2013: Need to allow trigger pending orders to be parked this way as well ( #6808 )
    If (IsOpenOrder(Order.Status, False) = True) Or (Order.Status = eTT_OrderStatus_TriggerPending) Or (Order.Status = eTT_OrderStatus_DataPending) Then
        bContinue = True
        If Order.AutoTradeItemID <> 0 Then
            If g.TradingItems.IsTradingItemActive(Order.AutoTradeItemID) = True Then
                InfBox "You cannot park an order generated by an active automated trading item", "!", "+-OK", "Park Order Error"
                bContinue = False
            End If
        End If
        
        If bContinue Then
            If g.Broker.ConnectionStatusForAccount(Order.AccountID) <> eGDConnectionStatus_Connected Then
                InfBox "You cannot park the order because you are not connected to the" & g.Broker.BrokerName(Order.Broker) & " servers", "!", "+-OK", "Park Order Error"
            Else
                g.Broker.BrokerDebug Order.Broker, "Parking Order from " & strSource & ": " & Order.OrderText, True
                ParkOrder Order
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ParkOrderFromOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrderFromGrid
'' Description: Park the selected open order from the grid
'' Inputs:      Grid, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ParkOrderFromGrid(Grid As VSFlexGrid, ByVal strSource As String)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to Park
    Dim bContinue As Boolean            ' Continue?

    If (Grid.Row >= Grid.FixedRows) And (Grid.Row < Grid.Rows) Then
        If TypeOf Grid.RowData(Grid.Row) Is cPtOrder Then
            Set Order = Grid.RowData(Grid.Row)
            ParkOrderFromOrder Order, strSource
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ParkOrderFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrderFromID
'' Description: Park the order with the given Order ID
'' Inputs:      Order ID, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ParkOrderFromID(ByVal lOrderID As Long, ByVal strSource As String)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to Park
    
    Set Order = New cPtOrder
    If Order.Load(lOrderID) Then
        ParkOrderFromOrder Order, strSource
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ParkOrderFromID"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrderFromOrder
'' Description: Submit the given order
'' Inputs:      Order, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SubmitOrderFromOrder(Order As cPtOrder, ByVal strSource As String)
On Error GoTo ErrSection:

    Dim bContinue As Boolean            ' Continue?

    If Order.Status = eTT_OrderStatus_Parked Then
        bContinue = True
        If Order.AutoTradeItemID <> 0 Then
            If g.TradingItems.IsTradingItemActive(Order.AutoTradeItemID) = True Then
                InfBox "You cannot submit an order generated by an active automated trading item", "!", "+-OK", "Submit Order Error"
                bContinue = False
            End If
        End If
        
        If bContinue Then
            If g.Broker.ConnectionStatusForAccount(Order.AccountID) <> eGDConnectionStatus_Connected Then
                InfBox "You cannot submit the order because you are not connected to the" & g.Broker.BrokerName(Order.Broker) & " servers", "!", "+-OK", "Submit Order Error"
            Else
                g.Broker.BrokerDebug Order.Broker, "Submitting Order from " & strSource & ": " & Order.OrderText, True
                SubmitOrder Order
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SubmitOrderFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrderFromGrid
'' Description: Submit the selected parked order from the grid
'' Inputs:      Grid, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SubmitOrderFromGrid(Grid As VSFlexGrid, ByVal strSource As String)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to submit

    If (Grid.Row >= Grid.FixedRows) And (Grid.Row < Grid.Rows) Then
        If TypeOf Grid.RowData(Grid.Row) Is cPtOrder Then
            Set Order = Grid.RowData(Grid.Row)
            SubmitOrderFromOrder Order, strSource
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SubmitOrderFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrderFromID
'' Description: Submit the order with the given order ID
'' Inputs:      Order ID, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SubmitOrderFromID(ByVal lOrderID As Long, ByVal strSource As String)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to submit
    
    Set Order = New cPtOrder
    If Order.Load(lOrderID) Then
        SubmitOrderFromOrder Order, strSource
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SubmitOrderFromID"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitAllOrdersFromGrid
'' Description: Submit all the parked orders from the grid
'' Inputs:      Grid, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SubmitAllOrdersFromGrid(Grid As VSFlexGrid, ByVal strSource As String)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to submit
    Dim lIndex As Long                  ' Index into a for loop
    Dim strIgnoreIds As String          ' String of ID's to ignore
    Dim strSubmitted As String          ' String of ID's of submitted orders
    Dim Orders As New cGdTree           ' Collection of orders to submit
    
    With Grid
        ' Pass One: Submit any orders that are not Triggered by any other orders...
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cPtOrder Then
                Set Order = .RowData(lIndex)
                
                If InStr(strIgnoreIds, "," & Str(Order.OrderID) & ",") = 0 Then
                    ' If there is an order-cancel-order situation here, submitting one
                    ' order will submit the other one, so we want to make sure not to
                    ' submit the other one here...
                    If Order.CancelOrderID <> 0 Then
                        strIgnoreIds = strIgnoreIds & "," & Str(Order.CancelOrderID) & ","
                    End If
                
                    If g.Broker.ConnectionStatusForAccount(Order.AccountID) = eGDConnectionStatus_Connected Then
                        If (Order.Status = eTT_OrderStatus_Parked) And (Order.AutoTradeItemID = 0&) Then
                            If Order.TriggerOrderID = 0& Then
                                Orders.Add Order, Str(Order.OrderID)
                            End If
                        End If
                    End If
                End If
            End If
        Next lIndex
        
        For lIndex = 1 To Orders.Count
            g.Broker.BrokerDebug Orders(lIndex).Broker, "Submitting Order from " & strSource & ": " & Orders(lIndex).OrderText, True
            SubmitOrder Orders(lIndex), , , True
            strSubmitted = strSubmitted & ",-" & Str(Orders(lIndex).OrderID) & ","
        Next lIndex
    
        ' Pass Two: Submit any orders that are Triggered by other orders...
        Orders.Clear
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cPtOrder Then
                Set Order = .RowData(lIndex)
                
                If InStr(strIgnoreIds, "," & Str(Order.OrderID) & ",") = 0 Then
                    If g.Broker.ConnectionStatusForAccount(Order.AccountID) = eGDConnectionStatus_Connected Then
                        If (Order.Status = eTT_OrderStatus_Parked) And (Order.AutoTradeItemID = 0&) Then
                            If Order.TriggerOrderID <> 0& Then
                                If InStr("," & strSubmitted & ",", "," & Str(Order.TriggerOrderID) & ",") = 0 Then
                                    If Order.Load(Order.OrderID) Then
                                        Orders.Add Order, Str(Order.OrderID)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next lIndex
    End With
    
    For lIndex = 1 To Orders.Count
        g.Broker.BrokerDebug Orders(lIndex).Broker, "Submitting Order from " & strSource & ": " & Orders(lIndex).OrderText, True
        SubmitOrder Orders(lIndex), , , True
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SubmitAllOrdersFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenPositionFromGrid
'' Description: Allow the user to flatten the selected position from the grid
'' Inputs:      Grid, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FlattenPositionFromGrid(Grid As VSFlexGrid, ByVal strSource As String)
On Error GoTo ErrSection:

    Dim AcctPos As cAccountPosition     ' Account position object
    Dim lIndex As Long                  ' Index into a for loop

    With Grid
        If (.Row >= .FixedRows) And (.Row < .Rows) Then
            If TypeOf .RowData(.Row) Is cAccountPosition Then
                Set AcctPos = .RowData(.Row)
                If AcctPos.AutoTradeItemID = -1& Then
                    If .GetNodeRow(.Row, flexNTFirstChild) <> -1& Then
                        For lIndex = .GetNodeRow(.Row, flexNTFirstChild) To .GetNodeRow(.Row, flexNTLastChild)
                            If TypeOf .RowData(lIndex) Is cAccountPosition Then
                                Set AcctPos = .RowData(lIndex)
                                g.Broker.BrokerDebug AcctPos.Broker, "Flattening Position for Account: " & g.Broker.AccountNumberForID(AcctPos.AccountID) & ", Symbol: " & AcctPos.Symbol & ", Auto Trade Item: " & Str(AcctPos.AutoTradeItemID) & " from " & strSource, True
                                FlattenForSymbol AcctPos.AccountID, AcctPos.SymbolOrSymbolID, AcctPos.AutoTradeItemID
                            End If
                        Next lIndex
                    End If
                Else
                    g.Broker.BrokerDebug AcctPos.Broker, "Flattening Position for Account: " & g.Broker.AccountNumberForID(AcctPos.AccountID) & ", Symbol: " & AcctPos.Symbol & ", Auto Trade Item: " & Str(AcctPos.AutoTradeItemID) & " from " & strSource, True
                    FlattenForSymbol AcctPos.AccountID, AcctPos.SymbolOrSymbolID, AcctPos.AutoTradeItemID
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.FlattenPositionFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReversePositionFromGrid
'' Description: Allow the user to reverse the selected position from the grid
'' Inputs:      Grid, Source
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReversePositionFromGrid(Grid As VSFlexGrid, ByVal strSource As String)
On Error GoTo ErrSection:

    Dim AcctPos As cAccountPosition     ' Account position object
    
    With Grid
        If (.Row >= .FixedRows) And (.Row < .Rows) Then
            If TypeOf .RowData(.Row) Is cAccountPosition Then
                Set AcctPos = .RowData(.Row)
                If AcctPos.AutoTradeItemID <> -1& Then
                    g.Broker.BrokerDebug AcctPos.Broker, "Reversing Position for Account: " & g.Broker.AccountNumberForID(AcctPos.AccountID) & ", Symbol: " & AcctPos.Symbol & ", Auto Trade Item: " & Str(AcctPos.AutoTradeItemID) & " from " & strSource, True
                    ReverseForSymbol AcctPos.AccountID, AcctPos.SymbolOrSymbolID, AcctPos.AutoTradeItemID
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ReversePositionFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllowAutoExits
'' Description: Determine whether to allow auto exits for an account/symbol or not
'' Inputs:      Account, Symbol, Force Verify?, Force Position Match?
'' Returns:     True if Allow Auto Exits, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllowAutoExits(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, Optional ByVal bForceVerify As Boolean = True, Optional ByVal bForcePositionMatch As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lAccountID As Long              ' Account ID for account passed in
    
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    
    If HasModule("RTG,RTE") = False Then
        bReturn = False
    ElseIf g.Broker.IsTradeableSymbol(vAccountNumberOrID, vSymbolOrSymbolID) = False Then
        bReturn = False
    ' 12/04/2014 DAJ: With the changes in exchange fees coming from the exchange, there is more of a
    ' possibility that traders could have data turned off, but trading turned on for symbols.  Because of
    ' this, we need to let them trade symbols that they don't get data for ( we will check for delayed data
    ' when they try to enable the automated item )...
    'ElseIf g.Broker.IsEnabledSymbol(vAccountNumberOrID, vSymbolOrSymbolID) = False Then
    '    bReturn = False
    ElseIf g.Broker.IsPitSymbol(vAccountNumberOrID, vSymbolOrSymbolID) = True Then
        bReturn = False
    ElseIf (g.Broker.PositionVerify(g.Broker.AccountTypeForID(lAccountID)) = True) And (bForceVerify = True) Then
        bReturn = False
    ElseIf (g.Broker.PositionMatch(vAccountNumberOrID, vSymbolOrSymbolID) = False) And (bForcePositionMatch = True) Then
        bReturn = False
    Else
        bReturn = True
    End If
    
    AllowAutoExits = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.AllowAutoExits"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PostFixVersion32
'' Description: Post fix the database tables for version 32 update
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PostFixVersion32()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim AcctPos As cAccountPosition     ' Account Position object
    Dim bNotified As Boolean            ' Has the user been notified?
    
    bNotified = False
    
    Set rs2 = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionTrades] " & _
                "WHERE [EntryFillID]<>0;", dbOpenDynaset)
    If (rs2.BOF And rs2.EOF) Then
        Set rs = g.dbPaper.OpenRecordset("SELECT DISTINCT AccountPositionID FROM [tblAccountPositionTrades] " & _
                    "WHERE [EntryFillID]=0;", dbOpenDynaset)
        Do While Not rs.EOF
            If bNotified = False Then
                frmSplash.Message -1, "Applying update to TradeTracker.MDB ..."
                bNotified = True
            End If
            
            Set AcctPos = New cAccountPosition
            If AcctPos.Load(rs!AccountPositionID) Then
                AcctPos.RecalculateHistory
            End If
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.PostFixVersion32"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActivateAutoExit
'' Description: Attempt to active an auto exit for the given information
'' Inputs:      Account, Symbol, Source
'' Returns:     Auto Exit Name, Blank if none or a problem
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ActivateAutoExit(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal strSource As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim lAccountID As Long              ' Account ID
    Dim strAccountNumber As String      ' Account number
    Dim strSymbol As String             ' Symbol
    
    strReturn = ""
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    strAccountNumber = g.Broker.AccountNumberForID(lAccountID)
    nBroker = g.Broker.AccountTypeForID(lAccountID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)

    If CanActivateAutomatedItem(vAccountNumberOrID, vSymbolOrSymbolID, "Auto Exit", strSource) Then
        strReturn = frmOrderStrategies.ShowMe(lAccountID, vSymbolOrSymbolID)
        If Len(strReturn) > 0 Then
            g.Broker.BrokerDebug nBroker, "Auto Exit Activate from " & strSource & " (" & strSymbol & ", " & strAccountNumber & "): Activating " & strReturn
            If g.OrderStrategies.ActivateExit(lAccountID, vSymbolOrSymbolID, strReturn) = False Then
                strReturn = ""
            End If
        End If
    End If
    
    ActivateAutoExit = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ActivateAutoExit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutomatedQuantityError
'' Description: Build error if the given quantity is not valid for automated trading
'' Inputs:      Account, Symbol, Automated Type, Source
'' Returns:     Error
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AutomatedQuantityError(ByVal TradeItem As cAutoTradeItem, ByVal strEditText As String, ByVal strSymbolError As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lNewValue As Long               ' New value for the quantity
    Dim lMax As Long                    ' Maximum quantity allowed
    Dim lMinQuantity As Long            ' Minimum quantity
    Dim lMinLotSize As Long             ' Minimum lot size
    Dim lMaxUnits As Long               ' Maximum number of units allowed for basket

    strReturn = ""
    lNewValue = CLng(Val(strEditText))
    lMinQuantity = g.Broker.MinimumOrderQuantity(TradeItem.AccountID, TradeItem.SymbolOrSymbolID)
    lMinLotSize = g.Broker.MinimumLotSize(TradeItem.AccountID, TradeItem.SymbolOrSymbolID)
    lMaxUnits = mSysNav.MaxUnitsForAutoTrade(TradeItem)
            
    If (lNewValue <> 0&) And (Len(strSymbolError) > 0) Then
        strReturn = strSymbolError
    ElseIf IsAlpha(strEditText) Then
        strReturn = "The quantity of the next entry must be a number"
    ElseIf lNewValue < 0 Then
        strReturn = "The quantity of the next entry cannot be a negative number"
    ElseIf lMaxUnits <> Abs(kNullData) Then
        lMax = TradeItem.StrategyBasketItemMult * lMaxUnits
        If lNewValue > lMax Then
            strReturn = "The quantity for this item cannot exceed " & Str(lMax)
        End If
    ElseIf (lNewValue < lMinQuantity) And (lNewValue <> 0&) Then
        strReturn = "The quantity of the next entry for " & TradeItem.Symbol & " cannot be less than " & Str(lMinQuantity)
    ElseIf lNewValue Mod lMinLotSize <> 0 Then
        strReturn = "The quantity of the next entry for " & TradeItem.Symbol & " must be an even multiple of " & Str(lMinLotSize)
    End If
    
    AutomatedQuantityError = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.AutomatedQuantityError"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutomatedSymbolError
'' Description: Build error if the given account/symbol are not valid for automated trading
'' Inputs:      Account, Symbol, Automated Type, Source
'' Returns:     Error
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AutomatedSymbolError(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal strAutomatedType As String, Optional ByVal bCheckAutoTradeEnable As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lAccountID As Long              ' Account ID
    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim strSymbol As String             ' Symbol
    Dim strBrokerBase As String         ' Broker base symbol
    Dim strBrokerExchange As String     ' Broker exchange
    Dim strBrokerName As String         ' Broker name
    Dim strType As String               ' Type of continuous contract
    
    strReturn = ""
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    nBroker = g.Broker.AccountTypeForID(lAccountID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    strBrokerName = g.Broker.BrokerName(nBroker)

    If (IsSpreadSymbol(strSymbol) = True) And ((HasModule("AUTOSPR") = False) And (FileExist(AddSlash(App.Path) & "AutoSpr.FLG") = False)) Then
        strReturn = "You cannot set up " & strAutomatedType & " for spread symbols"
    ElseIf (Left(strSymbol, 1) = "$") And (IsForex(strSymbol) = False) Then
        strReturn = "You cannot set up " & strAutomatedType & " for index symbols"
    ElseIf InStr(strSymbol, " ") > 0 Then
        strReturn = "You cannot set up " & strAutomatedType & " for option symbols"
    ElseIf g.Broker.IsTradeableSymbol(vAccountNumberOrID, vSymbolOrSymbolID) = False Then
        strReturn = g.Broker.UnknownSymbolError(strSymbol, nBroker, strBrokerName)
    ' 12/04/2014 DAJ: With the changes in exchange fees coming from the exchange, there is more of a
    ' possibility that traders could have data turned off, but trading turned on for symbols.  Because of
    ' this, we need to let them trade symbols that they don't get data for ( we will check for delayed data
    ' when they try to enable the automated item )...
    'ElseIf g.Broker.IsEnabledSymbol(vAccountNumberOrID, vSymbolOrSymbolID, strBrokerBase, strBrokerExchange) = False Then
    '    strReturn = g.Broker.NotEnabledForSymbolError(strSymbol, nBroker, strBrokerBase, strBrokerExchange, strBrokerName)
    ElseIf g.Broker.IsPitSymbol(vAccountNumberOrID, vSymbolOrSymbolID) = True Then
        strReturn = "You cannot set up " & strAutomatedType & " for pit session symbols"
    ElseIf bCheckAutoTradeEnable Then
        If IsForex(strSymbol) = True Then
            If (InStr(strSymbol, "@") <> 0) Then
                If ((nBroker <> eTT_AccountType_SimStream) And (nBroker <> eTT_AccountType_SimBroker)) Then
                    Select Case UCase(Parse(strSymbol, "@", 2))
                        Case "CNX"
                            If Not g.Broker.IsCurrenexBroker(nBroker) Then
                                strReturn = "You can only trade a Currenex Forex symbol in a Currenex or Trade Navigator account"
                            ElseIf (HasModule("AUTOCNXFX") = False) And (FileExist(AddSlash(App.Path) & "AutoCnxFx.FLG") = False) Then
                                strReturn = "You are not authorized to run an automated trading sytem on a Currenex Forex symbol"
                            End If
                        
                        Case "IB"
                            If Not g.Broker.IsIbBroker(nBroker) Then
                                strReturn = "You can only trade an Interactive Brokers|Forex symbol in an Interactive Brokers or|Trade Navigator account"
                            ElseIf (HasModule("AUTOIBFX") = False) And (FileExist(AddSlash(App.Path) & "AutoIbFx.FLG") = False) Then
                                strReturn = "You are not authorized to run an|automated trading sytem on an|Interactive Brokers Forex symbol"
                            End If
                        
'                        Case "PFG"
'                            If Not g.Broker.IsPfgBroker(nBroker) Then
'                                strReturn = "You can only trade a PFG Forex symbol in a PFG or Trade Navigator account"
'                            ElseIf Not FileExist(AddSlash(App.Path) & "AutoPfgFx.FLG") Then
'                                strReturn = "You are not authorized to run an automated trading sytem on a PFG Forex symbol"
'                            End If
                            
                        Case "OEC"
                            If Not g.Broker.IsOecBroker(nBroker) Then
                                strReturn = "You can only trade an Open E-Cry Forex symbol in an Open E-Cry or Trade Navigator account"
                            ElseIf (HasModule("AUTOOECFX") = False) And (FileExist(AddSlash(App.Path) & "AutoOecFx.FLG") = False) Then
                                strReturn = "You are not authorized to run an automated trading sytem on an Open E-Cry Forex symbol"
                            End If
                        
                        Case Else
                            strReturn = "You are not authorized to run an automated trading sytem on this symbol"
                    
                    End Select
                End If
            Else
                If g.Broker.IsLiveAccount(nBroker) Then
                    strReturn = "You cannot trade a Genesis Forex symbol in a broker account"
                End If
            End If
        
        ElseIf SecurityType(strSymbol, True) = "S" Then
            If (HasModule("AUTOSTK") = False) And (TypeOfAccount(vAccountNumberOrID) <> eGDTypeOfAccount_Simulated) Then
                strReturn = "Stock symbols are not allowed in|Automated Trading for a live account"
            End If
        
        ElseIf InStr(strSymbol, "-0") > 0 Then
            strType = Parse(strSymbol, "-", 2)
            If Not ((strType = "055") Or (strType = "065") Or (strType = "057") Or (strType = "067")) Then
                strReturn = "You cannot set up " & strAutomatedType & " for this kind of contract"
            End If
        End If
    End If
    
    AutomatedSymbolError = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.AutomatedSymbolError"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidAutomatedSymbol
'' Description: Determine if the given account/symbol are valid for automated trading
'' Inputs:      Account, Symbol, Automated Type, Source
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidAutomatedSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal strAutomatedType As String, ByVal strSource As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strError As String              ' Automated symbol error
    
    bReturn = True
    strError = AutomatedSymbolError(vAccountNumberOrID, vSymbolOrSymbolID, strAutomatedType)
    
    If (Len(strError) > 0) Then
        InfBox strError, "!", , strAutomatedType & " Error"
        bReturn = False
    End If
    
    ValidAutomatedSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ValidAutomatedSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanActivateAutomatedItem
'' Description: Determine if an automated item can be activated for the given account/symbol
'' Inputs:      Account, Symbol, Automated Type, Source, Show Message to User?
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CanActivateAutomatedItem(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal strAutomatedType As String, ByVal strSource As String, Optional ByVal bShowMessage As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lAccountID As Long              ' Account ID
    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim strAccountNumber As String      ' Account number
    Dim strSymbol As String             ' Symbol
    
    bReturn = False
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    strAccountNumber = g.Broker.AccountNumberForID(lAccountID)
    nBroker = g.Broker.AccountTypeForID(lAccountID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    
    If ValidAutomatedSymbol(vAccountNumberOrID, vSymbolOrSymbolID, strAutomatedType, strSource) Then
        If g.Broker.PositionVerify(nBroker) Then
            g.Broker.BrokerDebug nBroker, strAutomatedType & " Activate from " & strSource & " (" & strSymbol & ", " & strAccountNumber & "): Please wait until positions are verified for this symbol"
            If bShowMessage = True Then
                InfBox "Please wait until positions are verified for this symbol", "!", , strAutomatedType & " Error"
            End If
        ElseIf g.Broker.CarriedMatch(vAccountNumberOrID, vSymbolOrSymbolID) = False Then
            g.Broker.BrokerDebug nBroker, strAutomatedType & " Activate from " & strSource & " (" & strSymbol & ", " & strAccountNumber & "): You cannot set up " & strAutomatedType & " for this symbol because it currently has a carried position mismatch"
            If bShowMessage = True Then
                If InfBox("You cannot set up " & strAutomatedType & " for this symbol because the Trade Navigator carried position information does not match the broker carried position information.||Would you like to try to fix this now?|", "!", "+Fix|-Cancel", strAutomatedType & " Error") = "F" Then
                    g.Broker.FixPosition vAccountNumberOrID, vSymbolOrSymbolID
                End If
            End If
        ElseIf g.Broker.ConsistentBroker(vAccountNumberOrID, vSymbolOrSymbolID) = False Then
            g.Broker.BrokerDebug nBroker, strAutomatedType & " Activate from " & strSource & " (" & strSymbol & ", " & strAccountNumber & "): You cannot set up " & strAutomatedType & " for this symbol because the broker position data is inconsistent"
            If bShowMessage = True Then
                InfBox "You cannot set up " & strAutomatedType & " for this symbol because the broker position data is inconsistent", "!", , strAutomatedType & " Error"
            End If
        ElseIf HasModule("RTG,RTE") = False Then
            g.Broker.BrokerDebug nBroker, strAutomatedType & " Activate from " & strSource & " (" & strSymbol & ", " & strAccountNumber & "): You cannot set up " & strAutomatedType & " because you are not enabled for streaming"
            If bShowMessage = True Then
                InfBox "You cannot set up " & strAutomatedType & " because you are not enabled for streaming", "!", , strAutomatedType & " Error"
            End If
        Else
            bReturn = True
        End If
    End If

    CanActivateAutomatedItem = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.CanActivateAutomatedItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowTradeFilter
'' Description: Show the trade filter form
'' Inputs:      Account, Symbol, Auto Trade Item ID
'' Returns:     None
''
'' Settings:    chkDateRange, From Date, To Date, chkAccount, Account ID, chkSymbol,
''              Symbol, Direction, chkEntryRule, Entry Rule ID, chkExitRule
''              Exit Rule ID, RealSimFlag
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowTradeFilter(Optional ByVal vAccountNumberOrID As Variant = kNullData, Optional ByVal vSymbolOrSymbolID As Variant = kNullData, Optional ByVal lAutoTradeItemID As Long = kNullData)
On Error GoTo ErrSection:

    Dim Settings As cTradeFilterSettings ' Settings from the INI file
    
    If (vAccountNumberOrID <> kNullData) Or (vSymbolOrSymbolID <> kNullData) Or (lAutoTradeItemID <> kNullData) Then
        Set Settings = New cTradeFilterSettings
        Settings.LoadFromIni
        
        If vAccountNumberOrID <> kNullData Then
            Settings.UseAccount = True
            
            Settings.AccountIds.Clear
            Settings.AccountIds.Add g.Broker.GetAccountID(vAccountNumberOrID)
        End If
        
        If vSymbolOrSymbolID <> kNullData Then
            Settings.UseSymbol = True
            Settings.Symbol = GetSymbol(vSymbolOrSymbolID)
        End If
        
        If lAutoTradeItemID <> kNullData Then
            If lAutoTradeItemID = -1& Then
                Settings.UseAutoTrade = False
            Else
                Settings.UseAutoTrade = True
                Settings.AutoTradeID = lAutoTradeItemID
            End If
        End If
    End If
    
    frmTradeReportFilter.ShowMe Settings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ShowTradeFilter"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanMoveOrder
'' Description: Can we move the given order?
'' Inputs:      Order, Show Message?, Check Stop Order?, New Order Price
'' Returns:     True if can move, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CanMoveOrder(ByVal Order As cPtOrder, Optional ByVal bShowMessage As Boolean = True, Optional ByVal bCheckStop As Boolean = True, Optional ByVal dNewOrderPrice As Double = kNullData) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lNumTicks As Long               ' Number of ticks from the market
    
    bReturn = True
    
    ' Don't show the message for sure if this is an auto exit or automated trading order...
    If Order.AutoTradeItemID > 0 Then
        bShowMessage = False
    ' Go ahead and show the message for auto exit order when the user manually attempts to move them...
    'ElseIf g.OrderStrategies.OrderExistsInStrategy(Order) Then
        'bShowMessage = False
    End If

    If Order.Status <> eTT_OrderStatus_Parked Then
        If g.Broker.ConnectionStatusForAccount(Order.AccountID) <> eGDConnectionStatus_Connected Then
            If bShowMessage Then InfBox "You cannot move this order because you are not currently connected to the account", "!", "+-OK", "Modify Order Error", True
            g.Broker.BrokerDebug Order.Broker, "You cannot move this order because you are not currently connected to the account (" & Str(bShowMessage) & ")"
            bReturn = False
        ElseIf OrderIsPending(Order) = True Then
            If bShowMessage Then InfBox "This order cannot be modified because it is in a pending state.  Please wait for the order to be confirmed by the broker.", "!", "+-OK", "Modify Order Error", True
            g.Broker.BrokerDebug Order.Broker, "This order cannot be modified because it is in a pending state.  Please wait for the order to be confirmed by the broker. (" & Str(bShowMessage) & ")"
            bReturn = False
        ElseIf (bCheckStop = True) And ((Order.OrderType = eTT_OrderType_Stop) Or (Order.OrderType = eTT_OrderType_StopWithLimit) Or (Order.OrderType = eTT_OrderType_StopCloseOnly) Or (Order.OrderType = eTT_OrderType_StopWithLimitCloseOnly)) Then
            If (g.Broker.DontAllowStopMove = True) And (g.RealTime.Active = True) Then
                If dNewOrderPrice = kNullData Then
                    lNumTicks = NumTicksFromMarket(Order.OrderPrice(True), Order.SymbolOrSymbolID, , Order.TriggeredByPrice)
                Else
                    lNumTicks = NumTicksFromMarket(dNewOrderPrice, Order.SymbolOrSymbolID, , Order.TriggeredByPrice)
                End If
                If (lNumTicks <> kNullData) Then
                    If (Order.Buy = True) And (lNumTicks < 0&) Then
                        If bShowMessage Then InfBox "You cannot move a Buy Stop Order below the current market price", "!", "+-OK", "Modify Order Error", True
                        g.Broker.BrokerDebug Order.Broker, "You cannot move a Buy Stop Order below the current market price (Show Message = " & Str(bShowMessage) & ", Order Price = " & Str(Order.OrderPrice(True)) & ", Num Ticks = " & Str(lNumTicks) & ")"
                        bReturn = False
                    ElseIf (Order.Buy = False) And (lNumTicks > 0&) Then
                        If bShowMessage Then InfBox "You cannot move a Sell Stop Order above the current market price", "!", "+-OK", "Modify Order Error", True
                        g.Broker.BrokerDebug Order.Broker, "You cannot move a Sell Stop Order above the current market price (Show Message = " & Str(bShowMessage) & ", Order Price = " & Str(Order.OrderPrice(True)) & ", Num Ticks = " & Str(lNumTicks) & ")"
                        bReturn = False
                    ElseIf Abs(lNumTicks) <= g.Broker.NumTicksStopBuffer Then
                        If bShowMessage Then InfBox "You cannot move this order closer than " & Str(g.Broker.NumTicksStopBuffer) & " ticks from the current market price", "!", "+-OK", "Modify Order Error", True
                        g.Broker.BrokerDebug Order.Broker, "You cannot move this order closer than " & Str(g.Broker.NumTicksStopBuffer) & " ticks from the current market price (Show Message = " & Str(bShowMessage) & ", Order Price = " & Str(Order.OrderPrice(True)) & ", Num Ticks = " & Str(lNumTicks) & ")"
                        bReturn = False
                    End If
                End If
            End If
        End If
    End If
    
    CanMoveOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.CanMoveOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RollPosition
'' Description: Roll positions and orders from one contract to another contract
''              in the given account
'' Inputs:      Account, Old Contract, New Contract, Auto Trade Item ID, Confirm?
'' Returns:     True if rolled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RollPosition(ByVal vAccountNumberOrID As Variant, ByVal vOldSymbolOrSymbolID As Variant, ByVal vNewSymbolOrSymbolID As Variant, ByVal lAtID As Long, Optional ByVal bConfirm As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim Orders As cPtOrders             ' Orders for the old symbol
    Dim lPosition As Long               ' Current position in the old symbol
    Dim lTimeOut As Long                ' Timeout counter for waiting for the flatten
    Dim bContinue As Boolean            ' Do we want to continue?
    Dim dAdjustAmount As Double         ' Adjustment amount
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cPtOrder               ' Order object
    Dim strAutoExit As String           ' Auto exit for the old symbol
    Dim lAccountID As Long              ' Account ID for account passed in
    Dim nBroker As eTT_AccountType      ' Account type for the account passed in
    Dim OldBars As cGdBars              ' Bars for the old symbol
    Dim NewBars As cGdBars              ' Bars for the new symbol
    Dim OldAutoExit As cActiveExit      ' Active auto exit for the old symbol
    Dim NewAutoExit As cActiveExit      ' Active auto exit for the new symbol
    Dim strCurrentInfo As String        ' Current auto exit information

    bReturn = False
    
    ' Only allow rolling a symbol if streaming is active...
    If g.RealTime.Active Then
        lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
        nBroker = g.Broker.AccountTypeForID(lAccountID)
        
        ' Save information about orders and position for old symbol...
        Set Orders = g.Broker.WorkingOrders(vAccountNumberOrID, vOldSymbolOrSymbolID, lAtID)
        lPosition = g.Broker.CurrentPosition(vAccountNumberOrID, vOldSymbolOrSymbolID, lAtID)
        
        ' Confirm with the user if necessary...
        bContinue = True
        If (bConfirm = True) And (lPosition <> 0&) Then
            If InfBox("Are you sure that you want to roll your " & UCase(g.Broker.TextPosition(lPosition)) & " position from " & GetSymbol(vOldSymbolOrSymbolID) & " into " & GetSymbol(vNewSymbolOrSymbolID) & "?||(Only do this if the market is currently open)|", "?", "+Yes|-No", "Roll Symbol") = "N" Then
                bContinue = False
            End If
        End If
        
        g.Broker.BrokerDebug nBroker, "ROLLING: User rolling " & g.Broker.TextPosition(lPosition) & " position from " & GetSymbol(vOldSymbolOrSymbolID) & " to " & GetSymbol(vNewSymbolOrSymbolID)
        
        Set OldAutoExit = g.OrderStrategies.ExitObjectForAccountAndSymbol(lAccountID, vOldSymbolOrSymbolID)
        Set NewAutoExit = g.OrderStrategies.ExitObjectForAccountAndSymbol(lAccountID, vNewSymbolOrSymbolID)
        
        ' Get the current information from the auto exit for the old symbol...
        If Not OldAutoExit Is Nothing Then
            strCurrentInfo = OldAutoExit.GetCurrentInfo
            g.Broker.BrokerDebug nBroker, "ROLLING: Current Auto Exit Info = " & strCurrentInfo
        End If
    
        ' If confirmed, attempt to flatten the user's position in the old symbol...
        If bContinue Then
            ' Remove any auto exit orders for the old symbol from the collection...
            For lIndex = Orders.Count To 1 Step -1
                If g.OrderStrategies.OrderExistsInStrategy(Orders(lIndex)) Then
                    Orders.Remove lIndex
                End If
            Next lIndex
            
            If Orders.Count > 0 Then
                ' Make sure that the old symbol is streaming so that we can get a price for an adjustment...
                Set OldBars = New cGdBars
                DM_GetBars OldBars, vOldSymbolOrSymbolID, , Date
                g.RealTime.AddTickBuffer OldBars
                g.RealTime.SpliceBars OldBars
            
                ' Make sure that the new symbol is streaming so that we can get a price for an adjustment...
                Set NewBars = New cGdBars
                DM_GetBars NewBars, vNewSymbolOrSymbolID, , Date
                g.RealTime.AddTickBuffer NewBars
                g.RealTime.SpliceBars NewBars
            End If
            
            If lPosition <> 0& Then
                ' Flatten the old symbol...
                InfBox "Please wait while Trade Navigator flattens your " & UCase(g.Broker.TextPosition(lPosition)) & " position in " & GetSymbol(vOldSymbolOrSymbolID) & "...", , , "Roll Symbol", True
                g.Broker.BrokerDebug nBroker, "ROLLING: Flattening " & g.Broker.TextPosition(lPosition) & " position for " & GetSymbol(vOldSymbolOrSymbolID)
                
                FlattenForSymbol lAccountID, vOldSymbolOrSymbolID, lAtID
                    
                lTimeOut = 0&
                Do While (g.Broker.CurrentPosition(vAccountNumberOrID, vOldSymbolOrSymbolID, lAtID) <> 0) And (lTimeOut < 30&)
                    Sleep 1#
                    lTimeOut = lTimeOut + 1&
                Loop
                
                If g.Broker.CurrentPosition(vAccountNumberOrID, vOldSymbolOrSymbolID, lAtID) <> 0 Then bContinue = False
            ElseIf Orders.Count > 0 Then
                ' Cancel any working orders for the old contract...
                InfBox "Please wait while Trade Navigator cancels the working orders for " & GetSymbol(vOldSymbolOrSymbolID) & "...", , , "Roll Symbol", True
                g.Broker.BrokerDebug nBroker, "ROLLING: Cancelling working orders for " & GetSymbol(vOldSymbolOrSymbolID)
                
                g.Broker.CancelWorkingOrders vAccountNumberOrID, vOldSymbolOrSymbolID, lAtID
                
                lTimeOut = 0&
                Do While (g.Broker.HasWorkingOrders(vAccountNumberOrID, vOldSymbolOrSymbolID, lAtID) = True) And (lTimeOut < 30&)
                    Sleep 1#
                    lTimeOut = lTimeOut + 1&
                Loop
                
                bContinue = Not g.Broker.HasWorkingOrders(vAccountNumberOrID, vOldSymbolOrSymbolID, lAtID)
            End If
        End If
            
        ' If successfully flattened, attempt to open the position in the new contract...
        If bContinue Then
            If lPosition <> 0& Then
                ' Deactivate the auto exits for the new symbol...
                g.OrderStrategies.DeactivateExit lAccountID, vNewSymbolOrSymbolID, , "Rolling position"
                
                InfBox "Please wait while Trade Navigator enters you into a " & UCase(g.Broker.TextPosition(lPosition)) & " position in " & GetSymbol(vNewSymbolOrSymbolID) & "...", , , "Roll Symbol", True
                g.Broker.BrokerDebug nBroker, "ROLLING: Entering " & g.Broker.TextPosition(lPosition) & " position for " & GetSymbol(vNewSymbolOrSymbolID)
                EnterPositionForSymbol vAccountNumberOrID, vNewSymbolOrSymbolID, lAtID, lPosition
                
                lTimeOut = 0&
                Do While (g.Broker.CurrentPosition(vAccountNumberOrID, vNewSymbolOrSymbolID, lAtID) <> lPosition) And (lTimeOut < 30&)
                    Sleep 1#
                    lTimeOut = lTimeOut + 1&
                Loop
            
                If g.Broker.CurrentPosition(vAccountNumberOrID, vNewSymbolOrSymbolID, lAtID) <> lPosition Then bContinue = False
            End If
        End If
                
        ' If the user is in a position, recreate the orders that were working in the old symbol
        ' applying an adjustment based on the difference in price between the two symbols at the
        ' time of the flatten...
        If bContinue Then
            bReturn = True
            
            If (lPosition <> 0&) And (Orders.Count > 0) Then
                InfBox "Please wait while Trade Navigator submits your working orders for " & GetSymbol(vOldSymbolOrSymbolID) & "...", , , "Roll Symbol", True
                g.Broker.BrokerDebug nBroker, "ROLLING: Resubmitting working orders for new symbol"
                
                g.RealTime.UpdateBars OldBars
                g.RealTime.UpdateBars NewBars
                
                dAdjustAmount = NewBars(eBARS_Close, NewBars.Size - 1) - OldBars(eBARS_Close, OldBars.Size - 1)
                
                For lIndex = 1 To Orders.Count
                    Set Order = New cPtOrder
                    With Order
                        .AccountID = Orders(lIndex).AccountID
                        .AutoTradeItemID = Orders(lIndex).AutoTradeItemID
                        .Buy = Orders(lIndex).Buy
                        .Enter = Orders(lIndex).Enter
                        .Expiration = Orders(lIndex).Expiration
                        .GenesisOrderID = NextGenesisOrderID(g.Broker.GetAccountNumber(vAccountNumberOrID))
                        If Orders(lIndex).LimitPrice <> 0 Then .LimitPrice = Orders(lIndex).LimitPrice + dAdjustAmount
                        .OrderType = Orders(lIndex).OrderType
                        .Quantity = Orders(lIndex).RemainingQuantity
                        If Orders(lIndex).StopPrice <> 0 Then .StopPrice = Orders(lIndex).StopPrice + dAdjustAmount
                        .SymbolOrSymbolID = vNewSymbolOrSymbolID
                        .Save
                    End With
                    
                    SubmitOrder Order
                Next lIndex
            End If
            
            ' If auto-exit was set for old symbol, but no auto-exit is set for the new symbol, set
            ' the auto-exit for the new symbol to the auto-exit of the old symbol...
            If Len(g.OrderStrategies.ExitForAccountAndSymbol(lAccountID, vNewSymbolOrSymbolID)) = 0 Then
                strAutoExit = g.OrderStrategies.ExitForAccountAndSymbol(lAccountID, vOldSymbolOrSymbolID, True)
                If Len(strAutoExit) <> 0 Then
                    g.OrderStrategies.ActivateExit lAccountID, vNewSymbolOrSymbolID, strAutoExit, strCurrentInfo
                    g.OrderStrategies.DeactivateExit lAccountID, vOldSymbolOrSymbolID, , "Rolling Position"
                End If
            End If
        End If
        
        InfBox ""
    Else
        InfBox "You cannot roll positions if streaming is not active", "!", "", "Roll Symbol"
    End If
    
    RollPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.RollPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllowMIT
'' Description: Determine if we want to allow MIT orders or not
'' Inputs:      None
'' Returns:     True if Allow MIT, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllowMIT() As Boolean
On Error GoTo ErrSection:

    AllowMIT = False ' FileExist(AddSlash(App.Path) & "AllowMIT.FLG")
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.AllowMIT"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrependDateToPfgOrderIds
'' Description: Prepend the order date to the PFG order/fill Ids
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Sub PrependDateToPfgOrderIds()
'On Error GoTo ErrSection:
'
'    Dim rs As Recordset                 ' Recordset into the database
'    Dim strDate As String               ' Formatted version of the order date
'    Dim strSave As String               ' Previous version of the order ID
'    Dim lPos As Long                    ' Position of something in a sorted array
'    Dim astrChanged As cGdArray         ' Array of changed order IDs
'    Dim astrFillID As cGdArray          ' Fill ID split out into fields
'
'    Set astrChanged = New cGdArray
'
'    ' Search for orders for which there is no colon (:) in the BrokerOrderID...
'    Set rs = g.dbPaper.OpenRecordset("SELECT tblOrders.* " & _
'                                     "FROM tblOrders INNER JOIN tblAccounts ON tblOrders.AccountID=tblAccounts.AccountID " & _
'                                     "WHERE tblAccounts.AccountType=" & Str(eTT_AccountType_PFG) & " AND tblOrders.BrokerOrderID NOT LIKE '*:*';", dbOpenDynaset)
'    Do While Not rs.EOF
'        strDate = Format(rs!OrderDate, "YYYYMMDD")
'
'        If Len(rs!BrokerOrderID) > 0 Then
'            strSave = rs!BrokerOrderID
'
'            ' Prepend the order date to the PFG order ID (YYYYMMDD:Acct-OrderID)...
'            rs.Edit
'            If Left(rs!BrokerOrderID, 9) = strDate & "-" Then
'                rs!BrokerOrderID = strDate & ":" & Mid(rs!BrokerOrderID, 10)
'            Else
'                rs!BrokerOrderID = strDate & ":" & rs!BrokerOrderID
'            End If
'            rs.Update
'
'            ' Add the changed order ID to the sorted array...
'            If astrChanged.BinarySearch(strSave & vbTab, lPos) = False Then
'                astrChanged.Add strSave & vbTab & rs!BrokerOrderID, lPos
'            End If
'        End If
'
'        rs.MoveNext
'    Loop
'
'    ' Only go forward if we changed any order ID's...
'    If astrChanged.Size > 0 Then
'        ' Search for orders for which there is no colon (:) in the PreviousBrokerID...
'        Set rs = g.dbPaper.OpenRecordset("SELECT tblOrders.* " & _
'                                         "FROM tblOrders INNER JOIN tblAccounts ON tblOrders.AccountID=tblAccounts.AccountID " & _
'                                         "WHERE tblAccounts.AccountType=" & Str(eTT_AccountType_PFG) & " AND tblOrders.PreviousBrokerID NOT LIKE '*:*';", dbOpenDynaset)
'        Do While Not rs.EOF
'            If Len(rs!PreviousBrokerID) > 0 Then
'                ' Get the changed order ID from the array of changed order ID's...
'                If astrChanged.BinarySearch(rs!PreviousBrokerID & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
'                    rs.Edit
'                    rs!PreviousBrokerID = Parse(astrChanged(lPos), vbTab, 2)
'                    rs.Update
'                End If
'            End If
'
'            rs.MoveNext
'        Loop
'
'        ' Search for fills for which there is no colon (:) in the BrokerOrderID...
'        Set rs = g.dbPaper.OpenRecordset("SELECT tblFills.* " & _
'                                         "FROM tblFills INNER JOIN tblAccounts ON tblFills.AccountID=tblAccounts.AccountID " & _
'                                         "WHERE tblAccounts.AccountType=" & Str(eTT_AccountType_PFG) & " AND tblFills.BrokerOrderID NOT LIKE '*:*';", dbOpenDynaset)
'        Do While Not rs.EOF
'            If Len(rs!BrokerOrderID) > 0 Then
'                ' Get the changed order ID from the array of changed order ID's...
'                If astrChanged.BinarySearch(rs!BrokerOrderID & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
'                    Set astrFillID = New cGdArray
'
'                    astrFillID.SplitFields rs!BrokerFillID, "-"
'
'                    rs.Edit
'                    rs!BrokerOrderID = Parse(astrChanged(lPos), vbTab, 2)
'                    rs!BrokerFillID = rs!BrokerOrderID & "-" & astrFillID(astrFillID.Size - 1)
'                    rs.Update
'                End If
'            End If
'
'            rs.MoveNext
'        Loop
'    End If
'
'ErrExit:
'    Exit Sub
'
'ErrSection:
'    RaiseError "mTradeTracker.PrependDateToPfgOrderIds"
'
'End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextCustomTradeRuleID
'' Description: Determine the next available custom trade rule ID for the type
'' Inputs:      Rule Type
'' Returns:     Next available ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextCustomTradeRuleID(ByVal nRuleType As eGDTradeRuleTypes) As Long
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' File loaded up into an array
    Dim strFile As String               ' Filename to load
    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim lID As Long                     ' ID from the line in the file

    lReturn = kStartingCustomTradeRuleID
    If nRuleType = eGDTradeRuleType_Entry Then
        strFile = AddSlash(App.Path) & "Custom\ErFilter.TXT"
    Else
        strFile = AddSlash(App.Path) & "Custom\XrFilter.TXT"
    End If
    
    Set astrFile = New cGdArray
    If astrFile.FromFile(strFile) Then
        For lIndex = 0 To astrFile.Size - 1
            lID = CLng(Val(Parse(astrFile(lIndex), vbTab, 1)))
            If lID >= lReturn Then
                lReturn = lID + 1
            End If
        Next lIndex
    End If
    
    NextCustomTradeRuleID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.NextCustomTradeRuleID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransActSimulatedAccount
'' Description: Determine if the given account number is a TransAct sim account
'' Inputs:      Account Number
'' Returns:     True if TransAct simulated account, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TransActSimulatedAccount(ByVal strAccountNumber As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    If g.Transact Is Nothing Then
        Set g.Transact = New cTransact
    End If
    
    If strAccountNumber = kTransActSimUserAccount Then
        bReturn = True
    ElseIf InStr("," & g.Transact.DemoAccounts & ",", "," & strAccountNumber & ",") <> 0 Then
        bReturn = True
    Else
        bReturn = False
    End If
    
    TransActSimulatedAccount = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.TransActSimulatedAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransActLoginModeString
'' Description: Return a string representation of the login mode
'' Inputs:      Login Mode
'' Returns:     Login Mode String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TransActLoginModeString(ByVal nLoginMode As eGDTransActLoginModes) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nLoginMode
        Case eGDTransActLoginMode_Demo
            strReturn = "Demo"
        Case eGDTransActLoginMode_Live
            strReturn = "Live"
        Case eGDTransActLoginMode_SimLive
            strReturn = "Simulated Live"
        Case Else
            strReturn = Str(nLoginMode)
    End Select
    
    TransActLoginModeString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.TransActLoginModeString"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrependWeekNumToPfgOrderIds
'' Description: Prepend the week number of the order date to the PFG order/fill Ids
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Sub PrependWeekNumToPfgOrderIds()
'On Error GoTo ErrSection:
'
'    Dim rs As Recordset                 ' Recordset into the database
'    Dim strDate As String               ' Formatted version of the order date
'    Dim strWeekNum As String            ' Week number of the order date
'    Dim strSave As String               ' Previous version of the order ID
'    Dim lPos As Long                    ' Position of something in a sorted array
'    Dim astrChanged As cGdArray         ' Array of changed order IDs
'    Dim astrFillID As cGdArray          ' Fill ID split out into fields
'
'    Set astrChanged = New cGdArray
'
'    ' Search for orders for which there is no colon (:) in the BrokerOrderID...
'    Set rs = g.dbPaper.OpenRecordset("SELECT tblOrders.* " & _
'                                     "FROM tblOrders INNER JOIN tblAccounts ON tblOrders.AccountID=tblAccounts.AccountID " & _
'                                     "WHERE tblAccounts.AccountType=" & Str(eTT_AccountType_PFG) & " AND tblOrders.BrokerOrderID NOT LIKE '*|*';", dbOpenDynaset)
'    Do While Not rs.EOF
'        strWeekNum = Format(WkNum(rs!OrderDate), "00000")
'        strDate = Format(rs!OrderDate, "YYYYMMDD")
'
'        If Len(rs!BrokerOrderID) > 0 Then
'            strSave = rs!BrokerOrderID
'
'            ' Prepend the week number to the PFG order ID (00000|Acct-OrderID)...
'            rs.Edit
'            If Left(rs!BrokerOrderID, 9) = strDate & "-" Then
'                rs!BrokerOrderID = strWeekNum & "|" & Mid(rs!BrokerOrderID, 10)
'            ElseIf InStr(rs!BrokerOrderID, ":") <> 0 Then
'                rs!BrokerOrderID = strWeekNum & "|" & Parse(rs!BrokerOrderID, ":", 2)
'            Else
'                rs!BrokerOrderID = strWeekNum & "|" & rs!BrokerOrderID
'            End If
'            rs.Update
'
'            ' Add the changed order ID to the sorted array...
'            If astrChanged.BinarySearch(strSave & vbTab, lPos) = False Then
'                astrChanged.Add strSave & vbTab & rs!BrokerOrderID, lPos
'            End If
'        End If
'
'        rs.MoveNext
'    Loop
'
'    ' Only go forward if we changed any order ID's...
'    If astrChanged.Size > 0 Then
'        ' Search for orders for which there is no colon (:) in the PreviousBrokerID...
'        Set rs = g.dbPaper.OpenRecordset("SELECT tblOrders.* " & _
'                                         "FROM tblOrders INNER JOIN tblAccounts ON tblOrders.AccountID=tblAccounts.AccountID " & _
'                                         "WHERE tblAccounts.AccountType=" & Str(eTT_AccountType_PFG) & " AND tblOrders.PreviousBrokerID NOT LIKE '*|*';", dbOpenDynaset)
'        Do While Not rs.EOF
'            If Len(rs!PreviousBrokerID) > 0 Then
'                ' Get the changed order ID from the array of changed order ID's...
'                If astrChanged.BinarySearch(rs!PreviousBrokerID & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
'                    rs.Edit
'                    rs!PreviousBrokerID = Parse(astrChanged(lPos), vbTab, 2)
'                    rs.Update
'                End If
'            End If
'
'            rs.MoveNext
'        Loop
'
'        ' Search for fills for which there is no colon (:) in the BrokerOrderID...
'        Set rs = g.dbPaper.OpenRecordset("SELECT tblFills.* " & _
'                                         "FROM tblFills INNER JOIN tblAccounts ON tblFills.AccountID=tblAccounts.AccountID " & _
'                                         "WHERE tblAccounts.AccountType=" & Str(eTT_AccountType_PFG) & " AND tblFills.BrokerOrderID NOT LIKE '*|*';", dbOpenDynaset)
'        Do While Not rs.EOF
'            If Len(rs!BrokerOrderID) > 0 Then
'                ' Get the changed order ID from the array of changed order ID's...
'                If astrChanged.BinarySearch(rs!BrokerOrderID & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
'                    Set astrFillID = New cGdArray
'
'                    astrFillID.SplitFields rs!BrokerFillID, "-"
'
'                    rs.Edit
'                    rs!BrokerOrderID = Parse(astrChanged(lPos), vbTab, 2)
'                    rs!BrokerFillID = rs!BrokerOrderID & "-" & astrFillID(astrFillID.Size - 1)
'                    rs.Update
'                End If
'            End If
'
'            rs.MoveNext
'        Loop
'    End If
'
'ErrExit:
'    Exit Sub
'
'ErrSection:
'    RaiseError "mTradeTracker.PrependWeekNumToPfgOrderIds"
'
'End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsManExpressDemoAccount
'' Description: Determine if given account is a Man Express Demo account
'' Inputs:      Account Number
'' Returns:     True if Man Express Demo Account, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function IsManExpressDemoAccount(ByVal vAccountNumberOrID As Variant) As Boolean
'On Error GoTo ErrSection:
'
'    Dim strAccountNumber As String      ' Account number
'    Dim lAccountID As Long              ' Account ID
'    Dim bReturn As Boolean              ' Return value for the function
'    Dim nBroker As eTT_AccountType      ' Broker for the given account
'
'    strAccountNumber = g.Broker.GetAccountNumber(vAccountNumberOrID)
'    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
'    bReturn = False
'    nBroker = g.Broker.AccountTypeForID(lAccountID)
'
'    If (nBroker = eTT_AccountType_LindWaldock) Or (nBroker = eTT_AccountType_ManExpress) Then
'        If Len(strAccountNumber) >= 2 Then
'            If UCase(Left(strAccountNumber, 2)) = "PT" Then
'                bReturn = True
'            End If
'        End If
'    End If
'
'    IsManExpressDemoAccount = bReturn
'
'ErrExit:
'    Exit Function
'
'ErrSection:
'    RaiseError "mTradeTracker.IsManExpressDemoAccount"
'
'End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsPfgDemoAccount
'' Description: Determine if given account is a PFG Demo account
'' Inputs:      Account Number
'' Returns:     True if PFG Demo Account, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function IsPfgDemoAccount(ByVal vAccountNumberOrID As Variant) As Boolean
'On Error GoTo ErrSection:
'
'    Dim strAccountNumber As String      ' Account number
'    Dim lAccountID As Long              ' Account ID
'    Dim bReturn As Boolean              ' Return value for the function
'    Dim nBroker As eTT_AccountType      ' Broker for the given account
'
'    strAccountNumber = g.Broker.GetAccountNumber(vAccountNumberOrID)
'    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
'    bReturn = False
'    nBroker = g.Broker.AccountTypeForID(lAccountID)
'
'    If (nBroker = eTT_AccountType_CtgPfg) Or (nBroker = eTT_AccountType_FintecPfg) Or (nBroker = eTT_AccountType_PFG) Then
'        If Len(strAccountNumber) >= 1 Then
'            If UCase(Left(strAccountNumber, 1)) = "D" Then
'                bReturn = True
'            End If
'        End If
'    End If
'
'    IsPfgDemoAccount = bReturn
'
'ErrExit:
'    Exit Function
'
'ErrSection:
'    RaiseError "mTradeTracker.IsPfgDemoAccount"
'
'End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TypeOfAccount
'' Description: Determine the type of the given account
'' Inputs:      Account
'' Returns:     Account Type (Simulated, Broker Live, Broker Demo)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TypeOfAccount(ByVal vAccountNumberOrID As Variant) As eGDTypeOfAccount
On Error GoTo ErrSection:

    Dim nReturn As eGDTypeOfAccount     ' Return value for the function
    Dim Account As cPtAccount           ' Account object
    
    Set Account = g.Broker.Account(vAccountNumberOrID)
    If Not Account Is Nothing Then
        If g.Broker.IsLiveAccount(Account.AccountType) = False Then
            nReturn = eGDTypeOfAccount_Simulated
        
        ElseIf Account.AccountType = eTT_AccountType_DemoPats Then
            nReturn = eGDTypeOfAccount_BrokerDemo
        
        ElseIf g.Broker.IsIbBroker(Account.AccountType) = True Then
            If UCase(Left(Account.AccountNumber, 2)) = "DU" Then
                nReturn = eGDTypeOfAccount_BrokerDemo
            Else
                nReturn = eGDTypeOfAccount_BrokerLive
            End If
        
        ElseIf g.Broker.IsCqgBroker(Account.AccountType) = True Then
            If (UCase(Left(Account.FcmAccountNumber, 2)) = "PS") Or (UCase(Left(Account.FcmAccountNumber, 2)) = "TS") Then
                nReturn = eGDTypeOfAccount_BrokerDemo
            Else
                nReturn = eGDTypeOfAccount_BrokerLive
            End If
        
        Else
            nReturn = eGDTypeOfAccount_BrokerLive
        End If
    End If
    
    TypeOfAccount = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.TypeOfAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NotSent
'' Description: Determine if the order hasn't been sent yet by the status
'' Inputs:      Order Status
'' Returns:     True if Not Sent, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NotSent(ByVal nOrderStatus As eTT_OrderStatus) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    Select Case nOrderStatus
        Case eTT_OrderStatus_Open, eTT_OrderStatus_Parked, eTT_OrderStatus_TriggerPending, eTT_OrderStatus_DataPending
            bReturn = True
            
        Case Else
            bReturn = False
    End Select
    
    NotSent = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.NotSent"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasBeenSent
'' Description: Determine if the order has been sent by the status
'' Inputs:      Order Status
'' Returns:     True if Sent, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasBeenSent(ByVal nOrderStatus As eTT_OrderStatus) As Boolean
On Error GoTo ErrSection:

    HasBeenSent = Not NotSent(nOrderStatus)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.HasBeenSent"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanSubmitOCO
'' Description: Determine if we can submit both OCO orders
'' Inputs:      Order1, Order2
'' Returns:     True if can submit, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CanSubmitOCO(ByVal Order1 As cPtOrder, ByVal Order2 As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If Order1.WrongSideOfMarket Or Order2.WrongSideOfMarket Then
        bReturn = (InfBox("One or both sides of this OCO will be submitted on the wrong side of the market which could result in an immediate fill or an order rejection.||Do you want to continue?|", "!", "+Yes|No", "Submit OCO Warning") = "Y")
    End If
    
    CanSubmitOCO = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.CanSubmitOCO"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderLinkStatusString
'' Description: Convert the order link status to a string
'' Inputs:      Order Link Status
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function OrderLinkStatusString(ByVal nOrderLinkStatus As eGDOrderLinkStatus) As String
'On Error GoTo ErrSection:
'
'    Dim strReturn As String             ' Return value for the function
'
'    strReturn = ""
'    Select Case nOrderLinkStatus
'        Case eGDOrderLinkStatus_New
'            strReturn = "New"
'        Case eGDOrderLinkStatus_LinkSent
'            strReturn = "Link Sent"
'        Case eGDOrderLinkStatus_Confirmed
'            strReturn = "Confirmed"
'        Case eGDOrderLinkStatus_UnlinkSent
'            strReturn = "Unlink Sent"
'        Case eGDOrderLinkStatus_Parked
'            strReturn = "Parked"
'    End Select
'
'    OrderLinkStatusString = strReturn
'
'ErrExit:
'    Exit Function
'
'ErrSection:
'    RaiseError "mTradeTracker.OrderLinkStatusString"
'
'End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitAmend
'' Description: Submit an order amendment
'' Inputs:      Original Order, New Order, Allow merge?
'' Returns:     True if Submitted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SubmitAmend(OriginalOrder As cPtOrder, NewOrder As cPtOrder, Optional ByVal bAllowMerge As Boolean = True)
On Error GoTo ErrSection:

    Dim OtherOrder As cPtOrder          ' Other order in an order-cancel-order situation
    Dim strReturn As String             ' Return value from an InfBox
    
    'If OriginalOrder.BrokerCancelOrderID <> 0 Then
    '    g.OrderLinks.UnlinkAmendOneAndRelink OriginalOrder, NewOrder
    'Else
        ' If the order is not trigger pending or doesn't have a trigger, then try to submit...
        If (NewOrder.HasTrigger = False) Or (OriginalOrder.Status <> eTT_OrderStatus_TriggerPending) Then
        
            ' If important information about the order has changed, or the original was parked, then submit...
            If (NewOrder.OrderText <> OriginalOrder.OrderText) Or (OriginalOrder.Status = eTT_OrderStatus_Parked) Then
                ' DAJ 04/01/2013: Found an issue with a set of trade logs that changing the
                ' order status here to 'Amend Pending' for an order that was parked was causing
                ' the software to try to submit an Amend instead of an Add for CQG brokers.  So
                ' I figured out that I need to leave the order status parked ( at least for
                ' the CQG brokers )...
                
                ' Setup amend pending information depending on the broker for the account...
                Select Case g.Broker.AccountTypeForID(NewOrder.AccountID)
                    Case eTT_AccountType_AlpariCurrenex
                        If Not g.AlpariCnx Is Nothing Then
                            g.AlpariCnx.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                    
                    Case eTT_AccountType_AdvFut
                        NewOrder.ChangeOrderStatus eTT_OrderStatus_AmendPending
                        g.AdvFut.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                    
                    Case eTT_AccountType_AmpCqg
                        If OriginalOrder.Status = eTT_OrderStatus_Parked Then
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_Parked
                        Else
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_AmendPending
                            g.Broker.BrokerDebug OriginalOrder.Broker, vbTab & "Order added to amend orders array: '" & OriginalOrder.OrderText(True, True, True) & "'"
                            g.AmpCqg.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                    
                    Case eTT_AccountType_CQG
                        If OriginalOrder.Status = eTT_OrderStatus_Parked Then
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_Parked
                        Else
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_AmendPending
                            g.Broker.BrokerDebug OriginalOrder.Broker, vbTab & "Order added to amend orders array: '" & OriginalOrder.OrderText(True, True, True) & "'"
                            g.CQG.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                        
                    Case eTT_AccountType_CtgCqg
                        If OriginalOrder.Status = eTT_OrderStatus_Parked Then
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_Parked
                        Else
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_AmendPending
                            g.Broker.BrokerDebug OriginalOrder.Broker, vbTab & "Order added to amend orders array: '" & OriginalOrder.OrderText(True, True, True) & "'"
                            g.CtgCqg.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                        
                    Case eTT_AccountType_Currenex
                        If Not g.Currenex Is Nothing Then
                            g.Currenex.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                    
                    Case eTT_AccountType_FxddCurrenex
                        If Not g.FxddCnx Is Nothing Then
                            g.FxddCnx.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                    
                    Case eTT_AccountType_KnightCqg
                        If OriginalOrder.Status = eTT_OrderStatus_Parked Then
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_Parked
                        Else
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_AmendPending
                            g.Broker.BrokerDebug OriginalOrder.Broker, vbTab & "Order added to amend orders array: '" & OriginalOrder.OrderText(True, True, True) & "'"
                            g.KnightCqg.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                        
                    Case eTT_AccountType_KnightCurrenex
                        If Not g.KnightCnx Is Nothing Then
                            g.KnightCnx.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                    
'                        Case eTT_AccountType_LindWaldock
'                            NewOrder.ChangeOrderStatus eTT_OrderStatus_AmendPending
'                            g.LindWaldock.AddAmendPendingInfo NewOrder.BrokerID, OriginalOrder.OrderID, OriginalOrder.GenesisOrderID, NewOrder.OrderID, NewOrder.GenesisOrderID
'
'                        Case eTT_AccountType_ManExpress
'                            NewOrder.ChangeOrderStatus eTT_OrderStatus_AmendPending
'                            g.ManExpress.AddAmendPendingInfo NewOrder.BrokerID, OriginalOrder.OrderID, OriginalOrder.GenesisOrderID, NewOrder.OrderID, NewOrder.GenesisOrderID
                    
                    Case eTT_AccountType_RjoCqg
                        If OriginalOrder.Status = eTT_OrderStatus_Parked Then
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_Parked
                        Else
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_AmendPending
                            g.Broker.BrokerDebug OriginalOrder.Broker, vbTab & "Order added to amend orders array: '" & OriginalOrder.OrderText(True, True, True) & "'"
                            g.RjoCqg.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                    
                    Case eTT_AccountType_RobbinsCqg
                        If OriginalOrder.Status = eTT_OrderStatus_Parked Then
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_Parked
                        Else
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_AmendPending
                            g.Broker.BrokerDebug OriginalOrder.Broker, vbTab & "Order added to amend orders array: '" & OriginalOrder.OrderText(True, True, True) & "'"
                            g.RobbinsCqg.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                        
                    Case eTT_AccountType_TransAct
                        If Not g.Transact Is Nothing Then
                            g.Transact.ModifiedOrders.Add OriginalOrder, OriginalOrder.BrokerID
                        End If
                        
                    Case eTT_AccountType_TT
                        NewOrder.ChangeOrderStatus eTT_OrderStatus_AmendPending
                        g.TT.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        
                    Case eTT_AccountType_VanKarCurrenex
                        If Not g.VanKarCnx Is Nothing Then
                            g.VanKarCnx.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                    
                    Case eTT_AccountType_VisionCqg
                        If OriginalOrder.Status = eTT_OrderStatus_Parked Then
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_Parked
                        Else
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_AmendPending
                            g.Broker.BrokerDebug OriginalOrder.Broker, vbTab & "Order added to amend orders array: '" & OriginalOrder.OrderText(True, True, True) & "'"
                            g.VisionCqg.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                        
                    Case eTT_AccountType_ZanerCqg
                        If OriginalOrder.Status = eTT_OrderStatus_Parked Then
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_Parked
                        Else
                            NewOrder.ChangeOrderStatus eTT_OrderStatus_AmendPending
                            g.Broker.BrokerDebug OriginalOrder.Broker, vbTab & "Order added to amend orders array: '" & OriginalOrder.OrderText(True, True, True) & "'"
                            g.ZanerCqg.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                        
                    Case eTT_AccountType_ZanerCurrenex
                        If Not g.ZanerCnx Is Nothing Then
                            g.ZanerCnx.AmendOrders.Add OriginalOrder, OriginalOrder.GenesisOrderID
                        End If
                    
                End Select
                
                ' Submit the order...
                SubmitOrder NewOrder, OriginalOrder.GenesisOrderID, , , , bAllowMerge
                
            ' If none of the main order stuff changed, check the advanced information and deal with
            ' it as necessary...
            Else
                ' DAJ 08/14/2009: If this is an order that was parked as one side of an order-
                ' cancel-order situation, we need to see if the user wants to submit the other
                ' side of the OCO as well, break the OCO link, or cancel the submission altogether...
                If (NewOrder.AutoTradeItemID = 0&) And (NewOrder.IsAutoExit = False) Then
                    If (NewOrder.CancelOrderID <> 0) Then
                        Set OtherOrder = New cPtOrder
                        If OtherOrder.Load(NewOrder.CancelOrderID) Then
                            If OtherOrder.Status = eTT_OrderStatus_Parked Then
                                ' If we are in a submit all situation, then assume Submit here, otherwise
                                ' ask the user whether to submit or break the OCO...
                                strReturn = InfBox("You have chosen to submit one side of an Order-Cancel-Order.  What would you like to do with the other one?", "?", "+Submit|Break OCO|-Cancel", "Order Cancel Order")
                                Select Case strReturn
                                    Case "S"
                                        ' DAJ 08/27/2009: First need to check to make sure that both sides of the
                                        ' OCO will be on the correct side of the market...
                                        If CanSubmitOCO(NewOrder, OtherOrder) = True Then
                                            NewOrder.CancelOrderID = SubmitOrder(OtherOrder, OtherOrder.GenesisOrderID)
                                        End If
                                        
                                    Case "B"
                                        NewOrder.CancelOrderID = 0
                                    
                                    Case "C"
                                        Exit Sub
                                End Select
                            End If
                        End If
                    End If
                End If
                
                ' If we didn't end up submitting the order because none of the important information changed
                ' or it is just a parked order, we can get rid of the new copy based on the broker for the
                ' account and save the new information...
                Select Case g.Broker.AccountTypeForID(NewOrder.AccountID)
'                    Case eTT_AccountType_LindWaldock, eTT_AccountType_ManExpress
'                        NewOrder.Delete
'                        OriginalOrder.CopyAdvancedInfo NewOrder
'                        OriginalOrder.Save
                
                    Case Else
                        NewOrder.Save
                        g.Broker.AddOrder NewOrder
                        OrderCallback NewOrder
                
                End Select
            End If
            
        ' If this is a trigger pending order, no need submitting it, just set it to trigger pending...
        Else
            SetTriggerOrderStatus NewOrder
        End If
        
        AdjustTriggers OriginalOrder, NewOrder
    'End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SubmitAmend"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetupOrderToEdit
'' Description: Setup an order to be able to edit it
'' Inputs:      Original Order, New Order
'' Returns:     Order to Edit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SetupOrderToEdit(ByVal OriginalOrder As cPtOrder, bNewOrder As Boolean) As cPtOrder
On Error GoTo ErrSection:

    Dim OrderToEdit As cPtOrder         ' Order object
    
    bNewOrder = False
    
    Set OrderToEdit = OriginalOrder.MakeCopy
    If HasBeenSent(OriginalOrder.Status) Then
        If XpressAccount(OriginalOrder.Broker) Then
            bNewOrder = True
            
            OrderToEdit.OrderID = 0
            OrderToEdit.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(OrderToEdit.AccountID))
            OrderToEdit.PreviousBrokerID = ""
            OrderToEdit.OrderDate = 0#
            OrderToEdit.Quantity = OrderToEdit.Quantity - OrderToEdit.FillQuantity
            OrderToEdit.Fills.Clear
            OrderToEdit.History.Clear
            OrderToEdit.Save
        End If
    End If
    
    Set SetupOrderToEdit = OrderToEdit

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.SetupOrderToEdit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    XpressAccount
'' Description: Is the given account type an Xpress account?
'' Inputs:      Account Type
'' Returns:     True if Xpress account, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function XpressAccount(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    Select Case nBroker
'        Case eTT_AccountType_LindWaldock, eTT_AccountType_ManExpress
'            bReturn = True
        Case Else
            bReturn = False
    End Select
    
    XpressAccount = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.XpressAccount"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AdjustTriggers
'' Description: Adjust triggered by orders if necessary
'' Inputs:      Original Order, New Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AdjustTriggers(ByVal OriginalOrder As cPtOrder, ByVal NewOrder As cPtOrder)
On Error GoTo ErrSection:

    If (NewOrder.OrderType = OriginalOrder.OrderType) Then
        Select Case OriginalOrder.OrderType
            Case eTT_OrderType_Stop, eTT_OrderType_StopCloseOnly, eTT_OrderType_StopWithLimit, eTT_OrderType_StopWithLimitCloseOnly
                If NewOrder.StopPrice <> OriginalOrder.StopPrice Then
                    AdjustTriggeredOrders NewOrder, NewOrder.StopPrice - OriginalOrder.StopPrice
                End If
            
            Case eTT_OrderType_Limit, eTT_OrderType_LimitCloseOnly
                If NewOrder.LimitPrice <> OriginalOrder.LimitPrice Then
                    AdjustTriggeredOrders NewOrder, NewOrder.LimitPrice - OriginalOrder.LimitPrice
                End If
            
            Case eTT_OrderType_MIT
                If NewOrder.MitPrice <> OriginalOrder.MitPrice Then
                    AdjustTriggeredOrders NewOrder, NewOrder.MitPrice - OriginalOrder.MitPrice
                End If
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.AdjustTriggers"

End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    WaitListCommandString
'' Description: Convert a wait list command to a descriptive string
'' Inputs:      Wait List Command
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WaitListCommandString(ByVal nWaitListCommand As eGDWaitListCommands) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case nWaitListCommand
        Case eGDWaitListCommand_Link
            strReturn = "Link"
        Case eGDWaitListCommand_Cancel
            strReturn = "Cancel"
        Case eGDWaitListCommand_ParkOne
            strReturn = "ParkOne"
        Case eGDWaitListCommand_ParkBoth
            strReturn = "ParkBoth"
        Case eGDWaitListCommand_AmendOne
            strReturn = "AmendOne"
        Case eGDWaitListCommand_AmendBoth
            strReturn = "AmendBoth"
        Case eGDWaitListCommand_SubmitOne
            strReturn = "SubmitOne"
        Case eGDWaitListCommand_SubmitBoth
            strReturn = "SubmitBoth"
        Case Else
            strReturn = Str(nWaitListCommand)
    End Select
    
    WaitListCommandString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.WaitListCommandString"
    
End Function
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanModify
'' Description: Determine if we can modify the order
'' Inputs:      Order, New Price, Show Message?, Pending?
'' Returns:     True if we can modify the order, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CanModify(ByVal Order As cPtOrder, Optional ByVal dNewPrice As Double = 0#, Optional ByVal bShowMessage As Boolean = True, Optional bPending As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lNumTicks As Long               ' Number of ticks from the market
    
    bReturn = True
    bPending = False
    
    If IsOpenOrder(Order.Status, False) = False Then
        g.Broker.BrokerDebug Order.Broker, "Cannot move order because it is pending"
        bReturn = False
        bPending = True
    ElseIf Order.BrokerCancelOrderID < 0 Then
        g.Broker.BrokerDebug Order.Broker, "Cannot move order because broker OCO pending"
        bReturn = False
    ElseIf (Order.OrderType = eTT_OrderType_Stop) Or (Order.OrderType = eTT_OrderType_StopWithLimit) Then
        If (dNewPrice <> 0#) And (g.Broker.DontAllowStopMove = True) Then
            lNumTicks = NumTicksFromMarket(dNewPrice, Order.SymbolOrSymbolID)
            If Order.Buy Then
                If lNumTicks <= g.Broker.NumTicksStopBuffer Then
                    If bShowMessage Then InfBox "You cannot move this order closer than " & Str(g.Broker.NumTicksStopBuffer) & " ticks from the current market price", "!", "+-OK", "Modify Order Error", True
                    g.Broker.BrokerDebug Order.Broker, "Cannot move order because it would move closer than " & g.Broker.NumTicksStopBuffer & " ticks from the market (Show Message = " & Str(bShowMessage) & ", Order Price = " & Str(Order.OrderPrice(True)) & ", Num Ticks = " & Str(lNumTicks) & ")"""
                    bReturn = False
                End If
            Else
                If lNumTicks >= (g.Broker.NumTicksStopBuffer * -1) Then
                    If bShowMessage Then InfBox "You cannot move this order closer than " & Str(g.Broker.NumTicksStopBuffer) & " ticks from the current market price", "!", "+-OK", "Modify Order Error", True
                    g.Broker.BrokerDebug Order.Broker, "Cannot move order because it would move closer than " & g.Broker.NumTicksStopBuffer & " ticks from the market (Show Message = " & Str(bShowMessage) & ", Order Price = " & Str(Order.OrderPrice(True)) & ", Num Ticks = " & Str(lNumTicks) & ")"""
                    bReturn = False
                End If
            End If
        End If
    End If

    CanModify = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.CanModify"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveOldTriggerPending
'' Description: Remove old trigger pending orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveOldTriggerPending()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Order As cPtOrder               ' Order object
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                "WHERE [Status]=" & Str(eTT_OrderStatus_TriggerPending) & ";")
    Do While Not rs.EOF
        ' If it is a simulated account, then all of the orders should have been
        ' cancelled the last time that Trade Navigator was closed, so make
        ' sure to mark Trigger Pending orders as cancelled as well...
        If g.Broker.AccountTypeForID(rs!AccountID) = eTT_AccountType_SimStream Then
            rs.Edit
            rs!Status = eTT_OrderStatus_Cancelled
            rs.Update
            
        ' If it was not a simulated account, check to see if the Triggering order
        ' is closed -- if so, we don't want the order Trigger Pending anymore...
        ElseIf rs!TriggerOrderID <> 0 Then
            Set Order = New cPtOrder
            If Order.Load(rs!TriggerOrderID) Then
                If IsOpenOrder(Order.Status) = False Then
                    rs.Edit
                    rs!Status = eTT_OrderStatus_Cancelled
                    rs.Update
                End If
            End If
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.RemoveOldTriggerPending"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveOldSentOrders
'' Description: Remove old "sent" orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveOldSentOrders()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                "WHERE [Status]=" & Str(eTT_OrderStatus_Sent) & ";")
    Do While Not rs.EOF
        If g.Broker.AccountTypeForID(rs!AccountID) = eTT_AccountType_SimStream Then
            rs.Edit
            rs!Status = eTT_OrderStatus_Cancelled
            rs.Update
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.RemoveOldSentOrders"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveOldIbGenesisIds
'' Description: Clear the GenesisOrderID for Interactive Brokers history orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveOldIbGenesisIds()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * " & _
                "FROM [tblOrders] INNER JOIN [tblAccounts] ON tblOrders.AccountID=tblAccounts.AccountID " & _
                "WHERE [IsSnapshot]=0 AND LEN([GenesisOrderID])>0", dbOpenDynaset)
    Do While Not rs.EOF
        If g.Broker.IsIbBroker(rs!AccountType) = True Then
            rs.Edit
            rs!GenesisOrderID = ""
            rs.Update
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.RemoveOldIbGenesisIds"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixGroupIds
'' Description: Fix any duplicate Group Id's from Option Navigator on orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FixGroupIds()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim bNeedToFix As Boolean           ' Do we need to run the fix?
    Dim lMaxID As Long                  ' Max Group ID
    Dim astrNames As cGdArray           ' Array of group ID/group name pairs
    Dim lPos As Long                    ' Position in the array
    
    bNeedToFix = False
    lMaxID = 0&
    Set astrNames = New cGdArray
    astrNames.Create eGDARRAY_Strings
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [GroupID]>0;", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        Do While Not rs.EOF
            If Len(rs!GroupName) > 0 Then
                If astrNames.BinarySearch(Str(rs!GroupID) & "=", lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                    If Parse(astrNames(lPos), "=", 2) <> rs!GroupName Then
                        bNeedToFix = True
                    End If
                Else
                    astrNames.Add Str(rs!GroupID) & "=" & rs!GroupName, lPos
                End If
            Else
                bNeedToFix = True
            End If
            
            If rs!GroupID > lMaxID Then
                lMaxID = rs!GroupID
            End If
            
            rs.MoveNext
        Loop
        
        If bNeedToFix = True Then
            rs.MoveFirst
            Do While Not rs.EOF
                If Len(rs!GroupName) = 0 Then
                    rs.Edit
                    rs!GroupID = 0&
                    rs.Update
                ElseIf astrNames.BinarySearch(Str(rs!GroupID) & "=", lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                    If Parse(astrNames(lPos), "=", 2) <> rs!GroupName Then
                        lMaxID = lMaxID + 1&
                        rs.Edit
                        rs!GroupID = lMaxID
                        rs.Update
                    End If
                End If
                
                rs.MoveNext
            Loop
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.FixGroupIds"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PerformStartupFixes
'' Description: Perform fixes to trade information at program startup
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PerformStartupFixes()
On Error GoTo ErrSection:

    ' Make sure to do this before loading the online broker form which will load the SimTrade object
    ' which will load the SimTrade Orders and Fills, however it needs g.Broker to be alive...
    PostFixVersion28
    PostFixVersion32
    'PrependDateToPfgOrderIds
    'PrependWeekNumToPfgOrderIds
    RemoveOldTriggerPending
    RemoveOldSentOrders
    
    ' 05/25/2010 DAJ: Something is causing accounts to be added to the database with
    ' a blank account number and blank account name which is causing some issues in
    ' other places.  Remove them here...
    RemoveBlankAccounts
    
    ' 11/07/2011 DAJ: The version 57 update to the database caused a bunch of fills
    ' and orders to be fixed for Interactive Brokers / I-Deal accounts, so we need
    ' to rebuild the account position objects...
    PostFixVersion57
    
    ' 03/14/2012 DAJ: The version 61 update to the database needs to determine a symbol
    ' ID for a symbol, but the pool is not open yet, so do it here...
    PostFixVersion61
    
    ' 10/09/2014 DAJ: Customer ran into an issue where his GenesisOrderID's for Interactive
    ' Brokers ( the 'NextId' from the TWS ) reset.  This caused us to start finding old orders
    ' and think he was amending them.  To fix this, I think we can clear the GenesisOrderID
    ' when it moves to history for Interactive Brokers orders...
    RemoveOldIbGenesisIds

    ' 12/08/2015 DAJ: Option Navigator had a bug where every time a user would upgrade to a
    ' new version of Option Navigator, it started the Group ID counter over.  This caused the user
    ' to end up with orders grouped together with the same ID that shouldn't have been.  I put
    ' a fix in Option Navigator, but now we need to fix the database as well...
    FixGroupIds
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.PerformStartupFixes"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowOrderHistoryForID
'' Description: Show order history for the order with the given ID
'' Inputs:      Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowOrderHistoryForID(ByVal lOrderID As Long)
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order object
    
    If Order.Load(lOrderID) Then
        frmOrderHistory.ShowMe Order
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mTradeTracker.ShowOrderHistoryForID"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectXOS
'' Description: Allow the user to activate an exit order strategy
'' Inputs:      Account, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SelectXOS(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant, Optional ByVal strNewStrategy As String)
On Error GoTo ErrSection:

    Dim strStrategy As String           ' Strategy selected by the user

    If Not g.OrderStrategies Is Nothing Then
        If Len(strNewStrategy) > 0 Then
            strStrategy = strNewStrategy
        Else
            strStrategy = g.OrderStrategies.ExitForAccountAndSymbol(lAccountID, vSymbolOrSymbolID)
            strStrategy = frmOrderStrategies.ShowMe(lAccountID, vSymbolOrSymbolID, strStrategy, True)
        End If
        
        If Len(strStrategy) > 0 Then
            g.OrderStrategies.ActivateExit lAccountID, vSymbolOrSymbolID, strStrategy
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SelectXOS"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveXOS
'' Description: Allow the user to deactivate an exit order strategy
'' Inputs:      Account, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveXOS(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant, ByVal strSource As String)
On Error GoTo ErrSection:

    If Not g.OrderStrategies Is Nothing Then
        g.OrderStrategies.DeactivateExit lAccountID, vSymbolOrSymbolID, True, "Turned off from " & strSource
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.RemoveXOS"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowAccountHistory
'' Description: Show the trade tracker form for the given account
'' Inputs:      Account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowAccountHistory(ByVal lAccountID As Long, Optional ByVal nBroker As eTT_AccountType = -1&)
On Error GoTo ErrSection:

    If nBroker = -1& Then
        nBroker = g.Broker.AccountTypeForID(lAccountID)
    End If
    frmTTPositions.ShowMe lAccountID, nBroker

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ShowAccountHistory"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTradeBars
'' Description: Retrieve a set of bars for trading purposes from Trade Console
'' Inputs:      Symbol, Add to RT?
'' Returns:     Bars (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetTradeBars(ByVal vSymbolOrSymbolID As Variant, Optional ByVal bAddToRT As Boolean = True) As cGdBars
On Error GoTo ErrSection:

    Dim Bars As cGdBars                 ' Bars to return from the function
    
    If FormIsLoaded("frmTTSummary") Then
        Set Bars = frmTTSummary.GetBars(vSymbolOrSymbolID, bAddToRT)
    Else
        Set Bars = Nothing
    End If
    
    Set GetTradeBars = Bars

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.GetTradeBars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeBarsExist
'' Description: Does the Trade Console have Bars for the given Symbol?
'' Inputs:      Symbol
'' Returns:     True if exist, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TradeBarsExist(ByVal vSymbolOrSymbolID As Variant) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If FormIsLoaded("frmTTSummary") Then
        bReturn = frmTTSummary.BarsExist(vSymbolOrSymbolID)
    End If
    
    TradeBarsExist = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.TradeBarsExist"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearUpdatedColors
'' Description: Clear the updated colors on both grids if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearUpdatedColors()
On Error GoTo ErrSection:
    
    If FormIsLoaded("frmTTSumary") Then
        frmTTSummary.ClearUpdatedColors
    End If
    If FormIsLoaded("frmWorkingOrders") Then
        frmWorkingOrders.ClearUpdatedColors
    End If
    If FormIsLoaded("frmOpenPositions") Then
        frmOpenPositions.ClearUpdatedColors
    End If
    If FormIsLoaded("frmTTPositions") Then
        frmTTPositions.ClearUpdatedColors
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mTradeTracker.ClearUpdatedColors"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoBrokerTimer
'' Description: Tell the trade console forms to do a broker timer run
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoBrokerTimer()
On Error GoTo ErrSection:

    If FormIsLoaded("frmTTSummary") Then
        frmTTSummary.DoBrokerTimer
    End If
    If FormIsLoaded("frmWorkingOrders") Then
        frmWorkingOrders.DoBrokerTimer
    End If
    If FormIsLoaded("frmOpenPositions") Then
        frmOpenPositions.DoBrokerTimer
    End If
    If FormIsLoaded("frmAccounts") Then
        frmAccounts.DoBrokerTimer
    End If
    If FormIsLoaded("frmTodaysFills") Then
        frmTodaysFills.DoBrokerTimer
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.DoBrokerTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetOtosNegative
'' Description: Set the TriggerOrderID negative on any order that is triggered
''              by the given order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetOtosNegative(ByVal Order As cPtOrder, Optional ByVal bParkTriggers As Boolean = False)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Order2 As cPtOrder              ' Order
    
    If Order.OrderID > 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [TriggerOrderID]=" & Str(Order.OrderID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            Set Order2 = New cPtOrder
            If Order2.Load(rs!OrderID, rs) Then
                If IsOpenOrder(Order2.Status) Then
                    g.Broker.BrokerDebug Order2.Broker, "Order: '" & Order2.OrderText & "' OTO being changed to " & Str(Order2.TriggerOrderID * -1&)
                    Order2.TriggerOrderID = Order2.TriggerOrderID * -1&
                    Order2.Save
                    
                    OrderCallback Order2
                    g.Broker.AddOrder Order2
                    
                    If bParkTriggers Then
                        g.Broker.BrokerDebug Order2.Broker, "Order: '" & Order2.OrderText & "' being parked because triggered by order being parked"
                        ParkOrder Order2, , False
                    End If
                End If
            End If
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SetOtosNegative"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeNegativeOtos
'' Description: Change orders with a negative old order ID as their OTO to the
''              new order ID
'' Inputs:      Order, Old Order ID, Force Submit Orders?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeNegativeOtos(ByVal Order As cPtOrder, ByVal lOldOrderID As Long, Optional ByVal bForceSubmitOtos As Boolean = False)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Order2 As cPtOrder              ' Order
    Dim Orders As New cGdTree           ' Collection of orders
    Dim lIndex As Long                  ' Index into a for loop
    Dim strIgnoreList As String         ' List of orders not to submit
    Dim strSubmitOto As String          ' Submit the OTO orders?
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [TriggerOrderID]=" & Str(lOldOrderID * -1) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        Set Order2 = New cPtOrder
        If Order2.Load(rs!OrderID, rs) Then
            g.Broker.BrokerDebug Order2.Broker, "Order: '" & Order2.OrderText & "' OTO being changed to " & Str(Order.OrderID)
            Order2.TriggerOrderID = Order.OrderID
            Order2.Save
            
            OrderCallback Order2
            g.Broker.AddOrder Order2
            
            If InStr("," & strIgnoreList & ",", "," & Str(Order2.OrderID) & ",") = 0 Then
                Orders.Add Order2, Str(Order2.OrderID)
                If Order2.CancelOrderID <> 0 Then
                    strIgnoreList = strIgnoreList & "," & Str(Order2.CancelOrderID) & ","
                End If
            End If
        End If
        
        rs.MoveNext
    Loop
    
    If Orders.Count > 0 Then
        strSubmitOto = "Y"
        If bForceSubmitOtos = False Then
            strSubmitOto = InfBox("'" & Order.OrderText & "' has parked order that are triggered off of it.  Do you want to submit them as well?", "?", "+Yes|-No", "Submit Order")
            g.Broker.BrokerDebug Order.Broker, "User answered '" & strSubmitOto & "' to submit parked triggered by orders for order " & Str(Order.OrderID)
        End If
        
        If strSubmitOto = "Y" Then
            For lIndex = 1 To Orders.Count
                SubmitOrder Orders(lIndex), , , , False
            Next lIndex
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.ChangeNegativeOtos"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SameTif
'' Description: Determine if the two orders have the same time-in-force
'' Inputs:      Order1, Order2
'' Returns:     True if Same TIF, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SameTif(ByVal Order1 As cPtOrder, ByVal Order2 As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If (Order1.Expiration < 0) And (Order2.Expiration < 0) Then
        bReturn = True
    ElseIf Order1.Expiration = Order2.Expiration Then
        bReturn = True
    End If
    
    SameTif = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.SameTif"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveBlankAccounts
'' Description: Remove any blank accounts
'' Inputs:      None
'' Returns:     True if removed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemoveBlankAccounts() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    
    bReturn = False
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                "WHERE [AccountNumber]='' OR [Name]='';", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Delete
        bReturn = True
        
        rs.MoveNext
    Loop
    
    RemoveBlankAccounts = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.RemoveBlankAccounts"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowAdvancedTSOG
'' Description: Determine whether to show advanced controls for TradeSense
''              orders/order groups
'' Inputs:      None
'' Returns:     True if show, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowAdvancedTSOG() As Boolean
On Error GoTo ErrSection:

    ShowAdvancedTSOG = FileExist(AddSlash(App.Path) & "TsoInputs.FLG")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ShowAdvancedTSOG"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartDanielCodeProcess
'' Description: Start the DanielCode process
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub StartDanielCodeProcess()
On Error GoTo ErrSection:

    Dim strStandAlone As String         ' Stand alone application information
    Dim strDcFile As String             ' Daniel code file
    Dim strDcArgsFile As String         ' Filename for the arguments to the Daniel Code process
    Dim astrArgs As cGdArray            ' Array of arguments to pass

    strDcFile = GetIniFileProperty("OutputFile", "", "DanielCode", AddSlash(App.Path) & "Provided\Provided.INI")

    If frmMain.tbToolbar.Tools("ID_DanCodeWeb").Enabled = True Then
        strStandAlone = DanielCodeProcess
        If (Len(strStandAlone) > 0) Then
            If FileExist(strStandAlone) Then
                If HasDotNet Then
                    frmMain.tbToolbar.Tools("ID_DanCodeWeb").Enabled = False
                    frmMain.tbToolbar.Tools("ID_GmajPro").Enabled = False
                    
                    ' DAJ 08/10/2010: If the output file from the Daniel Code application is
                    ' still hanging around for some reason when we get in here, kill it
                    ' so that the timer doesn't handle it...
                    If Len(strDcFile) > 0 Then
                        If FileExist(AddSlash(App.Path) & strDcFile) Then
                            KillFile AddSlash(App.Path) & strDcFile
                        End If
                    End If
                    
                    ' DAJ 09/30/2010: We are now going to send a file name over to the Daniel Code process
                    ' as a command line argument that can eventually have arguments in it...
                    strDcArgsFile = GetIniFileProperty("ArgsFile", "", "DanielCode", AddSlash(App.Path) & "Provided\Provided.INI")
                    If Len(strDcArgsFile) > 0 Then
                        strDcArgsFile = AddSlash(App.Path) & strDcArgsFile
                        
                        Set astrArgs = New cGdArray
                        astrArgs.Create eGDARRAY_Strings
                        If HasModule("DCPLUS") Then
                            astrArgs.Add "DCPLUS"
                        ElseIf HasModule("DCFOREX") Then
                            astrArgs.Add "DCFOREX"
                        ElseIf HasModule("DCFUTURE") Then
                            astrArgs.Add "DCFUTURE"
                        Else
                            astrArgs.Add ""
                        End If
                        
                        astrArgs.ToFile strDcArgsFile
                        
                        If InStr(strDcArgsFile, " ") > 0 Then
                            strDcArgsFile = Chr(34) & strDcArgsFile & Chr(34)
                        End If
                    End If
                
                    ' Run the stand-alone application...
                    DebugLog "Starting '" & strStandAlone & "' '" & strDcArgsFile & "'"
                    RunProcess strStandAlone, strDcArgsFile
                
                    ' Start the timer waiting for file to exist...
                    frmOnlineBroker.tmrDanielCode.Enabled = True
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.StartDanielCodeProcess"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DanielCodeProcess
'' Description: Return the Daniel Code process path and name
'' Inputs:      None
'' Returns:     Process Path/Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DanielCodeProcess() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strName As String               ' Name of the product
    
    strReturn = ""
    strName = frmMain.tbToolbar.Tools("ID_DanCodeWeb").Name
    
    If HasLevel(eTN4_Gold, , strName) Then
        strReturn = GetIniFileProperty("ExeName", "", "DanielCode", AddSlash(App.Path) & "Provided\Provided.INI")
        If Len(strReturn) > 0 Then
            strReturn = AddSlash(App.Path) & strReturn
        End If
    End If
    
    DanielCodeProcess = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.DanielCodeProcess"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartGmajProcess
'' Description: Start the Gmaj process
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub StartGmajProcess()
On Error GoTo ErrSection:

    Dim strStandAlone As String         ' Stand alone application information
    Dim strDcFile As String             ' Daniel code file
    Dim strDcArgsFile As String         ' Filename for the arguments to the Daniel Code process
    Dim astrArgs As cGdArray            ' Array of arguments to pass

    strDcFile = GetIniFileProperty("OutputFile", "", "GmajPro", AddSlash(App.Path) & "Provided\Provided.INI")

    If frmMain.tbToolbar.Tools("ID_GmajPro").Enabled = True Then
        strStandAlone = GmajProcess
        If (Len(strStandAlone) > 0) Then
            If FileExist(strStandAlone) Then
                If HasDotNet Then
                    frmMain.tbToolbar.Tools("ID_DanCodeWeb").Enabled = False
                    frmMain.tbToolbar.Tools("ID_GmajPro").Enabled = False
                    
                    ' DAJ 08/10/2010: If the output file from the Daniel Code application is
                    ' still hanging around for some reason when we get in here, kill it
                    ' so that the timer doesn't handle it...
                    If Len(strDcFile) > 0 Then
                        If FileExist(AddSlash(App.Path) & strDcFile) Then
                            KillFile AddSlash(App.Path) & strDcFile
                        End If
                    End If
                    
                    ' DAJ 09/30/2010: We are now going to send a file name over to the Daniel Code process
                    ' as a command line argument that can eventually have arguments in it...
                    strDcArgsFile = GetIniFileProperty("ArgsFile", "", "GmajPro", AddSlash(App.Path) & "Provided\Provided.INI")
                    If Len(strDcArgsFile) > 0 Then
                        strDcArgsFile = AddSlash(App.Path) & strDcArgsFile
                        
                        Set astrArgs = New cGdArray
                        astrArgs.Create eGDARRAY_Strings
                        If HasModule("GMAJPRO") Then
                            astrArgs.Add "GMAJPRO"
                        ElseIf HasModule("DCPROFX") Then
                            astrArgs.Add "DCPROFX"
                        ElseIf HasModule("DCPROFUT") Then
                            astrArgs.Add "DCPROFUT"
                        Else
                            astrArgs.Add ""
                        End If
                        
                        astrArgs.ToFile strDcArgsFile
                        
                        If InStr(strDcArgsFile, " ") > 0 Then
                            strDcArgsFile = Chr(34) & strDcArgsFile & Chr(34)
                        End If
                    End If
                
                    ' Run the stand-alone application...
                    DebugLog "Starting '" & strStandAlone & "' '" & strDcArgsFile & "'"
                    RunProcess strStandAlone, strDcArgsFile
                
                    ' Start the timer waiting for file to exist...
                    frmOnlineBroker.tmrGmaj.Enabled = True
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.StartGmajProcess"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GmajProcess
'' Description: Return the Daniel Code Gmaj process path and name
'' Inputs:      None
'' Returns:     Process Path/Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GmajProcess() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strName As String               ' Name of the product
    
    strReturn = ""
    strName = frmMain.tbToolbar.Tools("ID_GmajPro").Name
    
    If HasLevel(eTN4_Gold, , strName) Then
        strReturn = GetIniFileProperty("ExeName", "", "GmajPro", AddSlash(App.Path) & "Provided\Provided.INI")
        If Len(strReturn) > 0 Then
            strReturn = AddSlash(App.Path) & strReturn
        End If
    End If
    
    GmajProcess = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.GmajProcess"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeactivateTrading
'' Description: Deactivate trading related stuff
'' Inputs:      Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeactivateTrading(ByVal strReason As String)
On Error GoTo ErrSection:

    g.TsoGroups.DeactivateGroups strReason, , 10
    g.OrderStrategies.DeactivateExits strReason, True, , 10
    g.TradingItems.DisableTradeItems strReason, True
    HandleDemoOrders 15
    
    'g.SimTrade.PositionVerify = True
    g.SimTradeStream.Broker.PositionVerify = Not g.RealTime.Reconnecting

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.DeactivateTrading"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionToString
'' Description: Convert a position to a display string
'' Inputs:      Position, Blank if flat?
'' Returns:     Display String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionToString(ByVal lPosition As Long, Optional ByVal bBlankIfFlat As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If lPosition > 0 Then
        strReturn = "Long " & Str(lPosition)
    ElseIf lPosition < 0 Then
        strReturn = "Short " & Str(Abs(lPosition))
    ElseIf bBlankIfFlat = False Then
        strReturn = "Flat"
    End If
    
    PositionToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.PositionToString"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerMessageTypeToString
'' Description: Convert a message type enumeration to a string
'' Inputs:      Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerMessageTypeToString(ByVal nBrokerMessage As eGDBrokerMessageTypes) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nBrokerMessage
        Case eGDBrokerMessageType_AppLoaded
            strReturn = "eGDBrokerMessageType_AppLoaded"
        Case eGDBrokerMessageType_Connect
            strReturn = "eGDBrokerMessageType_Connect"
        Case eGDBrokerMessageType_ConnectionInfo
            strReturn = "eGDBrokerMessageType_ConnectionInfo"
        Case eGDBrokerMessageType_Disconnect
            strReturn = "eGDBrokerMessageType_Disconnect"
        Case eGDBrokerMessageType_AddOrder
            strReturn = "eGDBrokerMessageType_AddOrder"
        Case eGDBrokerMessageType_Order
            strReturn = "eGDBrokerMessageType_Order"
        Case eGDBrokerMessageType_AmendOrder
            strReturn = "eGDBrokerMessageType_AmendOrder"
        Case eGDBrokerMessageType_CancelOrder
            strReturn = "eGDBrokerMessageType_CancelOrder"
        Case eGDBrokerMessageType_UnloadApp
            strReturn = "eGDBrokerMessageType_UnloadApp"
        Case eGDBrokerMessageType_AppUnloaded
            strReturn = "eGDBrokerMessageType_AppUnloaded"
        Case eGDBrokerMessageType_GetAccounts
            strReturn = "eGDBrokerMessageType_GetAccounts"
        Case eGDBrokerMessageType_AccountRefresh
            strReturn = "eGDBrokerMessageType_AccountRefresh"
        Case eGDBrokerMessageType_GetOrders
            strReturn = "eGDBrokerMessageType_GetOrders"
        Case eGDBrokerMessageType_OrderRefresh
            strReturn = "eGDBrokerMessageType_OrderRefresh"
        Case eGDBrokerMessageType_GetFills
            strReturn = "eGDBrokerMessageType_GetFills"
        Case eGDBrokerMessageType_FillRefresh
            strReturn = "eGDBrokerMessageType_FillRefresh"
        Case eGDBrokerMessageType_GetPositions
            strReturn = "eGDBrokerMessageType_GetPositions"
        Case eGDBrokerMessageType_PositionRefresh
            strReturn = "eGDBrokerMessageType_PositionRefresh"
        Case eGDBrokerMessageType_Heartbeat
            strReturn = "eGDBrokerMessageType_Heartbeat"
        Case eGDBrokerMessageType_Fill
            strReturn = "eGDBrokerMessageType_Fill"
        Case eGDBrokerMessageType_Position
            strReturn = "eGDBrokerMessageType_Position"
        Case eGDBrokerMessageType_CarriedFillRefresh
            strReturn = "eGDBrokerMessageType_CarriedFillRefresh"
        Case eGDBrokerMessageType_Subscribe
            strReturn = "eGDBrokerMessageType_Subscribe"
        Case eGDBrokerMessageType_PriceUpdate
            strReturn = "eGDBrokerMessageType_PriceUpdate"
        Case eGDBrokerMessageType_Unsubscribe
            strReturn = "eGDBrokerMessageType_Unsubscribe"
        Case eGDBrokerMessageType_GetAccountDetails
            strReturn = "eGDBrokerMessageType_GetAccountDetails"
        Case eGDBrokerMessageType_AccountDetails
            strReturn = "eGDBrokerMessageType_AccountDetails"
        Case eGDBrokerMessageType_GetSecurityDefinition
            strReturn = "eGDBrokerMessageType_GetSecurityDefinition"
        Case eGDBrokerMessageType_UserRequest
            strReturn = "eGDBrokerMessageType_UserRequest"
        Case eGDBrokerMessageType_GetSides
            strReturn = "eGDBrokerMessageType_GetSides"
        Case eGDBrokerMessageType_Sides
            strReturn = "eGDBrokerMessageType_Sides"
        Case eGDBrokerMessageType_GetTifs
            strReturn = "eGDBrokerMessageType_GetTifs"
        Case eGDBrokerMessageType_Tifs
            strReturn = "eGDBrokerMessageType_Tifs"
        Case eGDBrokerMessageType_GetOrderTypes
            strReturn = "eGDBrokerMessageType_GetOrderTypes"
        Case eGDBrokerMessageType_OrderTypes
            strReturn = "eGDBrokerMessageType_OrderTypes"
        Case eGDBrokerMessageType_GetSymbols
            strReturn = "eGDBrokerMessageType_GetSymbols"
        Case eGDBrokerMessageType_Symbols
            strReturn = "eGDBrokerMessageType_Symbols"
        Case eGDBrokerMessageType_GetNumberOfAccounts
            strReturn = "eGDBrokerMessageType_GetNumberOfAccounts"
        Case eGDBrokerMessageType_NumberOfAccounts
            strReturn = "eGDBrokerMessageType_NumberOfAccounts"
        Case eGDBrokerMessageType_AddOcoOrders
            strReturn = "eGDBrokerMessageType_AddOcoOrders"
        Case eGDBrokerMessageType_SpreadFill
            strReturn = "eGDBrokerMessageType_SpreadFill"
        Case eGDBrokerMessageType_SpecialFill
            strReturn = "eGDBrokerMessageType_SpecialFill"
        Case eGDBrokerMessageType_RejectMessage
            strReturn = "eGDBrokerMessageType_RejectMessage"
        Case eGDBrokerMessageType_PositionFillRefresh
            strReturn = "eGDBrokerMessageType_PositionFillRefresh"
        Case eGDBrokerMessageType_ConsumerInfo
            strReturn = "eGDBrokerMessageType_ConsumerInfo"
        Case eGDBrokerMessageType_LoginUrl
            strReturn = "eGDBrokerMessageType_LoginUrl"
        Case eGDBrokerMessageType_GetTransactions
            strReturn = "eGDBrokerMessageType_GetTransactions"
        Case eGDBrokerMessageType_Transactions
            strReturn = "eGDBrokerMessageType_Transactions"
            
        Case eGDBrokerMessageType_OecOrderIdChanged
            strReturn = "eGDBrokerMessageType_OecOrderIdChanged"
        Case eGDBrokerMessageType_CnxGetWorkingOrders
            strReturn = "eGDBrokerMessageType_CnxGetWorkingOrders"
        Case eGDBrokerMessageType_CnxWorkingOrderRefresh
            strReturn = "eGDBrokerMessageType_CnxWorkingOrderRefresh"
        Case eGDBrokerMessageType_CnxGetSingleOrder
            strReturn = "eGDBrokerMessageType_CnxGetSingleOrder"
        Case eGDBrokerMessageType_CnxSingleOrderRefresh
            strReturn = "eGDBrokerMessageType_CnxSingleOrderRefresh"

        Case Else
            strReturn = Str(nBrokerMessage)
    End Select
    
    BrokerMessageTypeToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.BrokerMessageTypeToString"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerMessageTypeFromString
'' Description: Convert a string to a message type enumeration
'' Inputs:      String
'' Returns:     Message Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerMessageTypeFromString(ByVal strMessageType As String) As eGDBrokerMessageTypes
On Error GoTo ErrSection:
    
    Dim nReturn As eGDBrokerMessageTypes  ' Return value for the function
    
    Select Case strMessageType
        Case "eGDBrokerMessageType_AppLoaded"
            nReturn = eGDBrokerMessageType_AppLoaded
        Case "eGDBrokerMessageType_Connect"
            nReturn = eGDBrokerMessageType_Connect
        Case "eGDBrokerMessageType_ConnectionInfo"
            nReturn = eGDBrokerMessageType_ConnectionInfo
        Case "eGDBrokerMessageType_Disconnect"
            nReturn = eGDBrokerMessageType_Disconnect
        Case "eGDBrokerMessageType_AddOrder"
            nReturn = eGDBrokerMessageType_AddOrder
        Case "eGDBrokerMessageType_Order"
            nReturn = eGDBrokerMessageType_Order
        Case "eGDBrokerMessageType_AmendOrder"
            nReturn = eGDBrokerMessageType_AmendOrder
        Case "eGDBrokerMessageType_CancelOrder"
            nReturn = eGDBrokerMessageType_CancelOrder
        Case "eGDBrokerMessageType_UnloadApp"
            nReturn = eGDBrokerMessageType_UnloadApp
        Case "eGDBrokerMessageType_AppUnloaded"
            nReturn = eGDBrokerMessageType_AppUnloaded
        Case "eGDBrokerMessageType_GetAccounts"
            nReturn = eGDBrokerMessageType_GetAccounts
        Case "eGDBrokerMessageType_AccountRefresh"
            nReturn = eGDBrokerMessageType_AccountRefresh
        Case "eGDBrokerMessageType_GetOrders"
            nReturn = eGDBrokerMessageType_GetOrders
        Case "eGDBrokerMessageType_OrderRefresh"
            nReturn = eGDBrokerMessageType_OrderRefresh
        Case "eGDBrokerMessageType_GetFills"
            nReturn = eGDBrokerMessageType_GetFills
        Case "eGDBrokerMessageType_FillRefresh"
            nReturn = eGDBrokerMessageType_FillRefresh
        Case "eGDBrokerMessageType_GetPositions"
            nReturn = eGDBrokerMessageType_GetPositions
        Case "eGDBrokerMessageType_PositionRefresh"
            nReturn = eGDBrokerMessageType_PositionRefresh
        Case "eGDBrokerMessageType_Heartbeat"
            nReturn = eGDBrokerMessageType_Heartbeat
        Case "eGDBrokerMessageType_Fill"
            nReturn = eGDBrokerMessageType_Fill
        Case "eGDBrokerMessageType_Position"
            nReturn = eGDBrokerMessageType_Position
        Case "eGDBrokerMessageType_CarriedFillRefresh"
            nReturn = eGDBrokerMessageType_CarriedFillRefresh
        Case "eGDBrokerMessageType_Subscribe"
            nReturn = eGDBrokerMessageType_Subscribe
        Case "eGDBrokerMessageType_PriceUpdate"
            nReturn = eGDBrokerMessageType_PriceUpdate
        Case "eGDBrokerMessageType_Unsubscribe"
            nReturn = eGDBrokerMessageType_Unsubscribe
        Case "eGDBrokerMessageType_GetAccountDetails"
            nReturn = eGDBrokerMessageType_GetAccountDetails
        Case "eGDBrokerMessageType_AccountDetails"
            nReturn = eGDBrokerMessageType_AccountDetails
        Case "eGDBrokerMessageType_GetSecurityDefinition"
            nReturn = eGDBrokerMessageType_GetSecurityDefinition
        Case "eGDBrokerMessageType_UserRequest"
            nReturn = eGDBrokerMessageType_UserRequest
        Case "eGDBrokerMessageType_GetSides"
            nReturn = eGDBrokerMessageType_GetSides
        Case "eGDBrokerMessageType_Sides"
            nReturn = eGDBrokerMessageType_Sides
        Case "eGDBrokerMessageType_GetTifs"
            nReturn = eGDBrokerMessageType_GetTifs
        Case "eGDBrokerMessageType_Tifs"
            nReturn = eGDBrokerMessageType_Tifs
        Case "eGDBrokerMessageType_GetOrderTypes"
            nReturn = eGDBrokerMessageType_GetOrderTypes
        Case "eGDBrokerMessageType_OrderTypes"
            nReturn = eGDBrokerMessageType_OrderTypes
        Case "eGDBrokerMessageType_GetSymbols"
            nReturn = eGDBrokerMessageType_GetSymbols
        Case "eGDBrokerMessageType_Symbols"
            nReturn = eGDBrokerMessageType_Symbols
        Case "eGDBrokerMessageType_GetNumberOfAccounts"
            nReturn = eGDBrokerMessageType_GetNumberOfAccounts
        Case "eGDBrokerMessageType_NumberOfAccounts"
            nReturn = eGDBrokerMessageType_NumberOfAccounts
        Case "eGDBrokerMessageType_AddOcoOrders"
            nReturn = eGDBrokerMessageType_AddOcoOrders
        Case "eGDBrokerMessageType_SpreadFill"
            nReturn = eGDBrokerMessageType_SpreadFill
        Case "eGDBrokerMessageType_SpecialFill"
            nReturn = eGDBrokerMessageType_SpecialFill
        Case "eGDBrokerMessageType_RejectMessage"
            nReturn = eGDBrokerMessageType_RejectMessage
        Case "eGDBrokerMessageType_PositionFillRefresh"
            nReturn = eGDBrokerMessageType_PositionFillRefresh
        Case "eGDBrokerMessageType_ConsumerInfo"
            nReturn = eGDBrokerMessageType_ConsumerInfo
        Case "eGDBrokerMessageType_LoginUrl"
            nReturn = eGDBrokerMessageType_LoginUrl
        Case "eGDBrokerMessageType_GetTransactions"
            nReturn = eGDBrokerMessageType_GetTransactions
        Case "eGDBrokerMessageType_Transactions"
            nReturn = eGDBrokerMessageType_Transactions
            
        Case "eGDBrokerMessageType_OecOrderIdChanged"
            nReturn = eGDBrokerMessageType_OecOrderIdChanged
        Case "eGDBrokerMessageType_CnxGetWorkingOrders"
            nReturn = eGDBrokerMessageType_CnxGetWorkingOrders
        Case "eGDBrokerMessageType_CnxWorkingOrderRefresh"
            nReturn = eGDBrokerMessageType_CnxWorkingOrderRefresh
        Case "eGDBrokerMessageType_CnxGetSingleOrder"
            nReturn = eGDBrokerMessageType_CnxGetSingleOrder
        Case "eGDBrokerMessageType_CnxSingleOrderRefresh"
            nReturn = eGDBrokerMessageType_CnxSingleOrderRefresh
        
        Case Else
            nReturn = -1&
    End Select
    
    BrokerMessageTypeFromString = nReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.BrokerMessageTypeFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PostFixVersion57
'' Description: Post fix the database tables for version 57 update
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PostFixVersion57()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim AcctPos As cAccountPosition     ' Account Position object
    Dim bNotified As Boolean            ' Has the user been notified?
    
    bNotified = False
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] WHERE [FillMatchMode]>253;", dbOpenDynaset)
    Do While Not rs.EOF
        If bNotified = False Then
            frmSplash.Message -1, "Applying update to TradeTracker.MDB ..."
            bNotified = True
        End If
        
        rs.Edit
        rs!FillMatchMode = rs!FillMatchMode - 254
        rs.Update
        
        Set AcctPos = New cAccountPosition
        If AcctPos.Load(rs!AccountPositionID) Then
            AcctPos.RecalculateHistory
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.PostFixVersion57"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleSecondTimePositionMismatch
'' Description: Handle a second time position mismatch
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleSecondTimePositionMismatch(ByVal strBrokerName As String, ByVal lBrokerPos As Long, ByVal lPosition As Long, ByVal AcctPos As cAccountPosition)
On Error GoTo ErrSection:

    Dim frm As frmAlertPopup            ' Alert popup form
    Dim strAccount As String            ' Account number

    With AcctPos
        strAccount = g.Broker.AccountNumberForID(.AccountID)
        
        Set frm = New frmAlertPopup
        frm.ShowMessageBox strBrokerName & " is reporting that you are currently in a " & UCase(g.Broker.TextPosition(lBrokerPos)) & " position for " & .Symbol & " in account " & strAccount & ", but your fills for the day imply that you are in a " & UCase(g.Broker.TextPosition(lPosition)) & " position.||Because this inconsistency could cause incorrect orders to be placed, auto exits and automated trading strategies are being disabled for this symbol.||PLEASE CALL YOUR BROKER AND VERIFY YOUR POSITION IN THIS ACCOUNT.|", "Inconsistent Broker Information", vbCenter
        
        g.OrderStrategies.DeactivateExit .AccountID, .SymbolOrSymbolID, , "Position mismatch"
        g.TradingItems.DisableTradeItemsForSymbol .AccountID, .SymbolOrSymbolID, "Position mismatch", True
        g.Alerts.PositionMismatch .Symbol, strAccount
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.HandleSecondTimePositionMismatch"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PostFixVersion61
'' Description: Determine a symbol ID for a symbol in the date journals table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PostFixVersion61()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblDateJournals] WHERE [SymbolID]=" & Str(kNullData) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Edit
        rs!SymbolID = GetSymbolID(rs!Symbol)
        rs.Update
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.PostFixVersion61"
    
End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectionStatusString
'' Description: Determine the displayable string for the given connection status
'' Inputs:      Connection Status
'' Returns:     Display Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConnectionStatusString(ByVal nConnectionStatus As eGDConnectionStatus) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nConnectionStatus
        Case eGDConnectionStatus_Disconnected
            strReturn = "Disconnected"
        Case eGDConnectionStatus_Disconnecting
            strReturn = "Disconnecting"
        Case eGDConnectionStatus_Connecting
            strReturn = "Connecting"
        Case eGDConnectionStatus_Connected
            strReturn = "Connected"
    End Select
    
    ConnectionStatusString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ConnectionStatusString"
    
End Function
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTypeAllowedForOco
'' Description: Determine if the given order type is allowed for OCO ( non-Market )
'' Inputs:      Order Type
'' Returns:     True if allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderTypeAllowedForOco(ByVal nOrderType As eTT_OrderType) As Boolean
On Error GoTo ErrSection:

    OrderTypeAllowedForOco = (nOrderType <> eTT_OrderType_Market) And (nOrderType <> eTT_OrderType_MarketOnClose)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.OrderTypeAllowedForOco"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FirstSimStreamAccountID
'' Description: Get the first SimStream account ID out of the database
'' Inputs:      None
'' Returns:     First SimStream AccountID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FirstSimStreamAccountID() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    
    lReturn = 0&
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountType]=" & Str(eTT_AccountType_SimStream) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        lReturn = rs!AccountID
    End If
    
    FirstSimStreamAccountID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.FirstSimStreamAccountID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetTriggerOrderStatus
'' Description: Set the order status appropriately on a trigger order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetTriggerOrderStatus(Order As cPtOrder)
On Error GoTo ErrSection:

    If Order.HasTrigger Then
        If Order.IsConditional(False) Then
            If Order.AllRtDataAvailable Then
                Order.ChangeOrderStatus eTT_OrderStatus_TriggerPending
            Else
                Order.ChangeOrderStatus eTT_OrderStatus_DataPending
            End If
        Else
            Order.ChangeOrderStatus eTT_OrderStatus_TriggerPending
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SetTriggerOrderStatus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportChartForJournal
'' Description: Export the given chart for a journal
'' Inputs:      Chart, Chart Caption ( out )
'' Returns:     Filename
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExportChartForJournal(Chart As Form, Optional strCaption As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strFileName As String           ' Filename for the exported file
    Dim lIndex As Long                  ' Index into a for loop
    Dim activeCht As Form               ' Active chart form
    Dim rc As Integer                   ' Return code from the save chart call
    Dim lHeight As Long                 ' Height of the image in pixels
    Dim lWidth As Long                  ' Width of the image in pixels
    
    strReturn = ""
    rc = -1
    lIndex = 0&
    
    If DirExist(AddSlash(App.Path) & "SavedImages") = False Then
        MkDir AddSlash(App.Path) & "SavedImages"
    End If
    strFileName = AddSlash(App.Path) & "SavedImages\" & Format(CurrentTime, "YYYYMMDD HHMMSS") & ".JPG"
    Do While FileExist(strFileName)
        lIndex = lIndex + 1
        strFileName = AddSlash(App.Path) & "SavedImages\" & Format(CurrentTime, "YYYYMMDD HHMMSS") & " " & Format(lIndex, "0000") & ".JPG"
    Loop
        
    Set activeCht = ActiveChart
    If (activeCht.hWnd <> Chart.hWnd) And (activeCht.WindowState = vbMaximized) Then
        If Screen.TwipsPerPixelX <> 0 And Screen.TwipsPerPixelY <> 0 Then
            lHeight = activeCht.Height / Screen.TwipsPerPixelY
            lWidth = activeCht.Width / Screen.TwipsPerPixelX
            
            rc = Chart.cChartObj.LoadExportData(lWidth, lHeight)
            If rc = 0 Then
                rc = geSaveChart(Chart.cChartObj.geChartObj, Chart.pbChart.hWnd, Chart.pbChart.hDC, lWidth, lHeight, 2, strFileName)
            End If
        End If
    Else
        rc = geSaveChart(Chart.cChartObj.geChartObj, Chart.pbChart.hWnd, Chart.pbChart.hDC, 0, 0, 2, strFileName)
    End If
    
    If rc = 0 Then
        strReturn = strFileName
        strCaption = ChartCaptionForJournal(Chart)
    End If
    
    ExportChartForJournal = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ExportChartForJournal"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChartImageForFill
'' Description: Create a chart image for the given fill
'' Inputs:      Fill, Chart Caption ( out )
'' Returns:     Filename ( Blank if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ChartImageForFill(ByVal Fill As cPtFill, Optional strCaption As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strFileName As String           ' Filename for the chart image
    Dim lIndex As Long                  ' Index into a for loop

    strReturn = ""
    For lIndex = 0 To Forms.Count - 1
        If IsFrmChart(Forms(lIndex)) = True Then
            If ConvertToTradeSymbol(Forms(lIndex).SymbolOrSymbolID, Fill.FillDate) = Fill.SymbolOrSymbolID Then
                If Forms(lIndex).TradeAccountID = Fill.AccountID Then
                    If Forms(lIndex).Chart.LastGoodDataBar(False, False) = Forms(lIndex).Chart.LastGoodDataBar(False, True) Then
                        strReturn = ExportChartForJournal(Forms(lIndex), strCaption)
                    End If
                End If
            End If
        End If
    Next lIndex
    
    ChartImageForFill = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ChartImageForFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChartImageForHwnd
'' Description: Create a chart image for the given hWnd
'' Inputs:      Fill, Chart Caption ( out )
'' Returns:     Filename ( Blank if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ChartImageForHwnd(ByVal hWnd As Long, Optional strCaption As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strFileName As String           ' Filename for the chart image
    Dim lIndex As Long                  ' Index into a for loop

    strReturn = ""
    For lIndex = 0 To Forms.Count - 1
        If IsFrmChart(Forms(lIndex)) = True Then
            If Forms(lIndex).hWnd = hWnd Then
                strReturn = ExportChartForJournal(Forms(lIndex), strCaption)
            End If
        End If
    Next lIndex
    
    ChartImageForHwnd = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ChartImageForHwnd"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChartCaptionForJournal
'' Description: Build an image caption from the given form
'' Inputs:      Chart Form
'' Returns:     Image Caption
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ChartCaptionForJournal(ChartForm As Form) As String
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol for the chart
    Dim strPeriod As String             ' Period for the chart
    Dim strNow As String                ' Current time
    
    strSymbol = ChartForm.Chart.Symbol
    strPeriod = GetPeriodStr(ChartForm.Periodicity)
    strNow = Format(CurrentTime, "YYYY-MM-DD HH:MM:SS")
    
    ChartCaptionForJournal = strSymbol & " " & strPeriod & " " & strNow

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.ChartCaptionForJournal"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitializeBrokerBridge
'' Description: Initialize the Broker DLL bridge
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitializeBrokerBridge()
On Error GoTo ErrSection:

    Set g.BrokerEnums = New cBrokerEnums
    Set g.TnBroker = New cTnBroker
    
    Set g.BrokerBridge = New cBrokerBridge
    With g.BrokerBridge
        .AppBridge = g.TnBroker
        .AppPath = g.strAppPath
        .IniFile = g.strIniFile
        .TnCore = g.TnCore
        .TradingDatabase = g.dbPaper
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.InitializeBrokerBridge"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitializeJournalBridge
'' Description: Initialize the Journal DLL bridge
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitializeJournalBridge()
On Error GoTo ErrSection:

    Set g.TnJournal = New cTnJournal
    
    Set g.JournalBridge = New cJournalBridge
    With g.JournalBridge
        .AltGridRowColor = ALT_GRID_ROW_COLOR
        .AppBridge = g.TnJournal
        .AppIsIde = IsIDE
        .AppPath = g.strAppPath
        .IniFile = g.strIniFile
        .TnCore = g.TnCore
        .TradingDatabase = g.dbPaper
        
        If Not IsIDE Then
            .MainForm = frmMain
        End If
        
        .Init
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.InitializeJournalBridge"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitializeCattleBridge
'' Description: Initialize the Cattle DLL bridge
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitializeCattleBridge()
On Error GoTo ErrSection:

    'Set g.CattleEnums = New cCattleEnums
    Set g.TnCattle = New cTnCattle
    
    Set g.CattleBridge = New cCattleBridge
    With g.CattleBridge
        .AltGridRowColor = ALT_GRID_ROW_COLOR
        .AppBridge = g.TnCattle
        .AppPath = g.strAppPath
        .DataServiceID = RI_GetLastDataServiceID
        .Help = g.Help
        .IniFile = g.strIniFile
        If Not IsIDE Then
            .MainForm = frmMain
        End If
        If g.RealTime Is Nothing Then
            .StreamActive = False
        Else
            .StreamActive = g.RealTime.Active
        End If
        If FormIsLoaded("frmQuotes") Then
            .StreamInterval = frmQuotes.tmrRealTime.Interval
        Else
            .StreamInterval = 250
        End If
        
        .Init
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.InitializeCattleBridge"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetDllBridgeDatabases
'' Description: Set the DLL Bridge databases appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetDllBridgeDatabases()
On Error GoTo ErrSection:

    If Not g.BrokerBridge Is Nothing Then
        g.BrokerBridge.TradingDatabase = g.dbPaper
    End If
    
    If Not g.JournalBridge Is Nothing Then
        g.JournalBridge.TradingDatabase = g.dbPaper
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mTradeTracker.SetDllBridgeDatabases"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecordsetToBrokerMessage
'' Description: Convert a recordset to a broker message
'' Inputs:      Recordset, Broker Message to Append to
'' Returns:     Broker Message
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RecordsetToBrokerMessage(rs As Recordset, Optional brokerMessage As cBrokerMessage = Nothing) As cBrokerMessage
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    If brokerMessage Is Nothing Then
        Set brokerMessage = New cBrokerMessage
    End If
    
    If Not rs Is Nothing Then
        For lIndex = 0 To rs.Fields.Count - 1
            brokerMessage.Add rs.Fields(lIndex).Name, Str(rs.Fields(lIndex).Value)
        Next lIndex
    End If
    
    Set RecordsetToBrokerMessage = brokerMessage

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.RecordsetToBrokerMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToBrokerMessage
'' Description: Convert an order from a recordset to a broker message
'' Inputs:      Order ID, Order
'' Returns:     Broker Message
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderToBrokerMessage(ByVal lOrderID As Long, Optional Order As cPtOrder = Nothing) As cBrokerMessage
On Error GoTo ErrSection:

    Dim orderMessage As cBrokerMessage  ' Return value for the function
    Dim OrderLeg As cBrokerMessage      ' Order leg object
    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [OrderID]=" & Str(lOrderID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        Set orderMessage = RecordsetToBrokerMessage(rs)
        
        orderMessage.Add "AccountNumber", g.Broker.AccountNumberForID(rs!AccountID)
        orderMessage.Add "AccountName", g.Broker.AccountNameForID(rs!AccountID)
        
        If Order Is Nothing Then
            Set Order = New cPtOrder
            Order.Load lOrderID, rs
        End If
        
        orderMessage.Add "OrderText", Order.OrderText(True, True)
        If Order.NumberOfLegs = 1 Then
            orderMessage.Add "Symbol", Order.Symbol
        Else
            orderMessage.Add "Symbol", Order.SpreadSymbol
        End If
        orderMessage.Add "OptionNavImageFile", Order.OptionNavImageFile
    
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderLegs] WHERE [OrderID]=" & Str(lOrderID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            orderMessage.Add "Leg" & Str(rs!LegNumber), RecordsetToBrokerMessage(rs).ToString(False)
            
            rs.MoveNext
        Loop
    End If
    
    Set OrderToBrokerMessage = orderMessage
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.OrderToBrokerMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillDisplay
'' Description: Build a display string for the fill
'' Inputs:      Include Symbol?, Format Price?, Include Date?, Include Account?,
''              Include ID's?, Include Category Action?, Include Category Pnl?
'' Returns:     Display String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillDisplay(Fill As cPtFill, Optional ByVal bIncludeSymbol As Boolean = True, Optional ByVal bFormatPrice As Boolean = False, Optional ByVal bIncludeDate As Boolean = True, Optional ByVal bIncludeAccount As Boolean = True, Optional ByVal bIncludeIDs As Boolean = False, Optional ByVal bIncludeCategoryAction As Boolean = False, Optional ByVal bIncludeCategoryPnl As Boolean = False) As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Information to join together to return from the function
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    If Fill.Buy = True Then
        astrReturn.Add "Bought"
    Else
        astrReturn.Add "Sold"
    End If
    
    astrReturn.Add Str(Fill.Quantity)
    
    If bIncludeSymbol = True Then
        astrReturn.Add Fill.Symbol
    End If
    
    If bFormatPrice = True Then
        astrReturn.Add "at " & Fill.PriceString
    Else
        astrReturn.Add "at " & Str(Fill.Price)
    End If

    If bIncludeDate = True Then
        astrReturn.Add "(" & DateFormat(Fill.FillDate, MM_DD_YYYY, HH_MM_SS) & ", Session = " & DateFormat(Fill.SessionDate, MM_DD_YYYY) & ")"
    End If
    
    If bIncludeAccount = True Then
        astrReturn.Add "in account " & g.Broker.AccountNumberForID(Fill.AccountID)
    End If
    
    If bIncludeIDs = True Then
        astrReturn.Add "(" & Str(Fill.FillID) & ", '" & Fill.BrokerID & "', '" & Fill.PreviousBrokerID & "', " & Str(Fill.OrderID) & ", '" & Fill.BrokerOrderID & "')"
    End If
    
    If Len(Fill.ActionCategory) > 0 Then
        If bIncludeCategoryAction = True Then
            If Fill.ActionCategory = "E" Then
                astrReturn.Add "to enter"
            ElseIf Fill.ActionCategory = "X" Then
                astrReturn.Add "to exit"
            Else
                astrReturn.Add "to reverse"
            End If
        End If
                        
        If (bIncludeCategoryPnl = True) And (Fill.ActionCategory <> "E") Then
            If Fill.ClosedProfitCategory > 0 Then
                astrReturn.Add "for a profit of " & Format(Fill.ClosedProfitCategory, "$#,##0.00")
            ElseIf Fill.ClosedProfitCategory < 0 Then
                astrReturn.Add "for a loss of " & Format(Abs(Fill.ClosedProfitCategory), "$#,##0.00")
            Else
                astrReturn.Add "for no profit"
            End If
        End If
    End If
    
    FillDisplay = astrReturn.JoinFields(" ")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.FillDisplay"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LimitPriceForMarketOrder
'' Description: Determine the limit price to use for a market order
'' Inputs:      Buy or Sell?, Symbol, Bars
'' Returns:     Limit price to use for a market order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LimitPriceForMarketOrder(ByVal bBuy As Boolean, ByVal strSymbol As String, Optional ByVal Bars As cGdBars = Nothing) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim strKey As String                ' Key for an INI file property
    Dim lNumTicks As Long               ' Number of ticks away from the market
    Dim dPrice As Double                ' Market price to use
    
    If Bars Is Nothing Then
        Set Bars = New cGdBars
        SetBarProperties Bars, strSymbol
    End If
    
    strKey = Parse(strSymbol, "-", 1)
    If InStr(strSymbol, "-") > 0 Then
        strKey = strKey & "-"
    End If
    lNumTicks = Bars.TickMove * GetIniFileProperty(strKey, 5&, "TicksForLimit", AddSlash(App.Path) & "Provided\Provided.INI")
    
    dReturn = kNullData
    If bBuy Then
        ' Put the limit order for a buy at 5 ticks above the ask which should be safely above the market...
        dPrice = g.RealTime.LastKnownPrice(strSymbol, 1)
        If dPrice = kNullData Then
            dPrice = g.RealTime.LastKnownPrice(strSymbol)
        End If
        
        If dPrice <> kNullData Then
            dReturn = dPrice + lNumTicks
        End If
    Else
        ' Put the limit order for a sell at 5 ticks below the bid which should be safely above the market...
        dPrice = g.RealTime.LastKnownPrice(strSymbol, -1)
        If dPrice = kNullData Then
            dPrice = g.RealTime.LastKnownPrice(strSymbol)
        End If
        
        If dPrice <> kNullData Then
            dReturn = dPrice - lNumTicks
        End If
    End If
    
    LimitPriceForMarketOrder = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.LimitPriceForMarketOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NeedToChangeMarketToLimit
'' Description: Do we need to change a market order into a limit order?
'' Inputs:      Symbol or Symbol ID, Bars, Time to Check
'' Returns:     True if need to change, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NeedToChangeMarketToLimit(ByVal vSymbolOrSymbolID As Variant, Optional ByVal Bars As cGdBars = Nothing, Optional ByVal dTimeToCheck As Double = kNullData) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strKey As String                ' Key for an INI file property
    
    bReturn = False
    If SecurityType(vSymbolOrSymbolID) = "F" Then
        strKey = Parse(GetSymbol(vSymbolOrSymbolID), "-", 1) & "-"
        If GetIniFileProperty(strKey, 0&, "ChangeMarketEth", AddSlash(App.Path) & "Provided\Provided.INI") <> 0& Then
            bReturn = mDataNav.IsEth(vSymbolOrSymbolID, Bars, dTimeToCheck)
        End If
    End If
    
    NeedToChangeMarketToLimit = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mTradeTracker.NeedToChangeMarketToLimit"
    
End Function
