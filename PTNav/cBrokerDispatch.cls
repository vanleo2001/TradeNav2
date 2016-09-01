VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerDispatch.cls
'' Description: Class to act as a dispatcher for broker related calls
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/21/2009   DAJ         Added settings for whether or not to warn the user if
''                          they are about to place an order on the wrong side
''                          of the market
'' 05/19/2009   DAJ         Fixed Positions/Orders for base symbols not to return
''                          options for the given base symbol
'' 08/21/2009   DAJ         Make changes for Parking/Submitting/Cancelling OCOs
'' 08/25/2009   DAJ         Added support for PFG account information
'' 09/01/2009   DAJ         Use new Parked order status
'' 10/07/2009   DAJ         Added support for Linked Orders held at the broker
'' 10/07/2009   DAJ         Implement price shaving for filling options
'' 12/01/2009   DAJ         Added support for automatic commissions on fills
'' 12/01/2009   DAJ         Don't cancel auto exit orders that are held at broker
''                          upon Trade Navigator exit
'' 01/04/2010   DAJ         Enhancements for CanDisconnect function
'' 02/18/2010   DAJ         Added the EntriesForSymbol function
'' 03/11/2010   DAJ         Use global collections, activity log object, added properties
'' 05/06/2010   DAJ         Handle submitting parked order with triggered by orders (#5715)
'' 05/25/2010   DAJ         Fixes for setting BrokerOCO flag for non-existent account
'' 06/03/2010   DAJ         Changes for new Trade Sense Order Groups
'' 06/30/2010   DAJ         Option to park all/cancel all TradeSense order groups upon close
'' 07/05/2010   DAJ         Handle TradeSense order groups on open/close
'' 08/11/2010   DAJ         Finished CanTrade function
'' 09/13/2010   DAJ         Added code for Rithmic
'' 09/20/2010   DAJ         Fix for auto exits/TradeSense order groups when stream stops
'' 09/24/2010   DAJ         Added the IsRithmicBroker function
'' 09/29/2010   DAJ         Changed order confirmation flags
'' 11/01/2010   DAJ         Added Optimus, OpVest, and Vision (Rithmic Brokers)
'' 11/09/2010   DAJ         Renamed object from cBroker.cls
'' 11/18/2010   DAJ         Added Starting Genesis ID for Option Navigator
'' 12/10/2010   DAJ         Added Zen-Fire, Changed over to the IsBrokerUser function, Added cBroker stuff
'' 12/15/2010   DAJ         Added ShowAccountOnline function
'' 01/19/2011   DAJ         Changed InitBrokerObjects to only call Init on objects if they were Nothing
'' 01/26/2011   DAJ         Fixes for submitting a parked order
'' 03/07/2011   DAJ         Added Oec/Options Express, IB/I-Deal/Rithmic/Gain now utilize cBroker,
''                          Added Change Password for LindXpress, utilize cPtOrder.ChangeOrderStatus,
''                          Changed some Option Nav calls, New Main Menu calls,
'' 03/10/2011   DAJ         Added OrderForAccessID call
'' 04/05/2011   DAJ         Added SymbolInformation call
'' 05/11/2011   DAJ         Added the CheckSimTradeStatus and EnableCheckStatusMenu calls
'' 05/16/2011   DAJ         Set OrderDate correctly in case delayed streaming
'' 06/21/2011   DAJ         Separate out Simulated trading types, Last connection status change
'' 07/15/2011   DAJ         Return ID from create for auto trade, Added GetPositions
'' 07/19/2011   DAJ         Added new, edit, and delete fill capabilities
'' 07/21/2011   DAJ         Added FillSummaryForTradeItem function
'' 07/22/2011   DAJ         Added AccountExists function
'' 08/15/2011   DAJ         Tweaks for modifying simulated fills
'' 08/25/2011   DAJ         Moved some code from frmOnlineBroker, mods for CQG/TT
'' 09/20/2011   DAJ         Entry point for showing TT data
'' 09/26/2011   DAJ         TT Fixes
'' 09/28/2011   DAJ         Fixes for automated journaling for automated orders
'' 10/07/2011   DAJ         Don't bring up auto journal for orders coming from Option Nav
'' 10/21/2011   DAJ         Pass along previous auto trade ID into AddFill if changed
'' 10/31/2011   DAJ         Send Option Nav a message when account deleted
'' 11/02/2011   DAJ         Added Amp Trading and RJ O'Brien as CQG brokers
'' 11/18/2011   DAJ         Added IsBrokerUserOptions function
'' 12/02/2101   DAJ         Added RJO (PATS)
'' 12/06/2011   DAJ         Added data calls for RJ O'Brien (PATS)
'' 12/09/2011   DAJ         Expiration Date for Account Positions, Added GFT Forex & OptionsHouse
'' 12/13/2011   DAJ         Added Capital Trading Group for PATS and CQG
'' 12/14/2011   DAJ         Added Capital Trading Group and Fintec for PFG
'' 01/26/2012   DAJ         Upon manual manipulation of fills, refresh auto trade item position
'' 01/30/2012   DAJ         Option Nav Journal Image
'' 03/14/2012   DAJ         Added Alpari(Currenex), Alpari(PATS), Penson(Currenex), Penson(CQG)
'' 03/21/2012   DAJ         Fix for amending a partially filled order
'' 04/05/2012   DAJ         Added broker mode code
'' 05/31/2012   DAJ         Turnkey implementation
'' 06/11/2012   DAJ         Make Turnkey work with all brokers
'' 06/26/2012   DAJ         Added calls to save to/load from new symbol info table in Trade Tracker
'' 07/16/2012   DAJ         ZanerCqg, ZanerPats, ZanerRithmic, ZanerZenFire, KnightCnx, KnightCqg
'' 07/17/2012   DAJ         AlpariZenFire
'' 07/17/2012   DAJ         RobbinsCqg
'' 07/18/2012   DAJ         RCG (New PATS)
'' 07/19/2012   DAJ         Added IsRcgPats broker call
'' 07/27/2012   DAJ         Demo (PATS)
'' 08/03/2012   DAJ         Remove Gain, FXCM, Photon, OptionsHouse, Alaron, Cadent, Lotus, OptXpress, Oec, ManChicago, ManLondon, Robbins
'' 08/17/2012   DAJ         Added IsIbBroker and IsCurrenexBroker calls
'' 08/23/2012   DAJ         Born (PATS), RJO Hong Kong (PATS)
'' 08/29/2012   DAJ         Zaner (Currenex)
'' 09/11/2012   DAJ         Check Broker View for account if loaded, Carried Fill calls
'' 09/12/2012   DAJ         Added Currenex, FXDD (Currenex), and VanKar (Currenex)
'' 09/12/2012   DAJ         Removed Rosenthal (Old PATS), Changed Generic PATS to New PATS
'' 10/02/2012   DAJ         Modifications for allowing options for CQG
'' 10/03/2012   DAJ         Lot size for forex symbols in TradeSense order groups
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 12/11/2012   DAJ         Use the flatten queue for position reversals
'' 12/11/2012   DAJ         Handle broker status in automated trading item
'' 12/11/2012   DAJ         Vision (CQG)
'' 01/18/2013   DAJ         Broker held OCO for Interactive Brokers
'' 01/30/2013   DAJ         Live/Demo/Test modes for Turnkey
'' 02/22/2013   DAJ         Added the MaxOrderLegs function
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
'' 03/11/2013   DAJ         Changed enabled symbols message for CQG brokers
'' 03/22/2013   DAJ         Added the OpenPositionsForAccount call
'' 04/15/2013   DAJ         Check for broker connection before attempting a cancel
'' 04/16/2013   DAJ         Added the RefreshOrder call
'' 04/17/2013   DAJ         Added direction filter to CancelWorkingOrders call
'' 06/12/2013   DAJ         Symbol and quantity validation for automated trading items
'' 06/28/2013   DAJ         Web export settings
'' 07/19/2013   DAJ         New TransAct API
'' 07/30/2013   DAJ         Keep a collection of journal categories
'' 08/27/2013   DAJ         Tweak to the UpdateBrokerContractInfo function
'' 09/04/2013   DAJ         Extra logging in the UpdateBrokerContractInfo function
'' 09/06/2013   DAJ         Handle the forms count changing while updating broker contract info
'' 10/16/2013   DAJ         Removed PFG/Xpress/OrderLinks, Added Oec/FptOec/FptCqg
'' 10/24/2013   DAJ         Contract size override for forex
'' 11/15/2013   DAJ         Moved "IsTurnkey...User" stuff to cTurnkey
'' 01/09/2014   DAJ         Fixed bug where broker menus were not showing in main menu
'' 02/05/2014   DAJ         Fixes for deleting an account
'' 02/24/2014   DAJ         Pass allowable security types on broker message
'' 02/25/2014   DAJ         Confirm orders if turnkey form up and associated account
'' 03/07/2014   DAJ         Moved Cattle stuff into NavCattle.DLL
'' 05/20/2014   DAJ         Added ReplayLogFile function
'' 06/26/2014   DAJ         Added automated trading items to symbols on forms check
'' 06/26/2014   DAJ         Added active TradeSense order groups to symbols on forms check
'' 07/10/2014   DAJ         Ensure order snapshot flag is the same as the fill
''                          snapshot flag after fill edit; Enhanced delete fill confirmation
'' 08/11/2014   DAJ         New flag for how to calculate open equity on options
'' 08/13/2014   DAJ         Added Interactive Brokers to the Not Enabled for Symbol error
'' 08/22/2014   DAJ         Added E-Trade
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 09/02/2014   DAJ         Commented out frmAccountInfo stuff
'' 09/10/2014   DAJ         Consolidate the delete account code
'' 10/14/2014   DAJ         Option to disable all auto trade items when stopping stream
'' 10/15/2014   DAJ         Added ClosePosition to DisableTradeItems
'' 10/24/2014   DAJ         Fill Display
'' 10/29/2014   DAJ         Remove old synthetic order/MIT code
'' 12/10/2014   DAJ         New DateIsSnapshot routines
'' 01/21/2015   DAJ         Log when PositionsToVerify changes
'' 01/29/2015   DAJ         Changed the GetContractInfo call
'' 04/28/2015   DAJ         Allow user to cancel a market order on live account if 'Suspended' or 'PreSubmitted'
'' 09/14/2015   DAJ         Added Tradier
'' 01/08/2016   DAJ         Update the account information cache and automated trading items UI when an account is updated
'' 01/28/2016   DAJ         Moved account delete code here from cAccountsUI
'' 03/18/2016   DAJ         Added TD Ameritrade
'' 04/12/2016   DAJ         Setup Tradier for a maximum of four order legs
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'    Select Case nBroker
'        Case eTT_AccountType_AdvFut
'        Case eTT_AccountType_AlpariCurrenex
'        Case eTT_AccountType_AlpariPats
'        Case eTT_AccountType_AlpariZenFire
'        Case eTT_AccountType_Ameritrade
'        Case eTT_AccountType_AmpCqg
'        Case eTT_AccountType_BornPats
'        Case eTT_AccountType_CQG
'        Case eTT_AccountType_CtgCqg
'        Case eTT_AccountType_CtgPats
'        Case eTT_AccountType_CtgPfg
'        Case eTT_AccountType_Currenex
'        Case eTT_AccountType_DemoPats
'        Case eTT_AccountType_Etrade
'        Case eTT_AccountType_FintecPfg
'        Case eTT_AccountType_FptCqg
'        Case eTT_AccountType_FptOec
'        Case eTT_AccountType_FxddCurrenex
'        Case eTT_AccountType_Gft
'        Case eTT_AccountType_Ideal
'        Case eTT_AccountType_IntBrokers
'        Case eTT_AccountType_KnightCqg
'        Case eTT_AccountType_KnightCurrenex
'        Case eTT_AccountType_LindWaldock
'        Case eTT_AccountType_ManExpress
'        Case eTT_AccountType_Oec
'        Case eTT_AccountType_Optimus
'        Case eTT_AccountType_OpVest
'        Case eTT_AccountType_PATS
'        Case eTT_AccountType_PFG
'        Case eTT_AccountType_RcgPats
'        Case eTT_AccountType_Rithmic
'        Case eTT_AccountType_RjoCqg
'        Case eTT_AccountType_RjoHkPats
'        Case eTT_AccountType_RjoPats
'        Case eTT_AccountType_RobbinsCqg
'        Case eTT_AccountType_SimBroker
'        Case eTT_AccountType_SimReplay
'        Case eTT_AccountType_SimStream
'        Case eTT_AccountType_SimTrade
'        Case eTT_AccountType_Tradier
'        Case eTT_AccountType_TransAct
'        Case eTT_AccountType_TT
'        Case eTT_AccountType_VanKarCurrenex
'        Case eTT_AccountType_Vision
'        Case eTT_AccountType_VisionCqg
'        Case eTT_AccountType_ZanerCqg
'        Case eTT_AccountType_ZanerCurrenex
'        Case eTT_AccountType_ZanerPats
'        Case eTT_AccountType_ZanerRithmic
'        Case eTT_AccountType_ZanerZenFire
'        Case eTT_AccountType_ZenFire
'    End Select

Private Const kSfeDat As String = "Provided\Sfe.DAT"

Private Type mPrivate
    astrAccountTypes As cGdArray        ' Array of account types for quick lookup
    alSfeAllowed As cGdArray            ' Array of last connected dates for SFE enablements
    
    dLastStatusChange As Double         ' Time of the last connection status change on any broker
    bPositionsToVerify As Boolean       ' Are there positions to verify?
    
    bShowCents As Boolean               ' Show cents in currency columns?
    strGridFont As String               ' Grid font
    bDontAllowStopMove As Boolean       ' Don't allow stop move if too close to market
    bConfirmManual As Boolean           ' Does the user want confirmation before a manual order gets submitted?
    bConfirmTriggered As Boolean        ' Does the user want confirmation before a triggered order gets submitted?
    bConfirmTradeSense As Boolean       ' Does the user want confirmation before a TradeSense order gets submitted?
    bWarnStopWrongSide As Boolean       ' Warn user if a Stop is going in on the wrong side of the market
    bWarnLimitWrongSide As Boolean      ' Warn user if a Limit is going in on the wrong side of the market
    lNumTicksStopBuffer As Long         ' Buffer to define "too close to market" to move a stop order
    bAutoJournalPopUp As Boolean        ' Automatically show journal dialog for manual orders?
    bAutoJournalAutomated As Boolean    ' Automatically create journal entry for automated orders?
    nOptionFillMethod As eGDOptionFill  ' Method for filling option orders
    nOptionOpenEquity As eGDOptionOpenEquity ' What to use to calculate open equity on options
    
    strWebAccounts As String            ' Account numbers to export portfolio information to the web
    lWebFontSize As Long                ' Font size for portfolio website
End Type
Private m As mPrivate

Public Property Get ShowCents() As Boolean
    ShowCents = m.bShowCents
End Property
Public Property Let ShowCents(ByVal bShowCents As Boolean)
    m.bShowCents = bShowCents
    SetIniFileProperty "ShowCents", bShowCents, "TTSummary", g.strIniFile
End Property

Public Property Get GridFont() As String
    GridFont = m.strGridFont
End Property
Public Property Let GridFont(ByVal strGridFont As String)
    m.strGridFont = strGridFont
    SetIniFileProperty "TTSummary", strGridFont, "Fonts", g.strIniFile
End Property

Public Property Get DontAllowStopMove() As Boolean
    DontAllowStopMove = m.bDontAllowStopMove
End Property
Public Property Let DontAllowStopMove(ByVal bDontAllowStopMove As Boolean)
    m.bDontAllowStopMove = bDontAllowStopMove
    SetIniFileProperty "DontAllowStopMove", m.bDontAllowStopMove, "Trading", g.strIniFile
End Property

Public Property Get ConfirmManual() As Boolean
    ConfirmManual = m.bConfirmManual
End Property
Public Property Let ConfirmManual(ByVal bConfirmManual As Boolean)
    m.bConfirmManual = bConfirmManual
    SetIniFileProperty "ConfirmManual", m.bConfirmManual, "Trading", g.strIniFile
End Property

Public Property Get ConfirmTriggered() As Boolean
    ConfirmTriggered = m.bConfirmTriggered
End Property
Public Property Let ConfirmTriggered(ByVal bConfirmTriggered As Boolean)
    m.bConfirmTriggered = bConfirmTriggered
    SetIniFileProperty "ConfirmTriggered", m.bConfirmTriggered, "Trading", g.strIniFile
End Property

Public Property Get ConfirmTradeSense() As Boolean
    ConfirmTradeSense = m.bConfirmTradeSense
End Property
Public Property Let ConfirmTradeSense(ByVal bConfirmTradeSense As Boolean)
    m.bConfirmTradeSense = bConfirmTradeSense
    SetIniFileProperty "ConfirmTradeSense", m.bConfirmTradeSense, "Trading", g.strIniFile
End Property

Public Property Get WarnStopWrongSide() As Boolean
    WarnStopWrongSide = m.bWarnStopWrongSide
End Property
Public Property Let WarnStopWrongSide(ByVal bWarnStopWrongSide As Boolean)
    m.bWarnStopWrongSide = bWarnStopWrongSide
    SetIniFileProperty "WarnStopWrongSide", m.bWarnStopWrongSide, "Trading", g.strIniFile
End Property

Public Property Get WarnLimitWrongSide() As Boolean
    WarnLimitWrongSide = m.bWarnLimitWrongSide
End Property
Public Property Let WarnLimitWrongSide(ByVal bWarnLimitWrongSide As Boolean)
    m.bWarnLimitWrongSide = bWarnLimitWrongSide
    SetIniFileProperty "WarnLimitWrongSide", m.bWarnLimitWrongSide, "Trading", g.strIniFile
End Property

Public Property Get NumTicksStopBuffer() As Long
    NumTicksStopBuffer = m.lNumTicksStopBuffer
End Property
Public Property Let NumTicksStopBuffer(ByVal lNumTicksStopBuffer As Long)
    m.lNumTicksStopBuffer = lNumTicksStopBuffer
    SetIniFileProperty "NumTicksStopBuffer", m.lNumTicksStopBuffer, "Trading", g.strIniFile
End Property

Public Property Get AutoJournalPopUp() As Boolean
    AutoJournalPopUp = m.bAutoJournalPopUp
End Property
Public Property Let AutoJournalPopUp(ByVal bAutoJournalPopUp As Boolean)
    m.bAutoJournalPopUp = bAutoJournalPopUp
    SetIniFileProperty "AutoJournalPopUp", m.bAutoJournalPopUp, "Trading", g.strIniFile
End Property

Public Property Get AutoJournalAutomated() As Boolean
    AutoJournalAutomated = m.bAutoJournalAutomated
End Property
Public Property Let AutoJournalAutomated(ByVal bAutoJournalAutomated As Boolean)
    m.bAutoJournalAutomated = bAutoJournalAutomated
    SetIniFileProperty "AutoJournalAutomated", m.bAutoJournalAutomated, "Trading", g.strIniFile
End Property

Public Property Get OptionFillMethod() As eGDOptionFill
    OptionFillMethod = m.nOptionFillMethod
End Property
Public Property Let OptionFillMethod(ByVal nOptionFillMethod As eGDOptionFill)
    m.nOptionFillMethod = nOptionFillMethod
    SetIniFileProperty "OptionFillMethod", m.nOptionFillMethod, "Trading", g.strIniFile
End Property

Public Property Get OptionOpenEquity() As eGDOptionOpenEquity
    OptionOpenEquity = m.nOptionOpenEquity
End Property
Public Property Let OptionOpenEquity(ByVal nOptionOpenEquity As eGDOptionOpenEquity)
    m.nOptionOpenEquity = nOptionOpenEquity
    SetIniFileProperty "OptionOpenEquity", m.nOptionOpenEquity, "Trading", g.strIniFile
End Property

Public Property Get LastStatusChange() As Double
    LastStatusChange = m.dLastStatusChange
End Property
Public Property Let LastStatusChange(ByVal dLastStatusChange As Double)
    m.dLastStatusChange = dLastStatusChange
End Property

Public Property Get PositionsToVerify() As Boolean
    PositionsToVerify = m.bPositionsToVerify
End Property
Public Property Let PositionsToVerify(ByVal bPositionsToVerify As Boolean)
    If bPositionsToVerify <> m.bPositionsToVerify Then
        BrokerDebugAll vbTab & vbTab & "g.Broker.PositionsToVerify changed from " & Str(m.bPositionsToVerify) & " to " & Str(bPositionsToVerify)
        m.bPositionsToVerify = bPositionsToVerify
    End If
End Property

Public Property Get WebAccounts() As String
    WebAccounts = m.strWebAccounts
End Property
Public Property Let WebAccounts(ByVal strWebAccounts As String)
    If strWebAccounts <> m.strWebAccounts Then
        m.strWebAccounts = strWebAccounts
        SetIniFileProperty "Accounts", strWebAccounts, "WebPortfolio", g.strIniFile
    End If
End Property

Public Property Get WebFontSize() As Long
    WebFontSize = m.lWebFontSize
End Property
Public Property Let WebFontSize(ByVal lWebFontSize As Long)
    If lWebFontSize <> m.lWebFontSize Then
        m.lWebFontSize = lWebFontSize
        SetIniFileProperty "FontSize", lWebFontSize, "WebPortfolio", g.strIniFile
    End If
End Property

Public Property Get HoldOcoAtBroker(ByVal vAccountNumberOrID As Variant) As Boolean
On Error GoTo ErrSection:

    Dim Acct As cPtAccount              ' Account object
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    Set Acct = Account(vAccountNumberOrID)
    If Not Acct Is Nothing Then
        bReturn = Acct.HoldOcoAtBroker
    End If
    
    HoldOcoAtBroker = bReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerDispatch.HoldOcoAtBroker.Get"
    
End Property
Public Property Let HoldOcoAtBroker(ByVal vAccountNumberOrID As Variant, ByVal bHoldOcoAtBroker As Boolean)
On Error GoTo ErrSection:

    Dim Acct As cPtAccount              ' Account object
    
    Set Acct = Account(vAccountNumberOrID)
    If Not Acct Is Nothing Then
        Acct.HoldOcoAtBroker = bHoldOcoAtBroker
        Acct.Save
        
        UpdateAccount Acct
    End If
    
    ' If it was the account number passed in, save off in the INI in case we need it later...
    If VarType(vAccountNumberOrID) = vbString Then
        SetIniFileProperty Str(vAccountNumberOrID), bHoldOcoAtBroker, "BrokerOCO", g.strIniFile
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerDispatch.HoldOcoAtBroker.Let"
    
End Property

Public Property Get SfeAllowed(ByVal nBroker As eTT_AccountType) As Long
    SfeAllowed = m.alSfeAllowed(nBroker)
End Property
Public Property Let SfeAllowed(ByVal nBroker As eTT_AccountType, ByVal lLastConnected As Long)
On Error GoTo ErrSection:

    Dim bSfeAllowed As Boolean          ' Is the SFE data currently allowed to be seen?

    If lLastConnected <> m.alSfeAllowed(nBroker) Then
        bSfeAllowed = mDataNav.SfeAllowed
        
        BrokerDebug nBroker, vbTab & "SfeAllowed changed from " & Str(m.alSfeAllowed(nBroker)) & " to " & Str(lLastConnected)
        m.alSfeAllowed(nBroker) = lLastConnected
        m.alSfeAllowed.Serialize AddSlash(App.Path) & kSfeDat, True
        
        If (bSfeAllowed = False) And (mDataNav.SfeAllowed = True) Then
        End If
    
        DebugLog "SfeAllowed = " & m.alSfeAllowed.JoinFields(",")
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerDispatch.SfeAllowed.Let"
    
End Property

Public Property Get PresetQuantity(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lNumber As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the property
    Dim strSecType As String            ' Security type for the given symbol
    Dim lDefaultQty As Long             ' Default quantity for the given security type
    
    strSecType = TradeSecType(vSymbolOrSymbolID)
    
    lReturn = GetIniFileProperty("Preset_" & strSecType & "_" & Str(lNumber), 0&, "Trading", g.strIniFile)
    If lReturn = 0& Then
        If strSecType = "F" Then
            lReturn = GetIniFileProperty("OrderQty" & Str(lNumber), 0&, "Charting", g.strIniFile)
        ElseIf strSecType = "S" Then
            lReturn = GetIniFileProperty("OrderQtyStocks" & Str(lNumber), 0&, "Charting", g.strIniFile)
        End If
        
        If lReturn = 0& Then
            lDefaultQty = DefaultOrderQuantity(vAccountNumberOrID, vSymbolOrSymbolID)
            Select Case lNumber
                Case 1:
                    lReturn = lDefaultQty
                Case 2:
                    lReturn = lDefaultQty * 2
                Case 3:
                    lReturn = lDefaultQty * 5
            End Select
        End If
    End If
    
    PresetQuantity = lReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerDispatch.PresetQuantity.Get"
    
End Property

Public Property Let PresetQuantity(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lNumber As Long, ByVal lPresetQuantity As Long)
On Error GoTo ErrSection:

    Dim strSecType As String            ' Security type for the given symbol
    
    strSecType = TradeSecType(vSymbolOrSymbolID)
    SetIniFileProperty "Preset_" & strSecType & "_" & Str(lNumber), lPresetQuantity, "Trading", g.strIniFile

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerDispatch.PresetQuantity.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountTypeForID
'' Description: Retrieve the account type for the given account ID
'' Inputs:      Account ID
'' Returns:     Account Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountTypeForID(ByVal lAccountID As Long) As eTT_AccountType
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the account ID found in the array?
    Dim rs As Recordset                 ' Recordset into the database
    Dim nReturn As eTT_AccountType      ' Return value from the function
    
    bFound = False
    For lIndex = 0 To m.astrAccountTypes.Size - 1
        If Parse(m.astrAccountTypes(lIndex), vbTab, 1) = Str(lAccountID) Then
            bFound = True
            nReturn = CLng(Val(Parse(m.astrAccountTypes(lIndex), vbTab, 3)))
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        If lAccountID = -10& Then
            m.astrAccountTypes.Add Str(-10&) & vbTab & "Alaron Broker Management" & vbTab & Str(8) & vbTab & "Alaron Broker Management"
            nReturn = 8
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                        "WHERE [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
            If Not (rs.BOF And rs.EOF) Then
                m.astrAccountTypes.Add Str(rs!AccountID) & vbTab & rs!AccountNumber & vbTab & Str(rs!AccountType) & vbTab & rs!Name
                nReturn = rs!AccountType
            Else
                nReturn = -1&
            End If
        End If
    End If
    
    AccountTypeForID = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AccountTypeForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountTypeForNumber
'' Description: Retrieve the account type for the given account number
'' Inputs:      Account Number
'' Returns:     Account Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountTypeForNumber(ByVal strAccountNumber As String) As eTT_AccountType
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the account ID found in the array?
    Dim rs As Recordset                 ' Recordset into the database
    Dim nReturn As eTT_AccountType      ' Return value from the function
    
    bFound = False
    For lIndex = 0 To m.astrAccountTypes.Size - 1
        If UCase(Parse(m.astrAccountTypes(lIndex), vbTab, 2)) = UCase(strAccountNumber) Then
            bFound = True
            nReturn = CLng(Val(Parse(m.astrAccountTypes(lIndex), vbTab, 3)))
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        If strAccountNumber = "Alaron Broker Management" Then
            m.astrAccountTypes.Add Str(-10&) & vbTab & "Alaron Broker Management" & vbTab & Str(8) & vbTab & "Alaron Broker Management"
            nReturn = 8
            bFound = True
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                        "WHERE [AccountNumber]='" & strAccountNumber & "';", dbOpenDynaset)
            If Not (rs.BOF And rs.EOF) Then
                m.astrAccountTypes.Add Str(rs!AccountID) & vbTab & rs!AccountNumber & vbTab & Str(rs!AccountType) & vbTab & rs!Name
                nReturn = rs!AccountType
                bFound = True
            Else
                nReturn = -1&
            End If
        End If
    End If
    
    If bFound = False Then
        If FormIsLoaded("frmBrokerView") Then
            If frmBrokerView.HasAccountNumber(strAccountNumber) Then
                nReturn = frmBrokerView.Broker
            End If
        End If
    End If
    
    AccountTypeForNumber = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AccountTypeForNumber"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountTypeForName
'' Description: Retrieve the account type for the given account name
'' Inputs:      Account Name
'' Returns:     Account Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountTypeForName(ByVal strAccountName As String) As eTT_AccountType
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the account ID found in the array?
    Dim rs As Recordset                 ' Recordset into the database
    Dim nReturn As eTT_AccountType      ' Return value from the function
    
    bFound = False
    For lIndex = 0 To m.astrAccountTypes.Size - 1
        If UCase(Parse(m.astrAccountTypes(lIndex), vbTab, 4)) = UCase(strAccountName) Then
            bFound = True
            nReturn = CLng(Val(Parse(m.astrAccountTypes(lIndex), vbTab, 3)))
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        If strAccountName = "Alaron Broker Management" Then
            m.astrAccountTypes.Add Str(-10&) & vbTab & "Alaron Broker Management" & vbTab & Str(8) & vbTab & "Alaron Broker Management"
            nReturn = 8
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                        "WHERE [Name]='" & strAccountName & "';", dbOpenDynaset)
            If Not (rs.BOF And rs.EOF) Then
                m.astrAccountTypes.Add Str(rs!AccountID) & vbTab & rs!AccountNumber & vbTab & Str(rs!AccountType) & vbTab & rs!Name
                nReturn = rs!AccountType
            Else
                nReturn = -1&
            End If
        End If
    End If
    
    AccountTypeForName = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AccountTypeForName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountNumberForID
'' Description: Retrieve the account number for the given account ID
'' Inputs:      Account ID
'' Returns:     Account Number
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountNumberForID(ByVal lAccountID As Long) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the account ID found in the array?
    Dim rs As Recordset                 ' Recordset into the database
    Dim strReturn As String             ' Return value from the function
    
    bFound = False
    For lIndex = 0 To m.astrAccountTypes.Size - 1
        If Parse(m.astrAccountTypes(lIndex), vbTab, 1) = Str(lAccountID) Then
            bFound = True
            strReturn = Parse(m.astrAccountTypes(lIndex), vbTab, 2)
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                    "WHERE [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            m.astrAccountTypes.Add Str(rs!AccountID) & vbTab & rs!AccountNumber & vbTab & Str(rs!AccountType) & vbTab & rs!Name
            strReturn = rs!AccountNumber
        Else
            strReturn = ""
        End If
    End If
    
    AccountNumberForID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AccountNumberForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountNumberForName
'' Description: Retrieve the account number for the given account name
'' Inputs:      Account Name
'' Returns:     Account Number
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountNumberForName(ByVal strAccountName As String) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the account ID found in the array?
    Dim rs As Recordset                 ' Recordset into the database
    Dim strReturn As String             ' Return value from the function
    
    bFound = False
    For lIndex = 0 To m.astrAccountTypes.Size - 1
        If UCase(Parse(m.astrAccountTypes(lIndex), vbTab, 4)) = UCase(strAccountName) Then
            bFound = True
            strReturn = Parse(m.astrAccountTypes(lIndex), vbTab, 2)
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                    "WHERE [Name]='" & strAccountName & "';", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            m.astrAccountTypes.Add Str(rs!AccountID) & vbTab & rs!AccountNumber & vbTab & Str(rs!AccountType) & vbTab & rs!Name
            strReturn = rs!AccountNumber
        Else
            strReturn = ""
        End If
    End If
    
    AccountNumberForName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AccountNumberForName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountIDForNumber
'' Description: Retrieve the account ID for the given account number
'' Inputs:      Account Number
'' Returns:     Account ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountIDForNumber(ByVal strAccountNumber As String) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the account ID found in the array?
    Dim rs As Recordset                 ' Recordset into the database
    Dim lReturn As Long                 ' Return value from the function
    
    bFound = False
    For lIndex = 0 To m.astrAccountTypes.Size - 1
        If UCase(Parse(m.astrAccountTypes(lIndex), vbTab, 2)) = UCase(strAccountNumber) Then
            bFound = True
            lReturn = CLng(Val(Parse(m.astrAccountTypes(lIndex), vbTab, 1)))
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                    "WHERE [AccountNumber]='" & strAccountNumber & "';", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            m.astrAccountTypes.Add Str(rs!AccountID) & vbTab & rs!AccountNumber & vbTab & Str(rs!AccountType) & vbTab & rs!Name
            lReturn = rs!AccountID
        Else
            lReturn = -1&
        End If
    End If
    
    AccountIDForNumber = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AccountIDForNumber"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountIDForName
'' Description: Retrieve the account ID for the given account name
'' Inputs:      Account Name
'' Returns:     Account ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountIDForName(ByVal strAccountName As String) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the account ID found in the array?
    Dim rs As Recordset                 ' Recordset into the database
    Dim lReturn As Long                 ' Return value from the function
    
    bFound = False
    For lIndex = 0 To m.astrAccountTypes.Size - 1
        If UCase(Parse(m.astrAccountTypes(lIndex), vbTab, 4)) = UCase(strAccountName) Then
            bFound = True
            lReturn = CLng(Val(Parse(m.astrAccountTypes(lIndex), vbTab, 1)))
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                    "WHERE [Name]='" & strAccountName & "';", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            m.astrAccountTypes.Add Str(rs!AccountID) & vbTab & rs!AccountNumber & vbTab & Str(rs!AccountType) & vbTab & rs!Name
            lReturn = rs!AccountID
        Else
            lReturn = -1&
        End If
    End If
    
    AccountIDForName = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AccountIDForName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountNameForID
'' Description: Retrieve the account name for the given account ID
'' Inputs:      Account ID
'' Returns:     Account Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountNameForID(ByVal lAccountID As Long) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the account ID found in the array?
    Dim rs As Recordset                 ' Recordset into the database
    Dim strReturn As String             ' Return value from the function
    
    bFound = False
    For lIndex = 0 To m.astrAccountTypes.Size - 1
        If Parse(m.astrAccountTypes(lIndex), vbTab, 1) = Str(lAccountID) Then
            bFound = True
            strReturn = Parse(m.astrAccountTypes(lIndex), vbTab, 4)
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                    "WHERE [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            m.astrAccountTypes.Add Str(rs!AccountID) & vbTab & rs!AccountNumber & vbTab & Str(rs!AccountType) & vbTab & rs!Name
            strReturn = rs!Name
        Else
            strReturn = ""
        End If
    End If
    
    AccountNameForID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AccountNameForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountNameForNumber
'' Description: Retrieve the account name for the given account number
'' Inputs:      Account Number
'' Returns:     Account Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountNameForNumber(ByVal strAccountNumber As String) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the account ID found in the array?
    Dim rs As Recordset                 ' Recordset into the database
    Dim strReturn As String             ' Return value from the function
    
    bFound = False
    For lIndex = 0 To m.astrAccountTypes.Size - 1
        If UCase(Parse(m.astrAccountTypes(lIndex), vbTab, 2)) = UCase(strAccountNumber) Then
            bFound = True
            strReturn = Parse(m.astrAccountTypes(lIndex), vbTab, 4)
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                    "WHERE [AccountNumber]='" & strAccountNumber & "';", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            m.astrAccountTypes.Add Str(rs!AccountID) & vbTab & rs!AccountNumber & vbTab & Str(rs!AccountType) & vbTab & rs!Name
            strReturn = rs!Name
        Else
            strReturn = ""
        End If
    End If
    
    AccountNameForNumber = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AccountNameForNumber"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccountID
'' Description: Pass back the account ID for the given ID or number
'' Inputs:      Account ID or Account Number
'' Returns:     Account ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAccountID(ByVal vAccountIDorNumber As Variant) As Long
On Error GoTo ErrSection:

    If VarType(vAccountIDorNumber) = vbString Then
        GetAccountID = AccountIDForNumber(Str(vAccountIDorNumber))
    Else
        GetAccountID = CLng(vAccountIDorNumber)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.GetAccountID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccountNumber
'' Description: Pass back the account Number for the given ID or number
'' Inputs:      Account ID or Account Number
'' Returns:     Account Number
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAccountNumber(ByVal vAccountIDorNumber As Variant) As String
On Error GoTo ErrSection:

    If VarType(vAccountIDorNumber) = vbString Then
        GetAccountNumber = Str(vAccountIDorNumber)
    Else
        GetAccountNumber = AccountNumberForID(CLng(vAccountIDorNumber))
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.GetAccountID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountExists
'' Description: Determine if the account exists in the database
'' Inputs:      Account
'' Returns:     True if exists in database, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountExists(ByVal vAccountNumberOrID As Variant) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lAccountID As Long              ' Account ID
    Dim rs As Recordset                 ' Recordset into the database
    
    bReturn = False
    lAccountID = GetAccountID(vAccountNumberOrID)
    If lAccountID > 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
        bReturn = Not (rs.BOF And rs.EOF)
    End If
    
    AccountExists = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AccountExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Broker
'' Description: Return the broker base class object for the given broker
'' Inputs:      Broker
'' Returns:     Broker Base Class object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Broker(ByVal nBroker As eTT_AccountType) As cBroker
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Return value for the function
    
    Set BrokerObj = Nothing
    Select Case nBroker
        Case eTT_AccountType_AdvFut
            If Not g.AdvFut Is Nothing Then
                Set BrokerObj = g.AdvFut.Broker
            End If
            
        Case eTT_AccountType_AlpariCurrenex
            If Not g.AlpariCnx Is Nothing Then
                Set BrokerObj = g.AlpariCnx.Broker
            End If
        
        Case eTT_AccountType_AlpariPats
            If Not g.AlpariPats Is Nothing Then
                Set BrokerObj = g.AlpariPats.Broker
            End If
        
        Case eTT_AccountType_AlpariZenFire
            If Not g.AlpariZenFire Is Nothing Then
                Set BrokerObj = g.AlpariZenFire.Broker
            End If
            
        Case eTT_AccountType_Ameritrade
            If Not g.Ameritrade Is Nothing Then
                Set BrokerObj = g.Ameritrade.Broker
            End If
    
        Case eTT_AccountType_AmpCqg
            If Not g.AmpCqg Is Nothing Then
                Set BrokerObj = g.AmpCqg.Broker
            End If
    
        Case eTT_AccountType_BornPats
            If Not g.BornPats Is Nothing Then
                Set BrokerObj = g.BornPats.Broker
            End If
    
        Case eTT_AccountType_CQG
            If Not g.CQG Is Nothing Then
                Set BrokerObj = g.CQG.Broker
            End If
            
        Case eTT_AccountType_CtgCqg
            If Not g.CtgCqg Is Nothing Then
                Set BrokerObj = g.CtgCqg.Broker
            End If
    
        Case eTT_AccountType_CtgPats
            If Not g.CtgPats Is Nothing Then
                Set BrokerObj = g.CtgPats.Broker
            End If
    
        Case eTT_AccountType_Currenex
            If Not g.Currenex Is Nothing Then
                Set BrokerObj = g.Currenex.Broker
            End If
            
        Case eTT_AccountType_DemoPats
            If Not g.DemoPats Is Nothing Then
                Set BrokerObj = g.DemoPats.Broker
            End If
            
        Case eTT_AccountType_Etrade
            If Not g.Etrade Is Nothing Then
                Set BrokerObj = g.Etrade.Broker
            End If
    
        Case eTT_AccountType_FptCqg
            If Not g.FptCqg Is Nothing Then
                Set BrokerObj = g.FptCqg.Broker
            End If
            
        Case eTT_AccountType_FptOec
            If Not g.FptOec Is Nothing Then
                Set BrokerObj = g.FptOec.Broker
            End If
            
        Case eTT_AccountType_FxddCurrenex
            If Not g.FxddCnx Is Nothing Then
                Set BrokerObj = g.FxddCnx.Broker
            End If
            
        Case eTT_AccountType_Gft
            If Not g.Gft Is Nothing Then
                Set BrokerObj = g.Gft.Broker
            End If
    
        Case eTT_AccountType_Ideal
            If Not g.Ideal Is Nothing Then
                Set BrokerObj = g.Ideal.Broker
            End If
        
        Case eTT_AccountType_IntBrokers
            If Not g.IntBroker Is Nothing Then
                Set BrokerObj = g.IntBroker.Broker
            End If
        
        Case eTT_AccountType_KnightCqg
            If Not g.KnightCqg Is Nothing Then
                Set BrokerObj = g.KnightCqg.Broker
            End If
        
        Case eTT_AccountType_KnightCurrenex
            If Not g.KnightCnx Is Nothing Then
                Set BrokerObj = g.KnightCnx.Broker
            End If
            
        Case eTT_AccountType_Oec
            If Not g.Oec Is Nothing Then
                Set BrokerObj = g.Oec.Broker
            End If
            
        Case eTT_AccountType_Optimus
            If Not g.Optimus Is Nothing Then
                Set BrokerObj = g.Optimus.Broker
            End If
            
        Case eTT_AccountType_OpVest
            If Not g.OpVest Is Nothing Then
                Set BrokerObj = g.OpVest.Broker
            End If
            
        Case eTT_AccountType_PATS
            If Not g.Pats Is Nothing Then
                Set BrokerObj = g.Pats.Broker
            End If
        
        Case eTT_AccountType_RcgPats
            If Not g.RcgPats Is Nothing Then
                Set BrokerObj = g.RcgPats.Broker
            End If
            
        Case eTT_AccountType_Rithmic
            If Not g.Rithmic Is Nothing Then
                Set BrokerObj = g.Rithmic.Broker
            End If
            
        Case eTT_AccountType_RjoCqg
            If Not g.RjoCqg Is Nothing Then
                Set BrokerObj = g.RjoCqg.Broker
            End If
    
        Case eTT_AccountType_RjoHkPats
            If Not g.RjoHkPats Is Nothing Then
                Set BrokerObj = g.RjoHkPats.Broker
            End If
    
        Case eTT_AccountType_RjoPats
            If Not g.RjoPats Is Nothing Then
                Set BrokerObj = g.RjoPats.Broker
            End If
    
        Case eTT_AccountType_RobbinsCqg
            If Not g.RobbinsCqg Is Nothing Then
                Set BrokerObj = g.RobbinsCqg.Broker
            End If
            
        Case eTT_AccountType_SimStream
            If Not g.SimTradeStream Is Nothing Then
                Set BrokerObj = g.SimTradeStream.Broker
            End If
            
        Case eTT_AccountType_SimReplay
            If Not g.SimTradeReplay Is Nothing Then
                Set BrokerObj = g.SimTradeReplay.Broker
            End If
            
        Case eTT_AccountType_SimBroker
            If Not g.SimTradeTs Is Nothing Then
                Set BrokerObj = g.SimTradeTs.Broker
            End If
            
        Case eTT_AccountType_Tradier
            If Not g.Tradier Is Nothing Then
                Set BrokerObj = g.Tradier.Broker
            End If
    
        Case eTT_AccountType_TT
            If Not g.TT Is Nothing Then
                Set BrokerObj = g.TT.Broker
            End If
            
        Case eTT_AccountType_VanKarCurrenex
            If Not g.VanKarCnx Is Nothing Then
                Set BrokerObj = g.VanKarCnx.Broker
            End If
            
        Case eTT_AccountType_Vision
            If Not g.Vision Is Nothing Then
                Set BrokerObj = g.Vision.Broker
            End If
            
        Case eTT_AccountType_VisionCqg
            If Not g.VisionCqg Is Nothing Then
                Set BrokerObj = g.VisionCqg.Broker
            End If
        
        Case eTT_AccountType_ZanerCqg
            If Not g.ZanerCqg Is Nothing Then
                Set BrokerObj = g.ZanerCqg.Broker
            End If
            
        Case eTT_AccountType_ZanerCurrenex
            If Not g.ZanerCnx Is Nothing Then
                Set BrokerObj = g.ZanerCnx.Broker
            End If
            
        Case eTT_AccountType_ZanerPats
            If Not g.ZanerPats Is Nothing Then
                Set BrokerObj = g.ZanerPats.Broker
            End If
            
        Case eTT_AccountType_ZanerRithmic
            If Not g.ZanerRithmic Is Nothing Then
                Set BrokerObj = g.ZanerRithmic.Broker
            End If
            
        Case eTT_AccountType_ZanerZenFire
            If Not g.ZanerZenFire Is Nothing Then
                Set BrokerObj = g.ZanerZenFire.Broker
            End If
            
        Case eTT_AccountType_ZenFire
            If Not g.ZenFire Is Nothing Then
                Set BrokerObj = g.ZenFire.Broker
            End If
            
    End Select
    
    Set Broker = BrokerObj

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cBrokerDispatch.Broker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsBrokerUser
'' Description: Determine if the user is entitled for the given broker
'' Inputs:      Broker, Is Broker?
'' Returns:     True if entitled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsBrokerUser(ByVal nBroker As eTT_AccountType, Optional bIsBroker As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bIsBroker = False
    bReturn = False
    Select Case nBroker
        Case eTT_AccountType_AdvFut
            bReturn = HasModule("B_ADVF") Or FileExist(AddSlash(App.Path) & "AdvFut.FLG")
        Case eTT_AccountType_AlpariCurrenex
            bReturn = HasModule("B_ALPCNX") Or FileExist(AddSlash(App.Path) & "AlpCnx.FLG")
        Case eTT_AccountType_AlpariPats
            bReturn = HasModule("B_ALPPAT") Or FileExist(AddSlash(App.Path) & "AlpPat.FLG")
        Case eTT_AccountType_AlpariZenFire
            bReturn = HasModule("B_ALPZEN") Or FileExist(AddSlash(App.Path) & "AlpZen.FLG")
        Case eTT_AccountType_Ameritrade
            bReturn = HasModule("B_TDA") Or FileExist(AddSlash(App.Path) & "TDA.FLG")
        Case eTT_AccountType_AmpCqg
            bReturn = HasModule("B_AMPCQG") Or FileExist(AddSlash(App.Path) & "AmpCqg.FLG")
        Case eTT_AccountType_BornPats
            bReturn = HasModule("B_BORNPAT") Or FileExist(AddSlash(App.Path) & "BornPats.FLG")
        Case eTT_AccountType_CQG
            bReturn = HasModule("B_CQG") Or FileExist(AddSlash(App.Path) & "CQG.FLG")
        Case eTT_AccountType_CtgCqg
            bReturn = HasModule("B_CTGCQG") Or FileExist(AddSlash(App.Path) & "CtgCqg.FLG")
        Case eTT_AccountType_CtgPats
            bReturn = HasModule("B_CTGPAT") Or FileExist(AddSlash(App.Path) & "CtgPats.FLG")
'        Case eTT_AccountType_CtgPfg
'            bReturn = HasModule("B_CTGPFG") Or FileExist(AddSlash(App.Path) & "CtgPfg.FLG")
        Case eTT_AccountType_Currenex
            bReturn = HasModule("B_CNX") Or FileExist(AddSlash(App.Path) & "Cnx.FLG")
        Case eTT_AccountType_DemoPats
            bReturn = HasModule("B_DEMOPAT") Or FileExist(AddSlash(App.Path) & "DemoPat.FLG")
        Case eTT_AccountType_Etrade
            bReturn = HasModule("B_ETRADE") Or FileExist(AddSlash(App.Path) & "Etrade.FLG")
'        Case eTT_AccountType_FintecPfg
'            bReturn = HasModule("B_FINPFG") Or FileExist(AddSlash(App.Path) & "FinPfg.FLG")
        Case eTT_AccountType_FptCqg
            bReturn = HasModule("B_FPTCQG") Or FileExist(AddSlash(App.Path) & "FptCqg.FLG")
        Case eTT_AccountType_FptOec
            bReturn = HasModule("B_FPTOEC") Or FileExist(AddSlash(App.Path) & "FptOec.FLG")
        Case eTT_AccountType_FxddCurrenex
            bReturn = HasModule("B_FXDDCNX") Or FileExist(AddSlash(App.Path) & "FxddCnx.FLG")
        Case eTT_AccountType_Gft
            bReturn = HasModule("B_GFT") Or FileExist(AddSlash(App.Path) & "Gft.FLG")
        Case eTT_AccountType_Ideal
            bReturn = HasModule("B_IDEAL") Or FileExist(AddSlash(App.Path) & "Ideal.FLG")
        Case eTT_AccountType_IntBrokers
            bReturn = HasModule("B_INTB,INTB") Or FileExist(AddSlash(App.Path) & "IntB.FLG")
        Case eTT_AccountType_KnightCqg
            bReturn = HasModule("B_PENCQG") Or FileExist(AddSlash(App.Path) & "PenCqg.FLG")
        Case eTT_AccountType_KnightCurrenex
            bReturn = HasModule("B_PENCNX") Or FileExist(AddSlash(App.Path) & "PenCnx.FLG")
'        Case eTT_AccountType_LindWaldock
'            bReturn = HasModule("B_LIND") Or FileExist(AddSlash(App.Path) & "Lind.FLG")
'        Case eTT_AccountType_ManExpress
'            bReturn = HasModule("B_MANX") Or FileExist(AddSlash(App.Path) & "ManExpress.FLG")
        Case eTT_AccountType_Oec
            bReturn = HasModule("B_OEC") Or FileExist(AddSlash(App.Path) & "Oec.FLG")
        Case eTT_AccountType_Optimus
            bReturn = HasModule("B_OPTMS") Or FileExist(AddSlash(App.Path) & "Optimus.FLG")
        Case eTT_AccountType_OpVest
            bReturn = HasModule("B_OPVEST") Or FileExist(AddSlash(App.Path) & "OpVest.FLG")
        Case eTT_AccountType_PATS
            bReturn = HasModule("B_PATS") Or FileExist(AddSlash(App.Path) & "PATS.FLG")
'        Case eTT_AccountType_PFG
'            bReturn = HasModule("B_PFG") Or FileExist(AddSlash(App.Path) & "PFG.FLG")
        Case eTT_AccountType_RcgPats
            bReturn = HasModule("B_RCGPATS") Or FileExist(AddSlash(App.Path) & "RcgPats.FLG")
        Case eTT_AccountType_Rithmic
            bReturn = HasModule("B_RITH") Or FileExist(AddSlash(App.Path) & "Rithmic.FLG")
        Case eTT_AccountType_RjoCqg
            bIsBroker = HasModule("B_RJOCQGB") Or FileExist(AddSlash(App.Path) & "RjoCqgB.FLG")
            bReturn = HasModule("B_RJOCQG") Or FileExist(AddSlash(App.Path) & "RjoCqg.FLG") Or bIsBroker
        Case eTT_AccountType_RjoHkPats
            bReturn = HasModule("B_RJOHPAT") Or FileExist(AddSlash(App.Path) & "RjoHkPats.FLG")
        Case eTT_AccountType_RjoPats
            bReturn = HasModule("B_RJOPAT") Or FileExist(AddSlash(App.Path) & "RjoPats.FLG")
        Case eTT_AccountType_RobbinsCqg
            bReturn = HasModule("B_ROBCQG") Or FileExist(AddSlash(App.Path) & "RobCqg.FLG")
        Case eTT_AccountType_SimBroker
            bReturn = True
        Case eTT_AccountType_SimReplay
            bReturn = True
        Case eTT_AccountType_SimStream
            bReturn = True
        Case eTT_AccountType_SimTrade
            bReturn = False
        Case eTT_AccountType_Standard
            bReturn = False
        Case eTT_AccountType_Tradier
            bReturn = HasModule("B_TRADIER") Or FileExist(AddSlash(App.Path) & "Tradier.FLG")
        Case eTT_AccountType_TransAct
            bReturn = HasModule("B_TRAN,TRAN,TRAN2") Or FileExist(AddSlash(App.Path) & "TransAct.FLG")
        Case eTT_AccountType_TT
            bReturn = HasModule("B_TT") Or FileExist(AddSlash(App.Path) & "TT.FLG")
        Case eTT_AccountType_VanKarCurrenex
            bReturn = HasModule("B_VKARCNX") Or FileExist(AddSlash(App.Path) & "VkarCnx.FLG")
        Case eTT_AccountType_Vision
            bReturn = HasModule("B_VISION") Or FileExist(AddSlash(App.Path) & "Vision.FLG")
        Case eTT_AccountType_VisionCqg
            bReturn = HasModule("B_VISCQG") Or FileExist(AddSlash(App.Path) & "VisCqg.FLG")
        Case eTT_AccountType_ZanerCqg
            bReturn = HasModule("B_ZANCQG") Or FileExist(AddSlash(App.Path) & "ZanCqg.FLG")
        Case eTT_AccountType_ZanerCurrenex
            bReturn = HasModule("B_ZANCNX") Or FileExist(AddSlash(App.Path) & "ZanCnx.FLG")
        Case eTT_AccountType_ZanerPats
            bReturn = HasModule("B_ZANPATS") Or FileExist(AddSlash(App.Path) & "ZanPats.FLG")
        Case eTT_AccountType_ZanerRithmic
            bReturn = HasModule("B_ZANRITH") Or FileExist(AddSlash(App.Path) & "ZanRith.FLG")
        Case eTT_AccountType_ZanerZenFire
            bReturn = HasModule("B_ZANZEN") Or FileExist(AddSlash(App.Path) & "ZanZen.FLG")
        Case eTT_AccountType_ZenFire
            bReturn = HasModule("B_ZFIRE") Or FileExist(AddSlash(App.Path) & "ZenFire.FLG")
    End Select
    
    IsBrokerUser = bReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cBrokerDispatch.IsBrokerUser"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsBrokerUserOptions
'' Description: Determine if the user is entitled for using Option Nav with the given broker
'' Inputs:      Broker
'' Returns:     True if entitled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsBrokerUserOptions(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If IsBrokerUser(nBroker) Then
        If IsCqgBroker(nBroker) Then
            bReturn = HasModule("O_CQG") Or FileExist(AddSlash(App.Path) & "CQG.FLG")
        Else
            Select Case nBroker
                Case eTT_AccountType_AdvFut
                    bReturn = False
                Case eTT_AccountType_AlpariCurrenex
                    bReturn = False
                Case eTT_AccountType_AlpariPats
                    bReturn = False
                Case eTT_AccountType_AlpariZenFire
                    bReturn = HasModule("O_ALPZEN") Or FileExist(AddSlash(App.Path) & "AlpZen.FLG")
                Case eTT_AccountType_Ameritrade
                    bReturn = HasModule("O_TDA") Or FileExist(AddSlash(App.Path) & "TDA.FLG")
                Case eTT_AccountType_BornPats
                    bReturn = False
                Case eTT_AccountType_CtgPats
                    bReturn = False
'                Case eTT_AccountType_CtgPfg
'                    bReturn = False
                Case eTT_AccountType_Currenex
                    bReturn = False
                Case eTT_AccountType_DemoPats
                    bReturn = False
                Case eTT_AccountType_Etrade
                    bReturn = False
'                Case eTT_AccountType_FintecPfg
'                    bReturn = False
                Case eTT_AccountType_FptCqg
                    bReturn = HasModule("O_FPTCQG") Or FileExist(AddSlash(App.Path) & "FptCqg.FLG")
                Case eTT_AccountType_FptOec
                    bReturn = HasModule("O_FPTOEC") Or FileExist(AddSlash(App.Path) & "FptOec.FLG")
                Case eTT_AccountType_FxddCurrenex
                    bReturn = False
                Case eTT_AccountType_Gft
                    bReturn = False
                Case eTT_AccountType_Ideal
                    bReturn = HasModule("O_IDEAL") Or FileExist(AddSlash(App.Path) & "Ideal.FLG")
                Case eTT_AccountType_IntBrokers
                    bReturn = HasModule("O_INTB") Or FileExist(AddSlash(App.Path) & "IntB.FLG")
                Case eTT_AccountType_KnightCurrenex
                    bReturn = False
'                Case eTT_AccountType_LindWaldock
'                    bReturn = False
'                Case eTT_AccountType_ManExpress
'                    bReturn = False
                Case eTT_AccountType_Oec
                    bReturn = HasModule("O_OEC") Or FileExist(AddSlash(App.Path) & "Oec.FLG")
                Case eTT_AccountType_Optimus
                    bReturn = HasModule("O_OPTMS") Or FileExist(AddSlash(App.Path) & "Optimus.FLG")
                Case eTT_AccountType_OpVest
                    bReturn = HasModule("O_OPVEST") Or FileExist(AddSlash(App.Path) & "OpVest.FLG")
                Case eTT_AccountType_PATS
                    bReturn = False
'                Case eTT_AccountType_PFG
'                    bReturn = HasModule("O_PFG") Or FileExist(AddSlash(App.Path) & "PFG.FLG")
                Case eTT_AccountType_RcgPats
                    bReturn = False
                Case eTT_AccountType_Rithmic
                    bReturn = HasModule("O_RITH") Or FileExist(AddSlash(App.Path) & "Rithmic.FLG")
                Case eTT_AccountType_RjoHkPats
                    bReturn = False
                Case eTT_AccountType_RjoPats
                    bReturn = False
                Case eTT_AccountType_SimBroker
                    bReturn = True
                Case eTT_AccountType_SimReplay
                    bReturn = True
                Case eTT_AccountType_SimStream
                    bReturn = True
                Case eTT_AccountType_SimTrade
                    bReturn = False
                Case eTT_AccountType_Standard
                    bReturn = False
                Case eTT_AccountType_Tradier
                    bReturn = HasModule("O_TRADIER") Or FileExist(AddSlash(App.Path) & "Tradier.FLG")
                Case eTT_AccountType_TransAct
                    bReturn = False
                Case eTT_AccountType_TT
                    bReturn = False
                Case eTT_AccountType_VanKarCurrenex
                    bReturn = False
                Case eTT_AccountType_Vision
                    bReturn = HasModule("O_VISION") Or FileExist(AddSlash(App.Path) & "Vision.FLG")
                Case eTT_AccountType_ZanerCurrenex
                    bReturn = False
                Case eTT_AccountType_ZanerPats
                    bReturn = False
                Case eTT_AccountType_ZanerRithmic
                    bReturn = HasModule("O_ZANRITH") Or FileExist(AddSlash(App.Path) & "ZanRith.FLG")
                Case eTT_AccountType_ZanerZenFire
                    bReturn = HasModule("O_ZANZEN") Or FileExist(AddSlash(App.Path) & "ZanZen.FLG")
                Case eTT_AccountType_ZenFire
                    bReturn = HasModule("O_ZFIRE") Or FileExist(AddSlash(App.Path) & "ZenFire.FLG")
            End Select
        End If
    End If
    
    IsBrokerUserOptions = bReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cBrokerDispatch.IsBrokerUserOptions"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsBrokerSimUser
'' Description: Determine if the user is entitled for the given broker and is
''              also a simulated user
'' Inputs:      Broker
'' Returns:     True if entitled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsBrokerSimUser(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    IsBrokerSimUser = IsBrokerUser(nBroker) And HasModule("BRKRDEMO")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsBrokerSimUser"

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
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim BrokerObj As cBroker            ' Broker object

    Select Case nBroker
        Case eTT_AccountType_AdvFut
            strReturn = "Advantage Futures"
'        Case eTT_AccountType_CtgPfg
'            strReturn = "CTG3"
'        Case eTT_AccountType_FintecPfg
'            strReturn = "Fintec"
'        Case eTT_AccountType_LindWaldock
'            strReturn = "Lind Waldock"
'        Case eTT_AccountType_ManExpress
'            strReturn = "Man Express"
        Case eTT_AccountType_PATS
            strReturn = "PATS"
'        Case eTT_AccountType_PFG
'            strReturn = "PFG"
        Case eTT_AccountType_TransAct
            strReturn = "TransAct"
        Case eTT_AccountType_TT
            strReturn = "Trading Technologies"
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                strReturn = BrokerObj.BrokerName
            End If
    
    End Select
    
    BrokerName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.BrokerName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerForControlName
'' Description: Determine the broker for the given control name
'' Inputs:      Control Name
'' Returns:     Broker
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerForControlName(ByVal strControlName As String) As eTT_AccountType
On Error GoTo ErrSection:

    Dim nReturn As eTT_AccountType      ' Return value for the function
    
    nReturn = -1&
    Select Case UCase(strControlName)
        Case "ADVANTAGEFUTURES"
            nReturn = eTT_AccountType_AdvFut
        Case "ALPARICNX"
            nReturn = eTT_AccountType_AlpariCurrenex
        Case "ALPARIPATS"
            nReturn = eTT_AccountType_AlpariPats
        Case "ALPARIZENFIRE"
            nReturn = eTT_AccountType_AlpariZenFire
        Case "AMERITRADE"
            nReturn = eTT_AccountType_Ameritrade
        Case "AMPCQG"
            nReturn = eTT_AccountType_AmpCqg
        Case "BORNPATS"
            nReturn = eTT_AccountType_BornPats
        Case "CQG"
            nReturn = eTT_AccountType_CQG
        Case "CTGCQG"
            nReturn = eTT_AccountType_CtgCqg
        Case "CTGPATS"
            nReturn = eTT_AccountType_CtgPats
'        Case "CTGPFG"
'            nReturn = eTT_AccountType_CtgPfg
        Case "CURRENEX"
            nReturn = eTT_AccountType_Currenex
        Case "DEMOPATS"
            nReturn = eTT_AccountType_DemoPats
        Case "ETRADE"
            nReturn = eTT_AccountType_Etrade
'        Case "FINTECPFG"
'            nReturn = eTT_AccountType_FintecPfg
        Case "FPTCQG"
            nReturn = eTT_AccountType_FptCqg
        Case "FPTOEC"
            nReturn = eTT_AccountType_FptOec
        Case "FXDDCNX"
            nReturn = eTT_AccountType_FxddCurrenex
        Case "GFT"
            nReturn = eTT_AccountType_Gft
        Case "IDEAL"
            nReturn = eTT_AccountType_Ideal
        Case "INTERACTIVEBROKERS"
            nReturn = eTT_AccountType_IntBrokers
        Case "KNIGHTCNX"
            nReturn = eTT_AccountType_KnightCurrenex
        Case "KNIGHTCQG"
            nReturn = eTT_AccountType_KnightCqg
'        Case "LINDWALDOCK"
'            nReturn = eTT_AccountType_LindWaldock
'        Case "MANEXPRESS"
'            nReturn = eTT_AccountType_ManExpress
        Case "OEC"
            nReturn = eTT_AccountType_Oec
        Case "OPTIMUS"
            nReturn = eTT_AccountType_Optimus
        Case "OPVEST"
            nReturn = eTT_AccountType_OpVest
        Case "PATS"
            nReturn = eTT_AccountType_PATS
'        Case "PFG"
'            nReturn = eTT_AccountType_PFG
        Case "RCGPATS"
            nReturn = eTT_AccountType_RcgPats
        Case "RITHMIC"
            nReturn = eTT_AccountType_Rithmic
        Case "RJOCQG"
            nReturn = eTT_AccountType_RjoCqg
        Case "RJOHKPATS"
            nReturn = eTT_AccountType_RjoHkPats
        Case "RJOPATS"
            nReturn = eTT_AccountType_RjoPats
        Case "ROBBINSCQG"
            nReturn = eTT_AccountType_RobbinsCqg
        Case "TRADIER"
            nReturn = eTT_AccountType_Tradier
        Case "TRANSACT"
            nReturn = eTT_AccountType_TransAct
        Case "TRADINGTECHNOLOGIES"
            nReturn = eTT_AccountType_TT
        Case "VANKARCNX"
            nReturn = eTT_AccountType_VanKarCurrenex
        Case "VISION"
            nReturn = eTT_AccountType_Vision
        Case "VISIONCQG"
            nReturn = eTT_AccountType_VisionCqg
        Case "ZANERCNX"
            nReturn = eTT_AccountType_ZanerCurrenex
        Case "ZANERCQG"
            nReturn = eTT_AccountType_ZanerCqg
        Case "ZANERPATS"
            nReturn = eTT_AccountType_ZanerPats
        Case "ZANERRITHMIC"
            nReturn = eTT_AccountType_ZanerRithmic
        Case "ZANERZENFIRE"
            nReturn = eTT_AccountType_ZanerZenFire
        Case "ZENFIRE"
            nReturn = eTT_AccountType_ZenFire
    End Select
    
    BrokerForControlName = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.BrokerForControlName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActivityLogSource
'' Description: Return the activity log source name of the given brokerage
'' Inputs:      Broker
'' Returns:     Activity Log Source Name of Brokerage
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ActivityLogSource(ByVal nBroker As eTT_AccountType) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case nBroker
        Case eTT_AccountType_AdvFut
            strReturn = "AdvFut"
        Case eTT_AccountType_IntBrokers
            strReturn = "IB"
        Case eTT_AccountType_SimBroker
            strReturn = "SimBroker"
        Case eTT_AccountType_TT
            strReturn = "TT"
        Case Else
            strReturn = BrokerName(nBroker)
    End Select
    
    ActivityLogSource = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ActivityLogSource"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Connect
'' Description: Make an attempt to connect to the appropriate broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Connect(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker object
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                g.CtgPfg.Connect
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                g.FintecPfg.Connect
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                g.LindWaldock.Connect
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                g.ManExpress.Connect
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                g.PFG.Connect
'            End If
    
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                g.Transact.Connect
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                BrokerObj.Connect
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.Connect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanDisconnect
'' Description: Can the user manually disconnect from the given broker?
'' Inputs:      Broker
'' Returns:     True if can disconnect, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CanDisconnect(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim strReturn As String             ' User's response to question
    Dim bReturn As Boolean              ' Return value from the function
    Dim lTimeOut As Long                ' Timeout variable
    Dim bNotAllHeldAtBroker As Boolean  ' Are all the OCO's held not at the broker?
    
    bReturn = True

    If ConnectionStatusForBroker(nBroker) = eGDConnectionStatus_Connected Then
        If g.OrderStrategies.HasWorkingExitsForBroker(nBroker, bNotAllHeldAtBroker) Then
            If bNotAllHeldAtBroker = True Then
                strReturn = InfBox("Disconnecting from the " & BrokerName(nBroker) & " servers will result in having your auto exit orders cancelled.||Do you want to continue?|", "?", "+Yes|-No", BrokerName(nBroker) & " Disconnect Warning")
                If strReturn = "Y" Then bReturn = True Else bReturn = False
            End If
        End If
        
        If (bReturn = True) And (HasWorkingOcoOrders(nBroker) = True) Then
            strReturn = InfBox("Disconnecting from the " & BrokerName(nBroker) & " servers will result in having your Order-Cancel-Order orders cancelled.||Do you want to continue?|", "?", "+Yes|-No", BrokerName(nBroker) & " Disconnect Warning")
            If strReturn = "Y" Then bReturn = True Else bReturn = False
        End If
        
        If (bReturn = True) And (g.TsoGroups.HasWorkingGroups(nBroker) = True) Then
            bReturn = frmActiveTsConfirmation.ShowMeForExit("Manual disconnect from broker", nBroker)
        End If
        
        If bReturn = True Then
            CancelWorkingOcoOrders nBroker
            g.OrderStrategies.DeactivateExitsForBroker nBroker, "Disconnecting from " & BrokerName(nBroker), True
            
            lTimeOut = 0&
            'Do While ((g.OrderStrategies.HasWorkingExitsForBroker(nBroker) = True) Or (HasWorkingOcoOrders(nBroker) = True)) And (lTimeOut < 30&)
            Do While ((g.OrderStrategies.HasDeactivatingExitsForBroker(nBroker) = True) Or (HasWorkingOcoOrders(nBroker) = True)) And (lTimeOut < 30&)
                Sleep 1&
                lTimeOut = lTimeOut + 1&
            Loop
        End If
    Else
        If g.OrderStrategies.HasWorkingExitsForBroker(nBroker, bNotAllHeldAtBroker) Then
            If bNotAllHeldAtBroker = True Then
                InfBox "You are not currently connected to the " & BrokerName(nBroker) & " servers, but you still have active auto exit orders.  These orders are still live at the broker, but Order Cancel Order, Break-Even Stops, or Trailing Stops are no longer in effect.|", "!", , BrokerName(nBroker) & " Disconnect Warning"
            End If
        End If
    End If
    
    CanDisconnect = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CanDisconnect"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disconnect
'' Description: Make an attempt to disconnect from the appropriate broker
'' Inputs:      Broker, Reason
'' Returns:     True if going to disconnect, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Disconnect(ByVal nBroker As eTT_AccountType, ByVal strReason As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim BrokerObj As cBroker            ' Broker object

    If CanDisconnect(nBroker) Then
        Select Case nBroker
'            Case eTT_AccountType_CtgPfg
'                If Not g.CtgPfg Is Nothing Then
'                    g.CtgPfg.Disconnect False, strReason
'                End If
'
'            Case eTT_AccountType_FintecPfg
'                If Not g.FintecPfg Is Nothing Then
'                    g.FintecPfg.Disconnect False, strReason
'                End If
'
'            Case eTT_AccountType_LindWaldock
'                If Not g.LindWaldock Is Nothing Then
'                    g.LindWaldock.Disconnect False, strReason
'                End If
'
'            Case eTT_AccountType_ManExpress
'                If Not g.ManExpress Is Nothing Then
'                    g.ManExpress.Disconnect False, strReason
'                End If
'
'            Case eTT_AccountType_PFG
'                If Not g.PFG Is Nothing Then
'                    g.PFG.Disconnect False, strReason
'                End If
        
            Case eTT_AccountType_TransAct
                If Not g.Transact Is Nothing Then
                    g.Transact.Disconnect "", False, strReason
                End If
                
            Case Else
                Set BrokerObj = Broker(nBroker)
                If Not BrokerObj Is Nothing Then
                    BrokerObj.Disconnect False, strReason
                End If
        
        End Select
        
        bReturn = True
    End If
    
    Disconnect = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.Disconnect"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisconnectFromAll
'' Description: Make an attempt to disconnect from all live brokers
'' Inputs:      Reason, Closing Trade Navigator?
'' Returns:     True if going to disconnect, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DisconnectFromAll(ByVal strReason As String, ByVal bClosingTn As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim lIndex As Long                  ' Index into a for loop

    bReturn = CanStopStreaming(strReason, bClosingTn)
    If bReturn Then
        For lIndex = 1 To kNumBrokers - 1
            If IsBrokerUser(lIndex) Then
                bReturn = Disconnect(lIndex, strReason)
                If bReturn = False Then
                    Exit For
                End If
            End If
        Next lIndex
    End If

    DisconnectFromAll = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.DisconnectFromAll"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanStopStreaming
'' Description: Determine if streaming can be stopped
'' Inputs:      Reason, Closing Trade Nav?
'' Returns:     True if can stop streaming, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CanStopStreaming(ByVal strReason As String, ByVal bClosingTn As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim strAction As String             ' Action being attempted
    Dim strReturn As String             ' Return value from an InfBox

    bReturn = True
    If bClosingTn Then
        strAction = "close Trade Navigator"
    Else
        strAction = "shut down streaming"
    End If

    If g.nReplaySession <> 0 Then
        InfBox "You cannot " & strAction & " until you disable stream replay", "!", , "Stream Replay Error"
        bReturn = False
    End If

    If bReturn And g.TradingItems.HasActiveAutoTradeItems Then
        strReturn = InfBox("You cannot " & strAction & " until you disable all auto trading items", "!", "Disable All|+-Cancel", "Streaming Error")
        bReturn = False
        
        If strReturn = "D" Then
            g.TradingItems.DisableTradeItems "Disable All button", False, True, False, False, False
            bReturn = Not g.TradingItems.HasActiveAutoTradeItems
        End If
    End If
    
    If bReturn And g.TsoGroups.HasWorkingGroups Then
        bReturn = frmActiveTsConfirmation.ShowMeForExit(strReason)
    End If
    
    CanStopStreaming = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CanStopStreaming"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerInfo
'' Description: Get the broker information structure for the given broker
'' Inputs:      Broker
'' Returns:     Broker Info
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerInfo(ByVal nBroker As eTT_AccountType) As cBrokerInfo
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker object
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If g.CtgPfg Is Nothing Then
'                Set BrokerInfo = Nothing
'            Else
'                Set BrokerInfo = g.CtgPfg.BrokerInfo
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If g.FintecPfg Is Nothing Then
'                Set BrokerInfo = Nothing
'            Else
'                Set BrokerInfo = g.FintecPfg.BrokerInfo
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If g.LindWaldock Is Nothing Then
'                Set BrokerInfo = Nothing
'            Else
'                Set BrokerInfo = g.LindWaldock.BrokerInfo
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If g.ManExpress Is Nothing Then
'                Set BrokerInfo = Nothing
'            Else
'                Set BrokerInfo = g.ManExpress.BrokerInfo
'            End If
'
'        Case eTT_AccountType_PFG
'            If g.PFG Is Nothing Then
'                Set BrokerInfo = Nothing
'            Else
'                Set BrokerInfo = g.PFG.BrokerInfo
'            End If
        
        Case eTT_AccountType_TransAct
            If g.Transact Is Nothing Then
                Set BrokerInfo = Nothing
            Else
                Set BrokerInfo = g.Transact.BrokerInfo
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                Set BrokerInfo = BrokerObj.BrokerInfo
            End If
    
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.BrokerInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Refreshing
'' Description: Is the broker in the process of refreshing anything?
'' Inputs:      Broker
'' Returns:     True if refreshing, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Refreshing(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim BrokerObj As cBroker            ' Broker object
        
    bReturn = False
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                bReturn = g.CtgPfg.Refreshing
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                bReturn = g.FintecPfg.Refreshing
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                bReturn = g.LindWaldock.Refreshing
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                bReturn = g.ManExpress.Refreshing
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                bReturn = g.PFG.Refreshing
'            End If
        
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                bReturn = g.Transact.Refreshing
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                bReturn = BrokerObj.Refreshing
            End If
    
    End Select
    
    Refreshing = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.Refreshing"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LastChanged
'' Description: Get the last time any broker information has changed for a broker
'' Inputs:      Broker
'' Returns:     Last Changed tick count
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastChanged(ByVal nBroker As eTT_AccountType) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        dReturn = BInfo.LastChanged
    Else
        dReturn = -2#
    End If
    
    LastChanged = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.LastChanged"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LastChangedForAll
'' Description: Get the last time broker information has changed for all brokers
'' Inputs:      None
'' Returns:     Array of Last Changed info
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastChangedForAll() As cGdArray
On Error GoTo ErrSection:

    Dim adReturn As cGdArray            ' Array to return from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    Set adReturn = New cGdArray
    adReturn.Create eGDARRAY_Doubles, kNumBrokers
    
    For lIndex = 1 To kNumBrokers - 1
        Set BInfo = BrokerInfo(lIndex)
        
        If BInfo Is Nothing Then
            adReturn(lIndex) = -2#
        ElseIf BInfo.Refreshing Then
            adReturn(lIndex) = -1#
        Else
            adReturn(lIndex) = BInfo.LastChanged
        End If
    Next lIndex
    
    Set LastChangedForAll = adReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.LastChangedForAll"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolLastChanged
'' Description: Get the last time any symbol information has changed for an account
'' Inputs:      Account ID, Symbol
'' Returns:     Last Changed tick count
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get SymbolLastChanged(ByVal vAccountIDorNumber As Variant, ByVal vSymbolOrSymbolID As Variant) As Double
On Error GoTo ErrSection:

    Dim strAccountNumber As String      ' Account number
    Dim strSymbol As String             ' Symbol
    Dim dReturn As Double               ' Return value for the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    strAccountNumber = GetAccountNumber(vAccountIDorNumber)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    
    Set BInfo = BrokerInfo(AccountTypeForNumber(strAccountNumber))
    If Not BInfo Is Nothing Then
        dReturn = BInfo.SymbolLastChanged(strAccountNumber, strSymbol)
    Else
        dReturn = -2
    End If
    
    SymbolLastChanged = dReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerDispatch.SymbolLastChanged.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolLastChanged
'' Description: Get the last time any symbol information has changed for an account
'' Inputs:      Account ID, Symbol
'' Returns:     Last Changed tick count
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let SymbolLastChanged(ByVal vAccountIDorNumber As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal dLastChanged As Double)
On Error GoTo ErrSection:

    Dim strAccountNumber As String      ' Account number
    Dim strSymbol As String             ' Symbol
    Dim nBroker As eTT_AccountType      ' Broker for the account passed in
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    nBroker = AccountTypeForID(GetAccountID(vAccountIDorNumber))
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        strAccountNumber = GetAccountNumber(vAccountIDorNumber)
        strSymbol = GetSymbol(vSymbolOrSymbolID)
        
        BInfo.SymbolLastChanged(strAccountNumber, strSymbol) = dLastChanged
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerDispatch.SymbolLastChanged.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectionStatusForAccount
'' Description: Determine the connection status for the given account
'' Inputs:      Account Number, Check Broker View?
'' Returns:     Connection Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConnectionStatusForAccount(ByVal vAccountNumberOrID As Variant, Optional ByVal bCheckBrokerView As Boolean = False) As eGDConnectionStatus
On Error GoTo ErrSection:

    Dim strAccountNumber As String      ' Account number for account passed in
    Dim nBroker As eTT_AccountType      ' Broker for the account passed in
    Dim nReturn As eGDConnectionStatus  ' Return value from the function
    Dim BrokerObj As cBroker            ' Broker object
    
    strAccountNumber = GetAccountNumber(vAccountNumberOrID)
    nBroker = AccountTypeForNumber(strAccountNumber)

    nReturn = eGDConnectionStatus_Disconnected
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                nReturn = g.CtgPfg.ConnectionStatusForAccount(strAccountNumber)
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                nReturn = g.FintecPfg.ConnectionStatusForAccount(strAccountNumber)
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                nReturn = g.LindWaldock.ConnectionStatusForAccount(strAccountNumber)
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                nReturn = g.ManExpress.ConnectionStatusForAccount(strAccountNumber)
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                nReturn = g.PFG.ConnectionStatusForAccount(strAccountNumber)
'            End If
        
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                nReturn = g.Transact.ConnectionStatusForAccount(strAccountNumber)
            End If
        
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                nReturn = BrokerObj.ConnectionStatusForAccount(strAccountNumber)
            End If
    
    End Select
    
    If bCheckBrokerView Then
        If nReturn = eGDConnectionStatus_Disconnected Then
            If FormIsLoaded("frmBrokerView") Then
                If frmBrokerView.Broker = nBroker Then
                    nReturn = frmBrokerView.ConnectionStatus
                End If
            End If
        End If
    End If
    
    ConnectionStatusForAccount = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ConnectionStatusForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectionStatusForBroker
'' Description: Get the connection status for the given broker
'' Inputs:      Broker
'' Returns:     Connection Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConnectionStatusForBroker(ByVal nBroker As eTT_AccountType) As eGDConnectionStatus
On Error GoTo ErrSection:

    Dim nReturn As eGDConnectionStatus  ' Return value for the function
    Dim BrokerObj As cBroker            ' Broker object

    nReturn = eGDConnectionStatus_Disconnected
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                nReturn = g.CtgPfg.ConnectionStatus
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                nReturn = g.FintecPfg.ConnectionStatus
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                nReturn = g.LindWaldock.ConnectionStatus
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                nReturn = g.ManExpress.ConnectionStatus
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                nReturn = g.PFG.ConnectionStatus
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                nReturn = g.Transact.ConnectionStatus
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                nReturn = BrokerObj.ConnectionStatus
            End If
    
    End Select
    
    ConnectionStatusForBroker = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ConnectionStatusForBroker"
    
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

    If Not g.ActivityLogs Is Nothing Then
        g.ActivityLogs.AddStatusToActivityLog nBroker, nNewStatus
    End If
    
    If Not g.TradingItems Is Nothing Then
        g.TradingItems.ConnectionStatusChanged nBroker, nNewStatus
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ConnectionStatusChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Refresh
'' Description: Send a refresh request to the given broker
'' Inputs:      Broker, Verbose Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Refresh(ByVal nBroker As eTT_AccountType, Optional ByVal bVerbose As Boolean = False)
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker object
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                g.CtgPfg.Refresh bVerbose
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                g.FintecPfg.Refresh bVerbose
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                g.LindWaldock.Refresh bVerbose
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                g.ManExpress.Refresh bVerbose
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                g.PFG.Refresh bVerbose
'            End If
    
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                g.Transact.RefreshAccount True, False, bVerbose
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                BrokerObj.Refresh bVerbose
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.Refresh"
    
End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrderFromBrokerID
'' Description: Attempt to Cancel the order given the broker order ID
'' Inputs:      Account, Broker ID, Genesis ID
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelOrderFromBrokerID(ByVal vAccountIDorNumber As Variant, ByVal strBrokerID As String, Optional ByVal strGenesisID As String = "")
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim BrokerObj As cBroker            ' Broker object
    
    strAccount = GetAccountNumber(vAccountIDorNumber)
    nBroker = AccountTypeForNumber(strAccount)
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                If g.CtgPfg.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    g.CtgPfg.CancelOrderFromBrokerID strBrokerID, strGenesisID
'                Else
'                    InfBox "Cannot cancel order because you are not currently connected to the " & g.CtgPfg.BrokerName & " servers", "!", , "Cancel Order Error"
'                End If
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                If g.FintecPfg.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    g.FintecPfg.CancelOrderFromBrokerID strBrokerID, strGenesisID
'                Else
'                    InfBox "Cannot cancel order because you are not currently connected to the " & g.FintecPfg.BrokerName & " servers", "!", , "Cancel Order Error"
'                End If
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                If g.LindWaldock.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    g.LindWaldock.CancelOrderFromBrokerID strBrokerID, strGenesisID
'                Else
'                    InfBox "Cannot cancel order because you are not currently connected to the Lind Waldock servers", "!", , "Cancel Order Error"
'                End If
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                If g.ManExpress.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    g.ManExpress.CancelOrderFromBrokerID strBrokerID, strGenesisID
'                Else
'                    InfBox "Cannot cancel order because you are not currently connected to the Man Express servers", "!", , "Cancel Order Error"
'                End If
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                If g.PFG.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    g.PFG.CancelOrderFromBrokerID strBrokerID, strGenesisID
'                Else
'                    InfBox "Cannot cancel order because you are not currently connected to the PFG servers", "!", , "Cancel Order Error"
'                End If
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                If g.Transact.ConnectionStatus = eGDConnectionStatus_Connected Then
                    g.Transact.CancelOrderFromBrokerID strBrokerID, strGenesisID
                Else
                    InfBox "Cannot cancel order because you are not currently connected to the TransAct servers", "!", , "Cancel Order Error"
                End If
            End If
        
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                BrokerObj.CancelOrderFromBrokerID strBrokerID, strGenesisID
            Else
                InfBox "Cannot cancel order because you are not currently connected to the " & BrokerName(nBroker) & " servers", "!", , "Cancel Order Error"
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.CancelOrderFromBrokerID"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelWorkingOrders
'' Description: Cancel working orders for the given account and symbol
'' Inputs:      Account, Symbol, Auto Trade Item ID, Skip Order ID, Direction
'' Returns:     Number of orders cancelled
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelWorkingOrders(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant, ByVal lAtID As Long, Optional ByVal lSkipOrderID As Long = -1&, Optional ByVal nDirection As eGDFilterDirection = eGDFilterDirection_All) As Long
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol
    Dim lSymbolID As Long               ' Symbol ID
    Dim lReturn As Long                 ' Return value for the function
    Dim BrokerObj As cBroker            ' Broker object
    
    lReturn = 0&
    strAccount = GetAccountNumber(vAccountIDorNumber)
    nBroker = AccountTypeForNumber(strAccount)
    strSymbol = GetSymbol(ConvertToTradeSymbol(vSymbolIDOrSymbol, Int(CurrentTime("", "", True))))
    lSymbolID = GetSymbolID(strSymbol)
    If lSymbolID = 0 Then vSymbolIDOrSymbol = strSymbol Else vSymbolIDOrSymbol = lSymbolID
    
    g.OrderStrategies.CancelAllOrders GetAccountID(vAccountIDorNumber), vSymbolIDOrSymbol
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                If g.CtgPfg.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    lReturn = g.CtgPfg.CancelWorkingOrders(strAccount, strSymbol, lAtID, lSkipOrderID, nDirection)
'                Else
'                    InfBox "Cannot cancel orders because you are not currently connected to the " & g.CtgPfg.BrokerName & " servers", "!", , "Cancel Order Error"
'                End If
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                If g.FintecPfg.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    lReturn = g.FintecPfg.CancelWorkingOrders(strAccount, strSymbol, lAtID, lSkipOrderID, nDirection)
'                Else
'                    InfBox "Cannot cancel orders because you are not currently connected to the " & g.FintecPfg.BrokerName & " servers", "!", , "Cancel Order Error"
'                End If
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                If g.LindWaldock.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    lReturn = g.LindWaldock.CancelWorkingOrders(strAccount, strSymbol, lAtID, lSkipOrderID, nDirection)
'                Else
'                    InfBox "Cannot cancel orders because you are not currently connected to the Lind Waldock servers", "!", , "Cancel Order Error"
'                End If
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                If g.ManExpress.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    lReturn = g.ManExpress.CancelWorkingOrders(strAccount, strSymbol, lAtID, lSkipOrderID, nDirection)
'                Else
'                    InfBox "Cannot cancel orders because you are not currently connected to the Man Express servers", "!", , "Cancel Order Error"
'                End If
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                If g.PFG.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    lReturn = g.PFG.CancelWorkingOrders(strAccount, strSymbol, lAtID, lSkipOrderID, nDirection)
'                Else
'                    InfBox "Cannot cancel orders because you are not currently connected to the PFG servers", "!", , "Cancel Order Error"
'                End If
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                If g.Transact.ConnectionStatus = eGDConnectionStatus_Connected Then
                    lReturn = g.Transact.CancelWorkingOrders(strAccount, strSymbol, lAtID, lSkipOrderID, nDirection)
                Else
                    InfBox "Cannot cancel orders because you are not currently connected to the TransAct servers", "!", , "Cancel Order Error"
                End If
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                lReturn = BrokerObj.CancelWorkingOrders(strAccount, strSymbol, lAtID, lSkipOrderID, nDirection)
            Else
                InfBox "Cannot cancel orders because you are not currently connected to the " & BrokerName(nBroker) & " servers", "!", , "Cancel Order Error"
            End If
    
    End Select
    
    CancelWorkingOrders = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CancelWorkingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenPosition
'' Description: Flatten position for the given account and symbol
'' Inputs:      Account, Symbol, Auto Trade Item ID, Confirm?, Genesis ID
'' Returns:     Did User Confirm?
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FlattenPosition(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant, ByVal lAtID As Long, Optional ByVal bConfirm As Boolean = True, Optional strGenesisOrderID As String) As Boolean
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol
    Dim bReturn As Boolean              ' Return value for the function
    Dim BrokerObj As cBroker            ' Broker object
    
    bReturn = False
    strAccount = GetAccountNumber(vAccountIDorNumber)
    nBroker = AccountTypeForNumber(strAccount)
    strSymbol = GetSymbol(vSymbolIDOrSymbol)
    
    ' For the time being, I really don't like the confirmation in this case (DAJ 07/19/2007)...
    bConfirm = False
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                If g.CtgPfg.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    ' The folks at PFG do not want us to confirm the market order with the user, so
'                    ' the bConfirm is set to False for now (2/6/2007) DAJ...
'                    bReturn = g.CtgPfg.FlattenPosition(strAccount, strSymbol, lAtID, False, strGenesisOrderID)
'                Else
'                    InfBox "Cannot flatten position because you are not currently connected to the " & g.CtgPfg.BrokerName & " servers", "!", , "Flatten Position Error"
'                End If
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                If g.FintecPfg.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    ' The folks at PFG do not want us to confirm the market order with the user, so
'                    ' the bConfirm is set to False for now (2/6/2007) DAJ...
'                    bReturn = g.FintecPfg.FlattenPosition(strAccount, strSymbol, lAtID, False, strGenesisOrderID)
'                Else
'                    InfBox "Cannot flatten position because you are not currently connected to the " & g.FintecPfg.BrokerName & " servers", "!", , "Flatten Position Error"
'                End If
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                If g.LindWaldock.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    bReturn = g.LindWaldock.FlattenPosition(strAccount, strSymbol, lAtID, bConfirm, strGenesisOrderID)
'                Else
'                    InfBox "Cannot flatten position because you are not currently connected to the Lind Waldock servers", "!", , "Flatten Position Error"
'                End If
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                If g.ManExpress.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    bReturn = g.ManExpress.FlattenPosition(strAccount, strSymbol, lAtID, bConfirm, strGenesisOrderID)
'                Else
'                    InfBox "Cannot flatten position because you are not currently connected to the Man Express servers", "!", , "Flatten Position Error"
'                End If
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                If g.PFG.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    ' The folks at PFG do not want us to confirm the market order with the user, so
'                    ' the bConfirm is set to False for now (2/6/2007) DAJ...
'                    bReturn = g.PFG.FlattenPosition(strAccount, strSymbol, lAtID, False, strGenesisOrderID)
'                Else
'                    InfBox "Cannot flatten position because you are not currently connected to the PFG servers", "!", , "Flatten Position Error"
'                End If
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                If g.Transact.ConnectionStatus = eGDConnectionStatus_Connected Then
                    bReturn = g.Transact.FlattenPosition(strAccount, strSymbol, lAtID, bConfirm, strGenesisOrderID)
                Else
                    InfBox "Cannot flatten position because you are not currently connected to the TransAct servers", "!", , "Flatten Position Error"
                End If
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                bReturn = BrokerObj.FlattenPosition(strAccount, strSymbol, lAtID, bConfirm, strGenesisOrderID)
            Else
                InfBox "Cannot flatten position because you are not currently connected to the " & BrokerName(nBroker) & " servers", "!", , "Flatten Position Error"
            End If
    
    End Select
    
    FlattenPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.FlattenPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReversePosition
'' Description: Reverse position for the given account and symbol
'' Inputs:      Account, Symbol, Auto Trade Item ID, Confirm?, Genesis ID
'' Returns:     Did User Confirm?
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReversePosition(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant, ByVal lAtID As Long, Optional ByVal bConfirm As Boolean = True, Optional strGenesisOrderID As String) As Boolean
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol
    Dim bReturn As Boolean              ' Return value for the function
    Dim BrokerObj As cBroker            ' Broker object
    
    bReturn = False
    strAccount = GetAccountNumber(vAccountIDorNumber)
    nBroker = AccountTypeForNumber(strAccount)
    strSymbol = GetSymbol(vSymbolIDOrSymbol)
    
    ' For the time being, I really don't like the confirmation in this case (DAJ 07/19/2007)...
    bConfirm = False
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                If g.CtgPfg.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    ' The folks at PFG do not want us to confirm the market order with the user, so
'                    ' the bConfirm is set to False for now (2/6/2007) DAJ...
'                    bReturn = g.CtgPfg.ReversePosition(strAccount, strSymbol, lAtID, False, strGenesisOrderID)
'                Else
'                    InfBox "Cannot reverse position because you are not currently connected to the " & g.CtgPfg.BrokerName & " servers", "!", , "Reverse Position Error"
'                End If
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                If g.FintecPfg.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    ' The folks at PFG do not want us to confirm the market order with the user, so
'                    ' the bConfirm is set to False for now (2/6/2007) DAJ...
'                    bReturn = g.FintecPfg.ReversePosition(strAccount, strSymbol, lAtID, False, strGenesisOrderID)
'                Else
'                    InfBox "Cannot reverse position because you are not currently connected to the " & g.FintecPfg.BrokerName & " servers", "!", , "Reverse Position Error"
'                End If
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                If g.LindWaldock.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    bReturn = g.LindWaldock.ReversePosition(strAccount, strSymbol, lAtID, bConfirm, strGenesisOrderID)
'                Else
'                    InfBox "Cannot reverse position because you are not currently connected to the Lind Waldock servers", "!", , "Reverse Position Error"
'                End If
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                If g.ManExpress.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    bReturn = g.ManExpress.ReversePosition(strAccount, strSymbol, lAtID, bConfirm, strGenesisOrderID)
'                Else
'                    InfBox "Cannot reverse position because you are not currently connected to the Man Express servers", "!", , "Reverse Position Error"
'                End If
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                If g.PFG.ConnectionStatus = eGDConnectionStatus_Connected Then
'                    ' The folks at PFG do not want us to confirm the market order with the user, so
'                    ' the bConfirm is set to False for now (2/6/2007) DAJ...
'                    bReturn = g.PFG.ReversePosition(strAccount, strSymbol, lAtID, False, strGenesisOrderID)
'                Else
'                    InfBox "Cannot reverse position because you are not currently connected to the PFG servers", "!", , "Reverse Position Error"
'                End If
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                If g.Transact.ConnectionStatus = eGDConnectionStatus_Connected Then
                    bReturn = g.Transact.ReversePosition(strAccount, strSymbol, lAtID, bConfirm, strGenesisOrderID)
                Else
                    InfBox "Cannot reverse position because you are not currently connected to the TransAct servers", "!", , "Reverse Position Error"
                End If
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                bReturn = BrokerObj.ReversePosition(strAccount, strSymbol, lAtID, bConfirm, strGenesisOrderID)
            Else
                InfBox "Cannot reverse position because you are not currently connected to the " & BrokerName(nBroker) & " servers", "!", , "Reverse Position Error"
            End If
    
    End Select
    
    ReversePosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ReversePosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionString
'' Description: Get the position string for charts and ladders
'' Inputs:      Account, Symbol, Auto Trade Item ID
'' Returns:     Position String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionString(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long) As String
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol passed in
    Dim strReturn As String             ' Return value from the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    strAccount = GetAccountNumber(vAccountNumberOrID)
    nBroker = AccountTypeForNumber(strAccount)
    strSymbol = GetSymbol(ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True))))
    strReturn = ""
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        strReturn = BInfo.PositionString(strAccount, strSymbol, lAtID)
    End If
    
    PositionString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.PositionString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSummary
'' Description: Get the fill summary for given symbol, account, and at id
'' Inputs:      Account, Symbol, Auto Trade Item ID
'' Returns:     Fill Summary
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillSummary(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long) As cAccountPosition
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol passed in
    Dim AcctPos As cAccountPosition     ' Fill summary to return from the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
        
    strAccount = GetAccountNumber(vAccountNumberOrID)
    nBroker = AccountTypeForNumber(strAccount)
    strSymbol = GetSymbol(ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True))))
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set AcctPos = BInfo.FillSummaryForSymbol(strAccount, strSymbol, lAtID)
    End If
    
    Set FillSummary = AcctPos

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.FillSummary"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSummaryForBroker
'' Description: Get the all of the fill summaries for given broker
'' Inputs:      Broker
'' Returns:     Fill Summaries
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillSummaryForBroker(ByVal nBroker As eTT_AccountType) As cAccountPositions
On Error GoTo ErrSection:

    Dim FillSumms As cAccountPositions  ' Account positions for the broker
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set FillSumms = BInfo.FillSummary.MakeCopy
    End If
    
    Set FillSummaryForBroker = FillSumms

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.FillSummaryForBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSummaryForTradeItem
'' Description: Get the fill summary for the given automated trading item
'' Inputs:      Trade Item
'' Returns:     Fill Summary
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillSummaryForTradeItem(ByVal TradeItem As cAutoTradeItem) As cAccountPosition
On Error GoTo ErrSection:

    Dim AcctPos As cAccountPosition     ' Fill summary to return from the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
        
    Set BInfo = BrokerInfo(TradeItem.Broker)
    If Not BInfo Is Nothing Then
        Set AcctPos = BInfo.FillSummaryForID(TradeItem.AccountPositionID)
    End If
    
    Set FillSummaryForTradeItem = AcctPos

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.FillSummaryForTradeItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    WorkingOrders
'' Description: Get the working orders for charts and ladders
'' Inputs:      Account, Symbol, Auto Trade Item ID
'' Returns:     Array of working orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WorkingOrders(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long) As cPtOrders
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol passed in
    Dim Orders As cPtOrders             ' Collection of orders to return
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    strAccount = GetAccountNumber(vAccountNumberOrID)
    nBroker = AccountTypeForNumber(strAccount)
    strSymbol = GetSymbol(ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True))))
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Orders = BInfo.OrdersForSymbol(strAccount, strSymbol, lAtID, True)
    End If
    
    Set WorkingOrders = Orders
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.WorkingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyPositions
'' Description: Verify positions with the user
'' Inputs:      Broker, Account, Symbol, Force the verify?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub VerifyPositions(ByVal nBroker As eTT_AccountType, Optional ByVal strAccount As String = "", Optional ByVal strSymbol As String = "", Optional ByVal bForce As Boolean = True)
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker object
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                g.CtgPfg.VerifyPositions strAccount, strSymbol, bForce
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                g.FintecPfg.VerifyPositions strAccount, strSymbol, bForce
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                g.LindWaldock.VerifyPositions strAccount, strSymbol, bForce
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                g.ManExpress.VerifyPositions strAccount, strSymbol, bForce
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                g.PFG.VerifyPositions strAccount, strSymbol, bForce
'            End If
        
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                g.Transact.VerifyPositions strAccount, strSymbol, bForce
            End If
        
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                BrokerObj.VerifyPositions strAccount, strSymbol, bForce
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.VerifyPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyBrokerPositions
'' Description: Verify all broker positions with the user if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub VerifyBrokerPositions()
On Error GoTo ErrSection:

    Dim bVerified As Boolean            ' Has the broker been verified?
    Dim lIndex As Long                  ' Index into a for loop
    Dim bPositionsToVerify As Boolean   ' Are there still positions to verify?
    Static bSimTradeShown As Boolean    ' Has the SimTrade message been shown at least once?

    bVerified = False
    
    If (g.bStarting = False) Then
        If (g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected) Or (g.RealTime.ConnectionStatus = eGDConnectionStatus_Disconnected) Then
            For lIndex = 1 To kNumBrokers - 1
                If IsBrokerUser(lIndex) Then
                    If (ConnectionStatusForBroker(lIndex) = eGDConnectionStatus_Connected) Then
                        If (PositionVerify(lIndex) = True) And (Refreshing(lIndex) = False) Then
                            If (lIndex <> eTT_AccountType_SimStream) Or ((g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected) Or (bSimTradeShown = False)) Then
                                If lIndex = eTT_AccountType_SimStream Then
                                    bSimTradeShown = True
                                End If
                                VerifyPositions lIndex
                                bVerified = True
                                
                                If PositionVerify(lIndex) = False Then
                                    frmActiveTsConfirmation.ShowMeForEntry lIndex
                                End If
                                
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next lIndex
        End If
    End If
                
    If bVerified Then
        UpdateVisibleCharts eRedo3_Settings
    
        bPositionsToVerify = False
        For lIndex = 1 To kNumBrokers - 1
            If IsBrokerUser(lIndex) Then
                If (ConnectionStatusForBroker(lIndex) = eGDConnectionStatus_Connected) Then
                    If PositionVerify(lIndex) Then
                        bPositionsToVerify = True
                        Exit For
                    End If
                End If
            End If
        Next lIndex
        
        PositionsToVerify = bPositionsToVerify
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.VerifyBrokerPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionVerify
'' Description: Do we still need to verify positions for the given broker?
'' Inputs:      Broker
'' Returns:     True if need to verify positions, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get PositionVerify(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker object
    Dim bReturn As Boolean              ' Return value from the function

    bReturn = False
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                bReturn = g.CtgPfg.PositionVerify
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                bReturn = g.FintecPfg.PositionVerify
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                bReturn = g.LindWaldock.PositionVerify
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                bReturn = g.ManExpress.PositionVerify
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                bReturn = g.PFG.PositionVerify
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                bReturn = g.Transact.PositionVerify
            End If
        
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                bReturn = BrokerObj.PositionVerify
            End If
    
    End Select
    
    PositionVerify = bReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerDispatch.PositionVerify.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionVerify
'' Description: Do we still need to verify positions for the given broker?
'' Inputs:      Broker
'' Returns:     True if need to verify positions, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let PositionVerify(ByVal nBroker As eTT_AccountType, ByVal bPositionVerify As Boolean)
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker object
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                g.CtgPfg.PositionVerify = bPositionVerify
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                g.FintecPfg.PositionVerify = bPositionVerify
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                g.LindWaldock.PositionVerify = bPositionVerify
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                g.ManExpress.PositionVerify = bPositionVerify
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                g.PFG.PositionVerify = bPositionVerify
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                ''g.Transact.PositionVerify = bPositionVerify
            End If
        
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                BrokerObj.PositionVerify = bPositionVerify
            End If
    
    End Select

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerDispatch.PositionVerify.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NotifyNewExchanges
'' Description: Notify the user about any new exchanges
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub NotifyNewExchanges()
On Error GoTo ErrSection:

    If g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
        If Not g.Ideal Is Nothing Then
            g.Ideal.NotifyNewExchanges
        End If
        
        If Not g.IntBroker Is Nothing Then
            g.IntBroker.NotifyNewExchanges
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cBrokerDispatch.NotifyNewExchanges"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAutoTradeItems
'' Description: Tell the appropriate broker to update their auto trade items
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshAutoTradeItems(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.RefreshAutoTradeItems
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.RefreshAutoTradeItems"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAutoExits
'' Description: Tell the appropriate broker to update their auto exits
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshAutoExits(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.RefreshAutoExits
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.RefreshAutoExits"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TextPosition
'' Description: Convert the given position into text
'' Inputs:      Position
'' Returns:     Text Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TextPosition(ByVal lPosition As Long) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from the function

    Select Case lPosition
        Case Is > 0
            strReturn = "Long " & Str(lPosition)
        Case Is < 0
            strReturn = "Short " & Str(Abs(lPosition))
        Case 0
            strReturn = "Flat"
    End Select
    
    TextPosition = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.TextPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HideAccount
'' Description: Should this account be hidden?
'' Inputs:      Account
'' Returns:     True if Hide, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HideAccount(ByVal vAccountNumberOrID As Variant) As Boolean
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim bReturn As Boolean
    
    nBroker = AccountTypeForID(GetAccountID(vAccountNumberOrID))
    bReturn = False
    
    Select Case nBroker
        Case eTT_AccountType_SimReplay
            bReturn = (g.nReplaySession = 0)
        
        Case eTT_AccountType_TransAct
            bReturn = (g.nReplaySession <> 0) Or (Not IsBrokerUser(eTT_AccountType_TransAct)) Or TransActSimulatedAccount(GetAccountNumber(vAccountNumberOrID))
            
        Case Else
            bReturn = (g.nReplaySession <> 0) Or (Not IsBrokerUser(nBroker))
        
    End Select
    
    HideAccount = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.HideAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HideBroker
'' Description: Should this broker be hidden?
'' Inputs:      Broker, Account, Allow TransAct Simulated?
'' Returns:     True if Hide, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HideBroker(ByVal nBroker As eTT_AccountType, Optional ByVal vAccountNumberOrID As Variant = "", Optional ByVal bAllowTransActSimulated As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean
    
    bReturn = False
    Select Case nBroker
        Case eTT_AccountType_SimReplay
            bReturn = (g.nReplaySession = 0)
        
        Case eTT_AccountType_TransAct
            If bAllowTransActSimulated Then
                bReturn = (g.nReplaySession <> 0) Or (Not IsBrokerUser(eTT_AccountType_TransAct))
            Else
                bReturn = (g.nReplaySession <> 0) Or (Not IsBrokerUser(eTT_AccountType_TransAct)) Or (TransActSimulatedAccount(GetAccountNumber(vAccountNumberOrID)))
            End If
            
        Case Else
            bReturn = (g.nReplaySession <> 0) Or (Not IsBrokerUser(nBroker))
        
    End Select
    
    HideBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.HideBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsStreamReplayAccount
'' Description: Is the given account a stream replay account?
'' Inputs:      Account
'' Returns:     True if Stream Replay Account, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsStreamReplayAccount(ByVal vAccountNumberOrID As Variant) As Boolean
On Error GoTo ErrSection:

    Dim strAccountNumber As String      ' Account number for account passed in
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    strAccountNumber = GetAccountNumber(vAccountNumberOrID)
    
    If Len(strAccountNumber) = 10 Then
        If UCase(Left(strAccountNumber, 5)) = "GENSR" Then
            bReturn = True
        End If
    End If
    
    IsStreamReplayAccount = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsStreamReplayAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Get orders for the specified broker and account
'' Inputs:      Broker, Account, Broker Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetOrders(ByVal nBroker As eTT_AccountType, Optional ByVal vAccountNumberOrID As Variant = "", Optional ByVal strBrokerOrderID As String = "")
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker base class object

    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                If Len(strBrokerOrderID) > 0 Then
'                    g.CtgPfg.GetSingleOrder strBrokerOrderID
'                Else
'                    g.CtgPfg.GetOrders GetAccountNumber(vAccountNumberOrID)
'                End If
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                If Len(strBrokerOrderID) > 0 Then
'                    g.FintecPfg.GetSingleOrder strBrokerOrderID
'                Else
'                    g.FintecPfg.GetOrders GetAccountNumber(vAccountNumberOrID)
'                End If
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                g.LindWaldock.GetOrders GetAccountNumber(vAccountNumberOrID)
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                g.ManExpress.GetOrders GetAccountNumber(vAccountNumberOrID)
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                If Len(strBrokerOrderID) > 0 Then
'                    g.PFG.GetSingleOrder strBrokerOrderID
'                Else
'                    g.PFG.GetOrders GetAccountNumber(vAccountNumberOrID)
'                End If
'            End If
        
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                g.Transact.GetOrders
            End If
        
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                BrokerObj.GetOrders GetAccountNumber(vAccountNumberOrID)
            End If
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.GetOrders"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositions
'' Description: Get positions for the specified broker and account
'' Inputs:      Broker, Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPositions(ByVal nBroker As eTT_AccountType, Optional ByVal vAccountNumberOrID As Variant = "")
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker base class object
    Dim strAccountNumber As String      ' Account number for the account passed in

    strAccountNumber = GetAccountNumber(vAccountNumberOrID)

    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                g.CtgPfg.GetPositions strAccountNumber
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                g.FintecPfg.GetPositions strAccountNumber
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                g.LindWaldock.GetPositions strAccountNumber
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                g.ManExpress.GetPositions strAccountNumber
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                g.PFG.GetPositions strAccountNumber
'            End If
        
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                g.Transact.GetPositions
            End If
        
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                BrokerObj.GetPositions GetAccountNumber(vAccountNumberOrID)
            End If
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.GetPositions"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowBrokerConnectionInfo
'' Description: Allow the user to view their broker connection information
'' Inputs:      Broker, New?, User Name, Connect?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowBrokerConnectionInfo(ByVal nBroker As eTT_AccountType, Optional ByVal bNew As Boolean = False, Optional ByVal strUserName As String = "", Optional ByVal bConnect As Boolean = True)
On Error GoTo ErrSection:

    Dim nOldLoginMode As eGDTransActLoginModes ' TransAct login mode
    Dim nNewLoginMode As eGDTransActLoginModes ' TransAct login mode
    Dim strOldUserName As String            ' Previous User Name
    Dim BrokerObj As cBroker            ' Broker object
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If frmPfgLogin.ShowMe(nBroker, strUserName, True) Then
'                If bConnect Then g.CtgPfg.Connect
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If frmPfgLogin.ShowMe(nBroker, strUserName, True) Then
'                If bConnect Then g.FintecPfg.Connect
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If frmLwLogin.ShowMe(nBroker, strUserName, True) Then
'                If bConnect Then g.LindWaldock.Connect
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If frmLwLogin.ShowMe(nBroker, strUserName, True) Then
'                If bConnect Then g.ManExpress.Connect
'            End If
'
'        Case eTT_AccountType_PFG
'            If frmPfgLogin.ShowMe(nBroker, strUserName, True) Then
'                If bConnect Then g.PFG.Connect
'            End If
            
        Case eTT_AccountType_TransAct
            strOldUserName = g.Transact.UserName
            nOldLoginMode = g.Transact.LoginMode
            
            If frmTransActLogin.ShowMe(strUserName, "", True, False, bNew) Then
                If bConnect Then
                    ' If the user is logged into the TransAct servers with the same user name, but
                    ' the login mode is now different, we need to switch login modes...
                    If (g.Transact.UserName = strOldUserName) And (g.Transact.LoginMode <> nOldLoginMode) And (g.Transact.ConnectionStatus = eGDConnectionStatus_Connected) Then
                        nNewLoginMode = g.Transact.LoginMode
                        g.Transact.LoginMode = nOldLoginMode
                        g.Transact.SwitchLoginModes nNewLoginMode
                        
                    ' Otherwise, we are either connecting or switching user names, so we can call
                    ' connect to do that...
                    Else
                        g.Transact.Connect
                    End If
                End If
            End If
        
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                BrokerObj.ShowBrokerConnectionInfo bNew, strUserName, bConnect
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ShowBrokerConnectionInfo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SwitchBrokerAccount
'' Description: Allow the user to switch their online broker account
'' Inputs:      Broker, New Login, New Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SwitchBrokerAccount(ByVal nBroker As eTT_AccountType, Optional ByVal strLogin As String = "", Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker object
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                g.CtgPfg.Connect strLogin, True
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                g.FintecPfg.Connect strLogin, True
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                g.LindWaldock.Connect strLogin, True
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                g.ManExpress.Connect strLogin, True
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                g.PFG.Connect strLogin, True
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                g.Transact.Connect strLogin, strAccount, True
            End If
    
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                BrokerObj.Connect strLogin, True
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.SwitchBrokerAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowActivityView
'' Description: Allow the user to view the activity view
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowActivityView(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim frm As frmBrokerInfo            ' Broker Information form

    Set frm = BrokerInfoForm(nBroker)
    If frm Is Nothing Then
        Set frm = New frmBrokerInfo
        frm.ShowMe nBroker
    Else
        frm.SetFocus
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ShowActivityView"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowBrokerView
'' Description: Allow the user to view the broker view
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowBrokerView(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim frm As frmBrokerInfo            ' Broker Information form

    Select Case nBroker
        Case eTT_AccountType_CtgPats
            frmBrokerSnapshot.ShowMe nBroker
            
        Case eTT_AccountType_DemoPats
            frmBrokerSnapshot.ShowMe nBroker
            
        Case eTT_AccountType_RjoCqg
            frmBrokerView.ShowMe nBroker
            
        Case eTT_AccountType_RjoPats
            frmBrokerSnapshot.ShowMe nBroker
            
        Case eTT_AccountType_TT
            If g.TT.ConnectToPrice Then
                frmBrokerSnapshot.ShowMe nBroker
            End If
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ShowBrokerView"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBrokerCombo
'' Description: Load a broker combo with the appropriate brokers
'' Inputs:      Combo Box
'' Returns:     True if non-simulated brokers, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadBrokerCombo(cboBroker As ctlUniComboImageXP) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim lIndex As Long                  ' Index into a for loop

    With cboBroker
#If 1 Then
        For lIndex = 0 To kNumBrokers - 1
            If (IsBrokerUser(lIndex) = True) And (lIndex <> eTT_AccountType_SimReplay) Then
                .AddItem BrokerName(lIndex)
                .ItemData(.NewIndex) = lIndex
                bReturn = True
            End If
        Next lIndex
#Else
        If IsBrokerUser(eTT_AccountType_AdvFut) Then
            .AddItem "Advantage Futures"
            .ItemData(.NewIndex) = eTT_AccountType_AdvFut
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_Ameritrade) Then
            .AddItem "TD Ameritrade"
            .ItemData(.NewIndex) = eTT_AccountType_Ameritrade
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_AmpCqg) Then
            .AddItem "Amp"
            .ItemData(.NewIndex) = eTT_AccountType_AmpCqg
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_BornPats) Then
            .AddItem "Born Capital"
            .ItemData(.NewIndex) = eTT_AccountType_BornPats
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_CQG) Then
            .AddItem "CQG"
            .ItemData(.NewIndex) = eTT_AccountType_CQG
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_CtgCqg) Then
            .AddItem "CTG2"
            .ItemData(.NewIndex) = eTT_AccountType_CtgCqg
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_CtgPats) Then
            .AddItem "CTG1"
            .ItemData(.NewIndex) = eTT_AccountType_CtgPats
            bReturn = True
        End If
'        If IsBrokerUser(eTT_AccountType_CtgPfg) Then
'            .AddItem "CTG3"
'            .ItemData(.NewIndex) = eTT_AccountType_CtgPfg
'            bReturn = True
'        End If
        If IsBrokerUser(eTT_AccountType_DemoPats) Then
            .AddItem "PATS Demo"
            .ItemData(.NewIndex) = eTT_AccountType_DemoPats
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_Etrade) Then
            .AddItem "E-Trade"
            .ItemData(.NewIndex) = eTT_AccountType_Etrade
            bReturn = True
        End If
'        If IsBrokerUser(eTT_AccountType_FintecPfg) Then
'            .AddItem "Fintec"
'            .ItemData(.NewIndex) = eTT_AccountType_FintecPfg
'            bReturn = True
'        End If
        If IsBrokerUser(eTT_AccountType_FptCqg) Then
            .AddItem "FptCqg"
            .ItemData(.NewIndex) = eTT_AccountType_FptCqg
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_FptOec) Then
            .AddItem "FptOec"
            .ItemData(.NewIndex) = eTT_AccountType_FptOec
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_Gft) Then
            .AddItem "GFT Forex"
            .ItemData(.NewIndex) = eTT_AccountType_Gft
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_Ideal) Then
            .AddItem "I-Deal"
            .ItemData(.NewIndex) = eTT_AccountType_Ideal
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_IntBrokers) Then
            .AddItem "Interactive Brokers"
            .ItemData(.NewIndex) = eTT_AccountType_IntBrokers
            bReturn = True
        End If
'        If IsBrokerUser(eTT_AccountType_LindWaldock) Then
'            .AddItem BrokerName(eTT_AccountType_LindWaldock)
'            .ItemData(.NewIndex) = eTT_AccountType_LindWaldock
'            bReturn = True
'        End If
'        If IsBrokerUser(eTT_AccountType_ManExpress) Then
'            .AddItem "Man Express"
'            .ItemData(.NewIndex) = eTT_AccountType_ManExpress
'            bReturn = True
'        End If
        If IsBrokerUser(eTT_AccountType_Oec) Then
            .AddItem "Oec"
            .ItemData(.NewIndex) = eTT_AccountType_Oec
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_Optimus) Then
            .AddItem "Optimus"
            .ItemData(.NewIndex) = eTT_AccountType_Optimus
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_OpVest) Then
            .AddItem "OpVest"
            .ItemData(.NewIndex) = eTT_AccountType_OpVest
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_PATS) Then
            .AddItem "PATS"
            .ItemData(.NewIndex) = eTT_AccountType_PATS
            bReturn = True
        End If
'        If IsBrokerUser(eTT_AccountType_PFG) Then
'            .AddItem "PFG"
'            .ItemData(.NewIndex) = eTT_AccountType_PFG
'            bReturn = True
'        End If
        If IsBrokerUser(eTT_AccountType_Rithmic) Then
            .AddItem "Rithmic"
            .ItemData(.NewIndex) = eTT_AccountType_Rithmic
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_RjoCqg) Then
            .AddItem "R.J. O'Brien"
            .ItemData(.NewIndex) = eTT_AccountType_RjoCqg
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_RjoHkPats) Then
            .AddItem "R.J. O'Brien (Hong Kong)"
            .ItemData(.NewIndex) = eTT_AccountType_RjoHkPats
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_RjoPats) Then
            .AddItem "R.J. O'Brien"
            .ItemData(.NewIndex) = eTT_AccountType_RjoPats
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_Etrade) Then
            .AddItem "E-Trade"
            .ItemData(.NewIndex) = eTT_AccountType_Etrade
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_Tradier) Then
            .AddItem "Tradier"
            .ItemData(.NewIndex) = eTT_AccountType_Tradier
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_SimBroker) Then
            .AddItem BrokerName(eTT_AccountType_SimBroker)
            .ItemData(.NewIndex) = eTT_AccountType_SimBroker
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_SimStream) Then
            .AddItem "SimTrade"
            .ItemData(.NewIndex) = eTT_AccountType_SimStream
        End If
        If IsBrokerUser(eTT_AccountType_TT) Then
            .AddItem "Trading Technologies"
            .ItemData(.NewIndex) = eTT_AccountType_TT
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_Vision) Then
            .AddItem "Vision"
            .ItemData(.NewIndex) = eTT_AccountType_Vision
            bReturn = True
        End If
        If IsBrokerUser(eTT_AccountType_ZenFire) Then
            .AddItem "Zen-Fire"
            .ItemData(.NewIndex) = eTT_AccountType_ZenFire
            bReturn = True
        End If
#End If
    End With
    
    LoadBrokerCombo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.LoadBrokerCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsLiveAccount
'' Description: Determine if this is a simulated account or live broker account
'' Inputs:      Account Type
'' Returns:     True if Live Account, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsLiveAccount(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    IsLiveAccount = Not ((nBroker = eTT_AccountType_SimStream) Or (nBroker = eTT_AccountType_SimReplay) Or (nBroker = eTT_AccountType_SimBroker))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsLiveAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TooltipForAccount
'' Description: Determine the tooltip for the given account
'' Inputs:      Account
'' Returns:     Tooltip
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TooltipForAccount(ByVal vAccountNumberOrID As Variant) As String
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Account type for the given account
    Dim strBrokerName As String         ' Broker name
    Dim strReturn As String             ' Return value from the function
    
    nBroker = AccountTypeForID(GetAccountID(vAccountNumberOrID))
    strBrokerName = BrokerName(nBroker)
    
    Select Case ConnectionStatusForAccount(vAccountNumberOrID)
        Case eGDConnectionStatus_Connected
            If nBroker = eTT_AccountType_TransAct Then
                If g.Transact Is Nothing Then
                    strReturn = strBrokerName & ": Connected (Click here to Disconnect)"
                Else
                    strReturn = strBrokerName & ": Connected in " & TransActLoginModeString(g.Transact.LoginMode) & " mode (Click here to Disconnect)"
                End If
            Else
                strReturn = strBrokerName & ": Connected (Click here to Disconnect)"
            End If
        
        Case eGDConnectionStatus_Disconnected
            strReturn = strBrokerName & ": Disconnected (Click here to Connect)"
        
        Case eGDConnectionStatus_Connecting
            If nBroker = eTT_AccountType_TransAct Then
                If g.Transact Is Nothing Then
                    strReturn = strBrokerName & ": Connecting"
                Else
                    strReturn = strBrokerName & ": Connecting in " & TransActLoginModeString(g.Transact.LoginMode) & " mode"
                End If
            Else
                strReturn = strBrokerName & ": Connecting"
            End If
        
        Case eGDConnectionStatus_Disconnecting
            strReturn = strBrokerName & ": Disconnecting"
    
    End Select

    TooltipForAccount = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.TooltipForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToggleConnectionForAccount
'' Description: Attempt to toggle the connection for the given account
'' Inputs:      Account, User Name, Reason for Disconnect
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ToggleConnectionForAccount(ByVal vAccountNumberOrID As Variant, ByVal strUserName As String, Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim BrokerObj As cBroker            ' Broker object
    
    nBroker = AccountTypeForID(GetAccountID(vAccountNumberOrID))
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                Select Case ConnectionStatusForAccount(vAccountNumberOrID)
'                    Case eGDConnectionStatus_Connected
'                        If CanDisconnect(eTT_AccountType_CtgPfg) Then
'                            g.CtgPfg.Disconnect False, strReason
'                        End If
'
'                    Case eGDConnectionStatus_Disconnected
'                        g.CtgPfg.Connect strUserName
'                End Select
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                Select Case ConnectionStatusForAccount(vAccountNumberOrID)
'                    Case eGDConnectionStatus_Connected
'                        If CanDisconnect(eTT_AccountType_FintecPfg) Then
'                            g.FintecPfg.Disconnect False, strReason
'                        End If
'
'                    Case eGDConnectionStatus_Disconnected
'                        g.FintecPfg.Connect strUserName
'                End Select
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                Select Case ConnectionStatusForAccount(vAccountNumberOrID)
'                    Case eGDConnectionStatus_Connected
'                        If CanDisconnect(eTT_AccountType_LindWaldock) Then
'                            g.LindWaldock.Disconnect False, strReason
'                        End If
'
'                    Case eGDConnectionStatus_Disconnected
'                        g.LindWaldock.Connect
'                End Select
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                Select Case ConnectionStatusForAccount(vAccountNumberOrID)
'                    Case eGDConnectionStatus_Connected
'                        If CanDisconnect(eTT_AccountType_ManExpress) Then
'                            g.ManExpress.Disconnect False, strReason
'                        End If
'
'                    Case eGDConnectionStatus_Disconnected
'                        g.ManExpress.Connect
'                End Select
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                Select Case ConnectionStatusForAccount(vAccountNumberOrID)
'                    Case eGDConnectionStatus_Connected
'                        If CanDisconnect(eTT_AccountType_PFG) Then
'                            g.PFG.Disconnect False, strReason
'                        End If
'
'                    Case eGDConnectionStatus_Disconnected
'                        g.PFG.Connect strUserName
'                End Select
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                Select Case ConnectionStatusForAccount(vAccountNumberOrID)
                    Case eGDConnectionStatus_Connected
                        If CanDisconnect(eTT_AccountType_TransAct) Then
                            g.Transact.Disconnect GetAccountNumber(vAccountNumberOrID), False, strReason
                        End If
                        
                    Case eGDConnectionStatus_Disconnected
                        g.Transact.Connect strUserName, GetAccountNumber(vAccountNumberOrID)
                
                End Select
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                Select Case ConnectionStatusForAccount(vAccountNumberOrID)
                    Case eGDConnectionStatus_Connected
                        If CanDisconnect(BrokerObj.Broker) Then
                            BrokerObj.Disconnect False, strReason
                        End If
                        
                    Case eGDConnectionStatus_Disconnected
                        BrokerObj.Connect
                        
                End Select
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ToggleConnectionForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TempBrokerAccounts
'' Description: Tell the Trade Console which temporary broker accounts to create
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TempBrokerAccounts()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To kNumBrokers - 1
        If IsLiveAccount(lIndex) Then
            TempBrokerAccount lIndex, IsBrokerUser(lIndex)
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.TempBrokerAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckPendingOrders
'' Description: Check for any pending orders that may be getting stale
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckPendingOrders()
On Error GoTo ErrSection:
    
    Dim astrOrders As cGdArray          ' Array of pending order id's
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim bHasSent As Boolean             ' Are there orders stuck in "Sent" status?
    Dim BInfo As cBrokerInfo            ' Broker info object
    Dim strBrokerName As String         ' Broker name
    
    For lIndex = 1 To kNumBrokers - 1
        Set BInfo = BrokerInfo(lIndex)
        If Not BInfo Is Nothing Then
            If ConnectionStatusForBroker(lIndex) = eGDConnectionStatus_Connected Then
                Set astrOrders = New cGdArray
                astrOrders.Create eGDARRAY_Strings
                
                If BInfo.HasOldPendingOrders(astrOrders, bHasSent) Then
                    BInfo.PendingOrders.Clear
                
                    strBrokerName = BrokerName(lIndex)
                    InfBox "One or more orders have been sent to the " & strBrokerName & " servers but were not confirmed in a timely manner.  This lack of confirmation may be due to various internet connections or conditions.||PLEASE CALL YOUR BROKER AND VERIFY YOUR CURRENT ACCOUNT STATUS.||We are forcing a reconnect to the " & strBrokerName & " servers now to see if this will confirm pending orders.  You will be notified if the orders were not confirmed.|", "!", "+OK", "Unconfirmed " & strBrokerName & " orders", True
                    
'                    If lIndex = eTT_AccountType_CtgPfg Then
'                        BrokerDebug lIndex, "Calling for a refresh since unconfirmed orders were found"
'                        If bHasSent Then
'                            Refresh lIndex
'                        Else
'                            ' Ask for single order refreshes on each of the pending orders...
'                            For lIndex2 = 0 To astrOrders.Size - 1
'                                g.CtgPfg.GetSingleOrder astrOrders(lIndex2)
'                            Next lIndex2
'                        End If
'                    ElseIf lIndex = eTT_AccountType_FintecPfg Then
'                        BrokerDebug lIndex, "Calling for a refresh since unconfirmed orders were found"
'                        If bHasSent Then
'                            Refresh lIndex
'                        Else
'                            ' Ask for single order refreshes on each of the pending orders...
'                            For lIndex2 = 0 To astrOrders.Size - 1
'                                g.FintecPfg.GetSingleOrder astrOrders(lIndex2)
'                            Next lIndex2
'                        End If
'                    ElseIf lIndex = eTT_AccountType_PFG Then
'                        BrokerDebug lIndex, "Calling for a refresh since unconfirmed orders were found"
'                        If bHasSent Then
'                            Refresh lIndex
'                        Else
'                            ' Ask for single order refreshes on each of the pending orders...
'                            For lIndex2 = 0 To astrOrders.Size - 1
'                                g.PFG.GetSingleOrder astrOrders(lIndex2)
'                            Next lIndex2
'                        End If
                    If lIndex = eTT_AccountType_TransAct Then
                        g.Transact.Disconnect "", True, "Pending Orders"
                    Else
                        BrokerDebug lIndex, "Calling for a refresh since unconfirmed orders were found"
                        Refresh lIndex
                    End If
                End If
            End If
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.CheckPendingOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitBrokerObjects
'' Description: Initialize the broker objects as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitBrokerObjects()
On Error GoTo ErrSection:

    Dim bIsBroker As Boolean            ' Is this user a broker?

    If IsBrokerUser(eTT_AccountType_AdvFut) Then
        If g.AdvFut Is Nothing Then
            Set g.AdvFut = New cBrokerTt
            g.AdvFut.Init eTT_AccountType_AdvFut
        End If
    End If

    If IsBrokerUser(eTT_AccountType_AlpariCurrenex) Then
        If g.AlpariCnx Is Nothing Then
            Set g.AlpariCnx = New cBrokerCurrenex
            g.AlpariCnx.Init eTT_AccountType_AlpariCurrenex
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_AlpariPats) Then
        If g.AlpariPats Is Nothing Then
            Set g.AlpariPats = New cBrokerPats
            g.AlpariPats.Init eTT_AccountType_AlpariPats
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_AlpariZenFire) Then
        If g.AlpariZenFire Is Nothing Then
            Set g.AlpariZenFire = New cRithmic
            g.AlpariZenFire.Init eTT_AccountType_AlpariZenFire
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_Ameritrade) Then
        If g.Ameritrade Is Nothing Then
            Set g.Ameritrade = New cBrokerAmeritrade
            g.Ameritrade.Init eTT_AccountType_Ameritrade
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_AmpCqg) Then
        If g.AmpCqg Is Nothing Then
            Set g.AmpCqg = New cBrokerCqg
            g.AmpCqg.Init eTT_AccountType_AmpCqg
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_BornPats) Then
        If g.BornPats Is Nothing Then
            Set g.BornPats = New cBrokerPats
            g.BornPats.Init eTT_AccountType_BornPats
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_CQG) Then
        If g.CQG Is Nothing Then
            Set g.CQG = New cBrokerCqg
            g.CQG.Init eTT_AccountType_CQG
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_CtgCqg) Then
        If g.CtgCqg Is Nothing Then
            Set g.CtgCqg = New cBrokerCqg
            g.CtgCqg.Init eTT_AccountType_CtgCqg
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_CtgPats) Then
        If g.CtgPats Is Nothing Then
            Set g.CtgPats = New cBrokerPats
            g.CtgPats.Init eTT_AccountType_CtgPats
        End If
    End If
    
'    If IsBrokerUser(eTT_AccountType_CtgPfg) Then
'        If g.CtgPfg Is Nothing Then
'            Set g.CtgPfg = New cPFG
'            g.CtgPfg.Init eTT_AccountType_CtgPfg
'        End If
'    End If
    
    If IsBrokerUser(eTT_AccountType_Currenex) Then
        If g.Currenex Is Nothing Then
            Set g.Currenex = New cBrokerCurrenex
            g.Currenex.Init eTT_AccountType_Currenex
        End If
    End If

    If IsBrokerUser(eTT_AccountType_DemoPats) Then
        If g.DemoPats Is Nothing Then
            Set g.DemoPats = New cBrokerPats
            g.DemoPats.Init eTT_AccountType_DemoPats
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_Etrade) Then
        If g.Etrade Is Nothing Then
            Set g.Etrade = New cBrokerEtrade
            g.Etrade.Init eTT_AccountType_Etrade
        End If
    End If
    
'    If IsBrokerUser(eTT_AccountType_FintecPfg) Then
'        If g.FintecPfg Is Nothing Then
'            Set g.FintecPfg = New cPFG
'            g.FintecPfg.Init eTT_AccountType_FintecPfg
'        End If
'    End If
    
    If IsBrokerUser(eTT_AccountType_FptCqg) Then
        If g.FptCqg Is Nothing Then
            Set g.FptCqg = New cBrokerCqg
            g.FptCqg.Init eTT_AccountType_FptCqg
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_FptOec) Then
        If g.FptOec Is Nothing Then
            Set g.FptOec = New cBrokerOec
            g.FptOec.Init eTT_AccountType_FptOec
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_FxddCurrenex) Then
        If g.FxddCnx Is Nothing Then
            Set g.FxddCnx = New cBrokerCurrenex
            g.FxddCnx.Init eTT_AccountType_FxddCurrenex
        End If
    End If

    If IsBrokerUser(eTT_AccountType_Gft) Then
        If g.Gft Is Nothing Then
            Set g.Gft = New cBrokerGft
            g.Gft.Init eTT_AccountType_Gft
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_Ideal) Then
        If g.Ideal Is Nothing Then
            Set g.Ideal = New cIntBrokers
            g.Ideal.Init eTT_AccountType_Ideal
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_IntBrokers) Then
        If g.IntBroker Is Nothing Then
            Set g.IntBroker = New cIntBrokers
            g.IntBroker.Init eTT_AccountType_IntBrokers
        End If
    End If

    If IsBrokerUser(eTT_AccountType_KnightCqg) Then
        If g.KnightCqg Is Nothing Then
            Set g.KnightCqg = New cBrokerCqg
            g.KnightCqg.Init eTT_AccountType_KnightCqg
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_KnightCurrenex) Then
        If g.KnightCnx Is Nothing Then
            Set g.KnightCnx = New cBrokerCurrenex
            g.KnightCnx.Init eTT_AccountType_KnightCurrenex
        End If
    End If

'    If IsBrokerUser(eTT_AccountType_LindWaldock) Then
'        If g.LindWaldock Is Nothing Then
'            Set g.LindWaldock = New cXpress
'            g.LindWaldock.Init eTT_AccountType_LindWaldock
'        End If
'    End If
'
'    If IsBrokerUser(eTT_AccountType_ManExpress) Then
'        If g.ManExpress Is Nothing Then
'            Set g.ManExpress = New cXpress
'            g.ManExpress.Init eTT_AccountType_ManExpress
'        End If
'    End If

    If IsBrokerUser(eTT_AccountType_Oec) Then
        If g.Oec Is Nothing Then
            Set g.Oec = New cBrokerOec
            g.Oec.Init eTT_AccountType_Oec
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_Optimus) Then
        If g.Optimus Is Nothing Then
            Set g.Optimus = New cRithmic
            g.Optimus.Init eTT_AccountType_Optimus
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_OpVest) Then
        If g.OpVest Is Nothing Then
            Set g.OpVest = New cRithmic
            g.OpVest.Init eTT_AccountType_OpVest
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_PATS) Then
        If g.Pats Is Nothing Then
            Set g.Pats = New cBrokerPats
            g.Pats.Init eTT_AccountType_PATS
        End If
    End If
    
'    If IsBrokerUser(eTT_AccountType_PFG) Then
'        If g.PFG Is Nothing Then
'            Set g.PFG = New cPFG
'            g.PFG.Init eTT_AccountType_PFG
'        End If
'    End If
    
    If IsBrokerUser(eTT_AccountType_RcgPats) Then
        If g.RcgPats Is Nothing Then
            Set g.RcgPats = New cBrokerPats
            g.RcgPats.Init eTT_AccountType_RcgPats
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_Rithmic) Then
        If g.Rithmic Is Nothing Then
            Set g.Rithmic = New cRithmic
            g.Rithmic.Init eTT_AccountType_Rithmic
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_RjoCqg, bIsBroker) Then
        If g.RjoCqg Is Nothing Then
            Set g.RjoCqg = New cBrokerCqg
            g.RjoCqg.Init eTT_AccountType_RjoCqg, bIsBroker
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_RjoHkPats) Then
        If g.RjoHkPats Is Nothing Then
            Set g.RjoHkPats = New cBrokerPats
            g.RjoHkPats.Init eTT_AccountType_RjoHkPats
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_RjoPats) Then
        If g.RjoPats Is Nothing Then
            Set g.RjoPats = New cBrokerPats
            g.RjoPats.Init eTT_AccountType_RjoPats
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_RobbinsCqg) Then
        If g.RobbinsCqg Is Nothing Then
            Set g.RobbinsCqg = New cBrokerCqg
            g.RobbinsCqg.Init eTT_AccountType_RobbinsCqg, bIsBroker
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_SimStream) Then
        If g.SimTradeStream Is Nothing Then
            Set g.SimTradeStream = New cSimTradeStream
            g.SimTradeStream.Init eTT_AccountType_SimStream
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_SimReplay) Then
        If g.SimTradeReplay Is Nothing Then
            Set g.SimTradeReplay = New cSimTradeStream
            g.SimTradeReplay.Init eTT_AccountType_SimReplay
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_SimBroker) Then
        If g.SimTradeTs Is Nothing Then
            Set g.SimTradeTs = New cSimTradeTs
            g.SimTradeTs.Init eTT_AccountType_SimBroker, frmOnlineBroker.txtSalmonCallbackTs, frmOnlineBroker.tmrTradeServer
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_Tradier) Then
        If g.Tradier Is Nothing Then
            Set g.Tradier = New cBrokerTradier
            g.Tradier.Init eTT_AccountType_Tradier
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_TransAct) Then
        If g.Transact Is Nothing Then
            Set g.Transact = New cTransact
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_TT) Then
        If g.TT Is Nothing Then
            Set g.TT = New cBrokerTt
            g.TT.Init eTT_AccountType_TT
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_VanKarCurrenex) Then
        If g.VanKarCnx Is Nothing Then
            Set g.VanKarCnx = New cBrokerCurrenex
            g.VanKarCnx.Init eTT_AccountType_VanKarCurrenex
        End If
    End If

    If IsBrokerUser(eTT_AccountType_Vision) Then
        If g.Vision Is Nothing Then
            Set g.Vision = New cRithmic
            g.Vision.Init eTT_AccountType_Vision
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_VisionCqg) Then
        If g.VisionCqg Is Nothing Then
            Set g.VisionCqg = New cBrokerCqg
            g.VisionCqg.Init eTT_AccountType_VisionCqg
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_ZanerCqg) Then
        If g.ZanerCqg Is Nothing Then
            Set g.ZanerCqg = New cBrokerCqg
            g.ZanerCqg.Init eTT_AccountType_ZanerCqg, bIsBroker
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_ZanerCurrenex) Then
        If g.ZanerCnx Is Nothing Then
            Set g.ZanerCnx = New cBrokerCurrenex
            g.ZanerCnx.Init eTT_AccountType_ZanerCurrenex
        End If
    End If

    If IsBrokerUser(eTT_AccountType_ZanerPats) Then
        If g.ZanerPats Is Nothing Then
            Set g.ZanerPats = New cBrokerPats
            g.ZanerPats.Init eTT_AccountType_ZanerPats
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_ZanerRithmic) Then
        If g.ZanerRithmic Is Nothing Then
            Set g.ZanerRithmic = New cRithmic
            g.ZanerRithmic.Init eTT_AccountType_ZanerRithmic
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_ZanerZenFire) Then
        If g.ZanerZenFire Is Nothing Then
            Set g.ZanerZenFire = New cRithmic
            g.ZanerZenFire.Init eTT_AccountType_ZanerZenFire
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_ZenFire) Then
        If g.ZenFire Is Nothing Then
            Set g.ZenFire = New cRithmic
            g.ZenFire.Init eTT_AccountType_ZenFire
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.InitBrokerObjects"
    Resume
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DestroyBrokerObjects
'' Description: Destroy the broker objects as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DestroyBrokerObjects()
On Error GoTo ErrSection:

    If Not g.AdvFut Is Nothing Then
        g.AdvFut.Broker.Disconnect False, "Online Broker Form Unloading"
        g.AdvFut.UnloadApp
        Set g.AdvFut = Nothing
    End If

    If Not g.AlpariCnx Is Nothing Then
        g.AlpariCnx.Broker.Disconnect False, "Online Broker Form Unloading"
        g.AlpariCnx.UnloadApp
        Set g.AlpariCnx = Nothing
    End If
    
    If Not g.AlpariPats Is Nothing Then
        g.AlpariPats.Broker.Disconnect False, "Online Broker Form Unloading"
        g.AlpariPats.UnloadApp
        Set g.AlpariPats = Nothing
    End If
    
    If Not g.AlpariZenFire Is Nothing Then
        g.AlpariZenFire.Broker.Disconnect False, "Online Broker Form Unloading"
        g.AlpariZenFire.UnloadApp
        Set g.AlpariZenFire = Nothing
    End If

    If Not g.Ameritrade Is Nothing Then
        g.Ameritrade.Broker.Disconnect False, "Online Broker Form Unloading"
        g.Ameritrade.UnloadApp
        Set g.Ameritrade = Nothing
    End If

    If Not g.AmpCqg Is Nothing Then
        g.AmpCqg.Broker.Disconnect False, "Online Broker Form Unloading"
        g.AmpCqg.UnloadApp
        Set g.AmpCqg = Nothing
    End If
    
    If Not g.BornPats Is Nothing Then
        g.BornPats.Broker.Disconnect False, "Online Broker Form Unloading"
        g.BornPats.UnloadApp
        Set g.BornPats = Nothing
    End If
    
    If Not g.CQG Is Nothing Then
        g.CQG.Broker.Disconnect False, "Online Broker Form Unloading"
        g.CQG.UnloadApp
        Set g.CQG = Nothing
    End If
    
    If Not g.CtgCqg Is Nothing Then
        g.CtgCqg.Broker.Disconnect False, "Online Broker Form Unloading"
        g.CtgCqg.UnloadApp
        Set g.CtgCqg = Nothing
    End If
    
    If Not g.CtgPats Is Nothing Then
        g.CtgPats.Broker.Disconnect False, "Online Broker Form Unloading"
        g.CtgPats.UnloadApp
        Set g.CtgPats = Nothing
    End If
    
'    If Not g.CtgPfg Is Nothing Then
'        g.CtgPfg.Disconnect False, "Online Broker Form Unloading"
'        g.CtgPfg.SendPfgMessage eGDPfgMessageType_UnloadApp, ""
'        Set g.CtgPfg = Nothing
'    End If
    
    If Not g.DemoPats Is Nothing Then
        g.DemoPats.Broker.Disconnect False, "Online Broker Form Unloading"
        g.DemoPats.UnloadApp
        Set g.DemoPats = Nothing
    End If
    
    If Not g.Etrade Is Nothing Then
        g.Etrade.Broker.Disconnect False, "Online Broker Form Unloading"
        g.Etrade.UnloadApp
        Set g.Etrade = Nothing
    End If
    
    If Not g.FptCqg Is Nothing Then
        g.FptCqg.Broker.Disconnect False, "Online Broker Form Unloading"
        g.FptCqg.UnloadApp
        Set g.FptCqg = Nothing
    End If
    
    If Not g.FptOec Is Nothing Then
        g.FptOec.Broker.Disconnect False, "Online Broker Form Unloading"
        g.FptOec.UnloadApp
        Set g.FptOec = Nothing
    End If
    
'    If Not g.FintecPfg Is Nothing Then
'        g.FintecPfg.Disconnect False, "Online Broker Form Unloading"
'        g.FintecPfg.SendPfgMessage eGDPfgMessageType_UnloadApp, ""
'        Set g.FintecPfg = Nothing
'    End If
    
    If Not g.Ideal Is Nothing Then
        g.Ideal.Broker.Disconnect False, "Online Broker Form Unloading"
        g.Ideal.UnloadApp
        Set g.Ideal = Nothing
    End If
    
    If Not g.IntBroker Is Nothing Then
        g.IntBroker.Broker.Disconnect False, "Online Broker Form Unloading"
        g.IntBroker.UnloadApp
        Set g.IntBroker = Nothing
    End If

    If Not g.KnightCnx Is Nothing Then
        g.KnightCnx.Broker.Disconnect False, "Online Broker Form Unloading"
        g.KnightCnx.UnloadApp
        Set g.KnightCnx = Nothing
    End If
    
    If Not g.KnightCqg Is Nothing Then
        g.KnightCqg.Broker.Disconnect False, "Online Broker Form Unloading"
        g.KnightCqg.UnloadApp
        Set g.KnightCqg = Nothing
    End If
    
'    If Not g.LindWaldock Is Nothing Then
'        g.LindWaldock.Disconnect
'        g.LindWaldock.SendLindXpressMessage eGDLindXpressMessageType_UnloadApp, ""
'        Set g.LindWaldock = Nothing
'    End If
'
'    If Not g.ManExpress Is Nothing Then
'        g.ManExpress.Disconnect
'        g.ManExpress.SendLindXpressMessage eGDLindXpressMessageType_UnloadApp, ""
'        Set g.ManExpress = Nothing
'    End If
    
    If Not g.Oec Is Nothing Then
        g.Oec.Broker.Disconnect False, "Online Broker Form Unloading"
        g.Oec.UnloadApp
        Set g.Oec = Nothing
    End If
    
    If Not g.Optimus Is Nothing Then
        g.Optimus.Broker.Disconnect False, "Online Broker Form Unloading"
        g.Optimus.UnloadApp
        Set g.Optimus = Nothing
    End If
    
    If Not g.OpVest Is Nothing Then
        g.OpVest.Broker.Disconnect False, "Online Broker Form Unloading"
        g.OpVest.UnloadApp
        Set g.OpVest = Nothing
    End If
    
    If Not g.Pats Is Nothing Then
        g.Pats.Broker.Disconnect False, "Online Broker Form Unloading"
        g.Pats.UnloadApp
        Set g.Pats = Nothing
    End If
    
'    If Not g.PFG Is Nothing Then
'        g.PFG.Disconnect False, "Online Broker Form Unloading"
'        g.PFG.SendPfgMessage eGDPfgMessageType_UnloadApp, ""
'        Set g.PFG = Nothing
'    End If
    
    If Not g.RcgPats Is Nothing Then
        g.RcgPats.Broker.Disconnect False, "Online Broker Form Unloading"
        g.RcgPats.UnloadApp
        Set g.RcgPats = Nothing
    End If
    
    If Not g.Rithmic Is Nothing Then
        g.Rithmic.Broker.Disconnect False, "Online Broker Form Unloading"
        g.Rithmic.UnloadApp
        Set g.Rithmic = Nothing
    End If
    
    If Not g.RjoCqg Is Nothing Then
        g.RjoCqg.Broker.Disconnect False, "Online Broker Form Unloading"
        g.RjoCqg.UnloadApp
        Set g.RjoCqg = Nothing
    End If
    
    If Not g.RjoHkPats Is Nothing Then
        g.RjoHkPats.Broker.Disconnect False, "Online Broker Form Unloading"
        g.RjoHkPats.UnloadApp
        Set g.RjoHkPats = Nothing
    End If
    
    If Not g.RjoPats Is Nothing Then
        g.RjoPats.Broker.Disconnect False, "Online Broker Form Unloading"
        g.RjoPats.UnloadApp
        Set g.RjoPats = Nothing
    End If
    
    If Not g.RobbinsCqg Is Nothing Then
        g.RobbinsCqg.Broker.Disconnect False, "Online Broker Form Unloading"
        g.RobbinsCqg.UnloadApp
        Set g.RobbinsCqg = Nothing
    End If
    
    If Not g.Tradier Is Nothing Then
        g.Tradier.Broker.Disconnect False, "Online Broker Form Unloading"
        g.Tradier.UnloadApp
        Set g.Tradier = Nothing
    End If
    
    If Not g.Transact Is Nothing Then
        g.Transact.Disconnect "", False, "Online Broker form unloading"
        g.Transact.SendTransactMessage eGDTransactMessageType_UnloadApp, ""
        Set g.Transact = Nothing
    End If
    
    If Not g.TT Is Nothing Then
        g.TT.Broker.Disconnect False, "Online Broker Form Unloading"
        g.TT.UnloadApp
        Set g.TT = Nothing
    End If
    
    If Not g.Vision Is Nothing Then
        g.Vision.Broker.Disconnect False, "Online Broker Form Unloading"
        g.Vision.UnloadApp
        Set g.Vision = Nothing
    End If
    
    If Not g.VisionCqg Is Nothing Then
        g.VisionCqg.Broker.Disconnect False, "Online Broker Form Unloading"
        g.VisionCqg.UnloadApp
        Set g.VisionCqg = Nothing
    End If
    
    If Not g.ZanerCqg Is Nothing Then
        g.ZanerCqg.Broker.Disconnect False, "Online Broker Form Unloading"
        g.ZanerCqg.UnloadApp
        Set g.ZanerCqg = Nothing
    End If
    
    If Not g.ZanerPats Is Nothing Then
        g.ZanerPats.Broker.Disconnect False, "Online Broker Form Unloading"
        g.ZanerPats.UnloadApp
        Set g.ZanerPats = Nothing
    End If
    
    If Not g.ZanerRithmic Is Nothing Then
        g.ZanerRithmic.Broker.Disconnect False, "Online Broker Form Unloading"
        g.ZanerRithmic.UnloadApp
        Set g.ZanerRithmic = Nothing
    End If
    
    If Not g.ZanerZenFire Is Nothing Then
        g.ZanerZenFire.Broker.Disconnect False, "Online Broker Form Unloading"
        g.ZanerZenFire.UnloadApp
        Set g.ZanerZenFire = Nothing
    End If

    If Not g.ZenFire Is Nothing Then
        g.ZenFire.Broker.Disconnect False, "Online Broker Form Unloading"
        g.ZenFire.UnloadApp
        Set g.ZenFire = Nothing
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.DestroyBrokerObjects"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPosition
'' Description: Determine the current position for the information given
'' Inputs:      Account, Symbol, Auto Trade ID
'' Returns:     Current Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentPosition(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant, ByVal lAtID As Long) As Long
On Error GoTo ErrSection:

    Dim strAccountNumber As String      ' Account number
    Dim strSymbol As String             ' Symbol
    Dim nBroker As eTT_AccountType      ' Broker
    Dim lReturn As Long                 ' Return value for the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    strAccountNumber = GetAccountNumber(vAccountIDorNumber)
    strSymbol = GetSymbol(vSymbolIDOrSymbol)
    nBroker = AccountTypeForNumber(strAccountNumber)
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        lReturn = BInfo.CurrentPosition(strAccountNumber, strSymbol, lAtID)
    End If
    
    CurrentPosition = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CurrentPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentBrokerPosition
'' Description: Determine the current broker position for the information given
'' Inputs:      Account, Symbol
'' Returns:     Current Broker Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentBrokerPosition(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant) As Long
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strAccountNumber As String      ' Account number for given account
    Dim strSymbol As String             ' Symbol for the given symbol
    Dim BrokerObj As cBroker            ' Broker object
    Dim lReturn As Long                 ' Return value for the function
    
    strAccountNumber = GetAccountNumber(vAccountNumberOrID)
    nBroker = AccountTypeForNumber(strAccountNumber)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    lReturn = 0&

    Select Case nBroker
        Case eTT_AccountType_AdvFut
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                lReturn = g.CtgPfg.CurrentPosition(strAccountNumber, strSymbol)
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                lReturn = g.FintecPfg.CurrentPosition(strAccountNumber, strSymbol)
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                lReturn = g.LindWaldock.CurrentPosition(strAccountNumber, strSymbol)
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                lReturn = g.ManExpress.CurrentPosition(strAccountNumber, strSymbol)
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                lReturn = g.PFG.CurrentPosition(strAccountNumber, strSymbol)
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                lReturn = g.Transact.CurrentPosition(strAccountNumber, strSymbol)
            End If
            
        Case eTT_AccountType_TT
    
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                lReturn = BrokerObj.CurrentPosition(strAccountNumber, strSymbol)
            End If
    
    End Select
    
    CurrentBrokerPosition = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CurrentBrokerPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CarriedBrokerPosition
'' Description: Determine the carried broker position for the information given
'' Inputs:      Account, Symbol
'' Returns:     Carried Broker Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CarriedBrokerPosition(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant) As Long
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strAccountNumber As String      ' Account number for given account
    Dim strSymbol As String             ' Symbol for the given symbol
    Dim BrokerObj As cBroker            ' Broker object
    Dim lReturn As Long                 ' Return value for the function
    
    strAccountNumber = GetAccountNumber(vAccountNumberOrID)
    nBroker = AccountTypeForNumber(strAccountNumber)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    lReturn = 0&

    Select Case nBroker
        Case eTT_AccountType_AdvFut
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                lReturn = g.CtgPfg.CarriedPosition(strAccountNumber, strSymbol)
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                lReturn = g.FintecPfg.CarriedPosition(strAccountNumber, strSymbol)
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                lReturn = g.LindWaldock.CarriedPosition(strAccountNumber, strSymbol)
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                lReturn = g.ManExpress.CarriedPosition(strAccountNumber, strSymbol)
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                lReturn = g.PFG.CarriedPosition(strAccountNumber, strSymbol)
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                lReturn = g.Transact.CarriedPosition(strAccountNumber, strSymbol)
            End If
            
        Case eTT_AccountType_TT
    
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                lReturn = BrokerObj.CarriedPosition(strAccountNumber, strSymbol)
            End If
    
    End Select
    
    CarriedBrokerPosition = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CarriedBrokerPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixPositions
'' Description: Allow the user to attempt to fix carried position mismatches
'' Inputs:      Only do first timers
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FixPositions(Optional ByVal bOnlyFirstTime As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim BrokerObj As cBroker            ' Broker object
    Dim bContinue As Boolean            ' Continue?

    If (g.bStarting = False) Then
        If (g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected) Or (g.RealTime.ConnectionStatus = eGDConnectionStatus_Disconnected) Then
            For lIndex = 1 To kNumBrokers - 1
                If IsBrokerUser(lIndex) Then
                    If (ConnectionStatusForBroker(lIndex) = eGDConnectionStatus_Connected) Then
                        If Refreshing(lIndex) = False Then
                            Select Case lIndex
'                                Case eTT_AccountType_CtgPfg
'                                    If Not g.CtgPfg Is Nothing Then
'                                        g.CtgPfg.FixPositions bOnlyFirstTime
'                                    End If
'
'                                Case eTT_AccountType_FintecPfg
'                                    If Not g.FintecPfg Is Nothing Then
'                                        g.FintecPfg.FixPositions bOnlyFirstTime
'                                    End If
'
'                                Case eTT_AccountType_LindWaldock
'                                    If Not g.LindWaldock Is Nothing Then
'                                        g.LindWaldock.FixPositions bOnlyFirstTime
'                                    End If
'
'                                Case eTT_AccountType_ManExpress
'                                    If Not g.ManExpress Is Nothing Then
'                                        g.ManExpress.FixPositions bOnlyFirstTime
'                                    End If
'
'                                Case eTT_AccountType_PFG
'                                    If Not g.PFG Is Nothing Then
'                                        g.PFG.FixPositions bOnlyFirstTime
'                                    End If
                                    
                                Case eTT_AccountType_TransAct
                                    If Not g.Transact Is Nothing Then
                                        g.Transact.FixPositions bOnlyFirstTime
                                    End If
                                    
                                Case Else
                                    Set BrokerObj = Broker(lIndex)
                                    If Not BrokerObj Is Nothing Then
                                        BrokerObj.FixPositions bOnlyFirstTime
                                    End If
                                    
                            End Select
                        End If
                    End If
                End If
            Next lIndex
        End If
    End If
                    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.FixPositions"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixPosition
'' Description: Allow the user to try to fix the carried position mismatch
'' Inputs:      Account, Symbol
'' Returns:     True if Fixed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FixPosition(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant) As Boolean
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim BrokerObj As cBroker            ' Broker object
    Dim bReturn As Boolean              ' Return value for the function
    
    nBroker = AccountTypeForNumber(GetAccountNumber(vAccountNumberOrID))
    bReturn = True

    Select Case nBroker
        Case eTT_AccountType_AdvFut
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                bReturn = g.CtgPfg.FixPosition(vAccountNumberOrID, vSymbolOrSymbolID)
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                bReturn = g.FintecPfg.FixPosition(vAccountNumberOrID, vSymbolOrSymbolID)
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                bReturn = g.LindWaldock.FixPosition(vAccountNumberOrID, vSymbolOrSymbolID)
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                bReturn = g.ManExpress.FixPosition(vAccountNumberOrID, vSymbolOrSymbolID)
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                bReturn = g.PFG.FixPosition(vAccountNumberOrID, vSymbolOrSymbolID)
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                bReturn = g.Transact.FixPosition(vAccountNumberOrID, vSymbolOrSymbolID)
            End If
            
        Case eTT_AccountType_TT
    
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                bReturn = BrokerObj.FixPosition(vAccountNumberOrID, vSymbolOrSymbolID)
            End If
    
    End Select
    
    FixPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.FixPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradesForSymbol
'' Description: Get the trade-by-trade file for the information given
'' Inputs:      Account, Symbol, Auto Trade ID, Period
'' Returns:     Trade-by-trade For Symbol (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TradesForSymbol(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant, ByVal lAtID As Long, Optional strPeriod As String = "Daily") As cGdArray
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID
    Dim strAccountNumber As String      ' Account number
    Dim strSymbol As String             ' Symbol
    Dim nBroker As eTT_AccountType      ' Broker
    Dim Account As cPtAccount           ' Account from the Trade console
    Dim astrReturn As cGdArray          ' Return value for the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    lAccountID = GetAccountID(vAccountIDorNumber)
    strAccountNumber = GetAccountNumber(vAccountIDorNumber)
    strSymbol = ConvertToTradeSymbol(GetSymbol(vSymbolIDOrSymbol), Int(CurrentTime("", "", True)))
    nBroker = AccountTypeForNumber(strAccountNumber)

    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set astrReturn = BInfo.TradesForSymbol(strAccountNumber, strSymbol, lAtID, strPeriod)
    End If
    
    Set TradesForSymbol = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.TradesForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToTradesFile
'' Description: Dump the trade-by-trade file for the information given to a file
'' Inputs:      Filename, Account, Symbol, Auto Trade ID, Period
'' Returns:     True if dumped, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToTradesFile(ByVal strFileName As String, ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant, ByVal lAtID As Long, Optional strPeriod As String = "Daily") As Boolean
On Error GoTo ErrSection:

    Dim astrTrades As cGdArray          ' Trade-by-trade array to dump
    Dim bReturn As Boolean              ' Return value from the function
    
    Set astrTrades = TradesForSymbol(vAccountIDorNumber, vSymbolIDOrSymbol, lAtID, strPeriod)
    If Not astrTrades Is Nothing Then
        astrTrades.ToFile strFileName
        bReturn = True
    End If
    
    ToTradesFile = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ToTradesFile"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTradeableSymbol
'' Description: Is the given symbol tradeable for the given account?
'' Inputs:      Account, Symbol
'' Returns:     True if Tradeable Symbol, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsTradeableSymbol(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant) As Boolean
On Error GoTo ErrSection:

    Dim strAccountNumber As String      ' Account number
    Dim strSymbol As String             ' Symbol
    Dim nBroker As eTT_AccountType      ' Broker
    Dim BrokerObj As cBroker            ' Broker object
    Dim bReturn As Boolean              ' Return value from the function
    
    strAccountNumber = GetAccountNumber(vAccountIDorNumber)
    strSymbol = ConvertToTradeSymbol(GetSymbol(vSymbolIDOrSymbol), Int(CurrentTime("", "", True)))
    nBroker = AccountTypeForNumber(strAccountNumber)
    bReturn = False

    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                bReturn = g.CtgPfg.IsTradeableSymbol(strSymbol)
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                bReturn = g.FintecPfg.IsTradeableSymbol(strSymbol)
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                bReturn = g.LindWaldock.IsTradeableSymbol(strSymbol)
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                bReturn = g.ManExpress.IsTradeableSymbol(strSymbol)
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                bReturn = g.PFG.IsTradeableSymbol(strSymbol)
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                bReturn = g.Transact.IsTradeableSymbol(strSymbol)
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                bReturn = BrokerObj.IsTradeableSymbol(strSymbol)
            End If
    
    End Select
    
    IsTradeableSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsTradeableSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsEnabledSymbol
'' Description: Is the given symbol enabled for trading for the given account?
'' Inputs:      Account, Symbol, Broker Base, Broker Exchange
'' Returns:     True if Enabled Symbol, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEnabledSymbol(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant, Optional strBrokerBase As String, Optional strBrokerExchange As String) As Boolean
On Error GoTo ErrSection:

    Dim strAccountNumber As String      ' Account number
    Dim strSymbol As String             ' Symbol
    Dim nBroker As eTT_AccountType      ' Broker
    Dim BrokerObj As cBroker            ' Broker object
    Dim bReturn As Boolean              ' Return value from the function
    
    strAccountNumber = GetAccountNumber(vAccountIDorNumber)
    strSymbol = ConvertToTradeSymbol(GetSymbol(vSymbolIDOrSymbol), Int(CurrentTime("", "", True)))
    nBroker = AccountTypeForNumber(strAccountNumber)
    bReturn = False

    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                bReturn = g.CtgPfg.IsEnabledSymbol(strSymbol, strBrokerBase, strBrokerExchange)
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                bReturn = g.FintecPfg.IsEnabledSymbol(strSymbol, strBrokerBase, strBrokerExchange)
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                bReturn = g.LindWaldock.IsEnabledSymbol(strSymbol, strBrokerBase, strBrokerExchange)
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                bReturn = g.ManExpress.IsEnabledSymbol(strSymbol, strBrokerBase, strBrokerExchange)
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                bReturn = g.PFG.IsEnabledSymbol(strSymbol, strBrokerBase, strBrokerExchange)
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                bReturn = g.Transact.IsEnabledSymbol(strSymbol, strBrokerBase, strBrokerExchange)
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                bReturn = BrokerObj.IsEnabledSymbol(strSymbol, strBrokerBase, strBrokerExchange)
            End If
    
    End Select
    
    IsEnabledSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsEnabledSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnabledSymbols
'' Description: Get the list of symbols the user is enabled to trade
'' Inputs:      Broker
'' Returns:     Enabled Symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnabledSymbols(ByVal nBroker As eTT_AccountType) As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Return value for the function
    Dim BrokerObj As cBroker            ' Broker object

    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                Set astrReturn = g.CtgPfg.EnabledSymbols
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                Set astrReturn = g.FintecPfg.EnabledSymbols
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                Set astrReturn = g.LindWaldock.EnabledSymbols
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                Set astrReturn = g.ManExpress.EnabledSymbols
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                Set astrReturn = g.PFG.EnabledSymbols
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                Set astrReturn = g.Transact.EnabledSymbols
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                Set astrReturn = BrokerObj.EnabledSymbols
            End If
    
    End Select
    
    Set EnabledSymbols = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.EnabledSymbols"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsPitSymbol
'' Description: Is the given symbol a pit symbol?
'' Inputs:      Account, Symbol
'' Returns:     True if Pit Symbol, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsPitSymbol(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant) As Boolean
On Error GoTo ErrSection:

    Dim strAccountNumber As String      ' Account number
    Dim strSymbol As String             ' Symbol
    Dim nBroker As eTT_AccountType      ' Broker
    Dim bReturn As Boolean              ' Return value from the function
    
    strAccountNumber = GetAccountNumber(vAccountIDorNumber)
    strSymbol = ConvertToTradeSymbol(GetSymbol(vSymbolIDOrSymbol), Int(CurrentTime("", "", True)))
    nBroker = AccountTypeForNumber(strAccountNumber)
    bReturn = False

    IsPitSymbol = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsPitSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanTrade
'' Description: Can the given symbol be traded right now in the given account?
'' Inputs:      Account, Symbol
'' Returns:     True if Can be Traded, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CanTrade(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant) As Boolean
On Error GoTo ErrSection:

    Dim strAccountNumber As String      ' Account number
    Dim strSymbol As String             ' Symbol
    Dim nBroker As eTT_AccountType      ' Broker
    Dim BrokerObj As cBroker            ' Broker object
    Dim bReturn As Boolean              ' Return value from the function
    
    strAccountNumber = GetAccountNumber(vAccountIDorNumber)
    strSymbol = ConvertToTradeSymbol(GetSymbol(vSymbolIDOrSymbol), Int(CurrentTime("", "", True)))
    nBroker = AccountTypeForNumber(strAccountNumber)
    
    bReturn = False
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                bReturn = g.CtgPfg.IsTradeableSymbol(strSymbol)
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                bReturn = g.FintecPfg.IsTradeableSymbol(strSymbol)
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.ManExpress Is Nothing Then
'                bReturn = g.LindWaldock.IsTradeableSymbol(strSymbol)
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                bReturn = g.ManExpress.IsTradeableSymbol(strSymbol)
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                bReturn = g.PFG.IsTradeableSymbol(strSymbol)
'            End If
        
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                bReturn = g.Transact.IsTradeableSymbol(strSymbol)
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                bReturn = BrokerObj.IsTradeableSymbol(strSymbol)
            End If
    
    End Select
        
    CanTrade = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CanTrade"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingOrders
'' Description: Are there working orders for the given account and symbol?
'' Inputs:      Account, Symbol, Auto Trade Item ID, Skip Order ID
'' Returns:     True if Working Orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingOrders(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant, ByVal lAtID As Long, Optional ByVal lSkipOrderID As Long = -1&) As Boolean
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol
    Dim bReturn As Boolean              ' Return value from the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
        
    strAccount = GetAccountNumber(vAccountIDorNumber)
    nBroker = AccountTypeForNumber(strAccount)
    vSymbolIDOrSymbol = ConvertToTradeSymbol(vSymbolIDOrSymbol, Int(CurrentTime("", "", True)))
    strSymbol = GetSymbol(vSymbolIDOrSymbol)
    bReturn = False
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        bReturn = BInfo.HasWorkingOrdersForSymbol(strAccount, strSymbol, lAtID, lSkipOrderID)
    End If
    
    HasWorkingOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.HasWorkingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasPendingOrders
'' Description: Are there pending orders for the given account and symbol?
'' Inputs:      Account, Symbol, Auto Trade Item ID
'' Returns:     True if Pending Orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasPendingOrders(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant, ByVal lAtID As Long) As Boolean
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol
    Dim bReturn As Boolean              ' Return value from the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
        
    strAccount = GetAccountNumber(vAccountIDorNumber)
    nBroker = AccountTypeForNumber(strAccount)
    vSymbolIDOrSymbol = ConvertToTradeSymbol(vSymbolIDOrSymbol, Int(CurrentTime("", "", True)))
    strSymbol = GetSymbol(vSymbolIDOrSymbol)
    bReturn = False
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        bReturn = BInfo.HasPendingOrdersForSymbol(strAccount, strSymbol, lAtID)
    End If
    
    HasPendingOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.HasPendingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionMatch
'' Description: Do we have a position match for the given symbol?
'' Inputs:      Account, Symbol
'' Returns:     True if Positions Match, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionMatch(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant) As Boolean
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol
    Dim bReturn As Boolean              ' Return value from the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    strAccount = GetAccountNumber(vAccountIDorNumber)
    nBroker = AccountTypeForNumber(strAccount)
    vSymbolIDOrSymbol = ConvertToTradeSymbol(vSymbolIDOrSymbol, Int(CurrentTime("", "", True)))
    strSymbol = GetSymbol(vSymbolIDOrSymbol)
    bReturn = True

    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        bReturn = BInfo.PositionMatch(strAccount, strSymbol)
    End If
    
    PositionMatch = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.PositionMatch"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CarriedMatch
'' Description: Do we have a carried position match for the given symbol?
'' Inputs:      Account, Symbol
'' Returns:     True if Carried Positions Match, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CarriedMatch(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant) As Boolean
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol
    Dim bReturn As Boolean              ' Return value from the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    strAccount = GetAccountNumber(vAccountIDorNumber)
    nBroker = AccountTypeForNumber(strAccount)
    vSymbolIDOrSymbol = ConvertToTradeSymbol(vSymbolIDOrSymbol, Int(CurrentTime("", "", True)))
    strSymbol = GetSymbol(vSymbolIDOrSymbol)
    bReturn = True

    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        bReturn = BInfo.CarriedMatch(strAccount, strSymbol)
    End If
    
    CarriedMatch = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CarriedMatch"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConsistentBroker
'' Description: Do we have consistent broker data for the given symbol?
'' Inputs:      Account, Symbol
'' Returns:     True if Consistent Broker Data, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConsistentBroker(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant) As Boolean
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol
    Dim bReturn As Boolean              ' Return value from the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    strAccount = GetAccountNumber(vAccountIDorNumber)
    nBroker = AccountTypeForNumber(strAccount)
    vSymbolIDOrSymbol = ConvertToTradeSymbol(vSymbolIDOrSymbol, Int(CurrentTime("", "", True)))
    strSymbol = GetSymbol(vSymbolIDOrSymbol)
    bReturn = True

    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        bReturn = BInfo.ConsistentBroker(strAccount, strSymbol)
    End If
    
    ConsistentBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ConsistentBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConfirmOrder
'' Description: Do we need to confirm the order with the user first?
'' Inputs:      Account, Symbol
'' Returns:     True if Confirm Order, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConfirmOrder(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim strAccount As String            ' Account Number
    Dim nBroker As eTT_AccountType      ' Account type for the account passed in
    
    bReturn = False
    strAccount = GetAccountNumber(vAccountIDorNumber)
    nBroker = AccountTypeForNumber(strAccount)
    
    If (ConfirmManual = True) Or (g.RealTime.Active = False) Then
        bReturn = True
    ElseIf (IsPitSymbol(vAccountIDorNumber, ConvertToTradeSymbol(vSymbolIDOrSymbol, Int(CurrentTime("", "", True)))) = True) And (IsLiveAccount(nBroker) = True) Then
        bReturn = True
    ElseIf Not g.CattleBridge Is Nothing Then
        bReturn = g.CattleBridge.ConfirmOrder(strAccount, nBroker)
    End If
    
    ConfirmOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ConfirmOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerDebug
'' Description: Send a message to the appropriate debug log
'' Inputs:      Broker, Message, Send to Activity Log as well?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BrokerDebug(ByVal nBroker As eTT_AccountType, ByVal strMessage As String, Optional ByVal bActivityLog As Boolean = False)
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker object
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then g.CtgPfg.DumpDebug strMessage
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then g.FintecPfg.DumpDebug strMessage
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then g.LindWaldock.DumpDebug strMessage
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then g.ManExpress.DumpDebug strMessage
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then g.PFG.DumpDebug strMessage
        
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then g.Transact.DumpDebug strMessage
        
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                BrokerObj.DumpDebug strMessage
            End If
    
    End Select
    
    If bActivityLog Then
        g.ActivityLogs.AddActivityLogEntry nBroker, strMessage
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.BrokerDebug"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerDebugAll
'' Description: Send a message to all of the appropriate broker logs
'' Inputs:      Message, Send to Activity Log as well?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BrokerDebugAll(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To kNumBrokers - 1
        BrokerDebug lIndex, strMessage
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.BrokerDebugAll"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTypeAllowed
'' Description: Is the given order type allowed for the given account and symbol?
'' Inputs:      Account, Symbol, Order Type
'' Returns:     True if Allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderTypeAllowed(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant, ByVal nOrderType As eTT_OrderType) As Boolean
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol passed in
    Dim BrokerObj As cBroker            ' Broker object
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = True
    nBroker = AccountTypeForID(GetAccountID(vAccountIDorNumber))
    strSymbol = GetSymbol(vSymbolIDOrSymbol)

    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                bReturn = g.CtgPfg.OrderTypeAllowed(nOrderType, strSymbol)
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                bReturn = g.FintecPfg.OrderTypeAllowed(nOrderType, strSymbol)
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                bReturn = g.LindWaldock.OrderTypeAllowed(nOrderType, strSymbol)
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                bReturn = g.ManExpress.OrderTypeAllowed(nOrderType, strSymbol)
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                bReturn = g.PFG.OrderTypeAllowed(nOrderType, strSymbol)
'            End If
        
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                bReturn = g.Transact.OrderTypeAllowed(nOrderType, strSymbol)
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                bReturn = BrokerObj.OrderTypeAllowed(nOrderType, strSymbol)
            End If
    
    End Select
    
    OrderTypeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.OrderTypeAllowed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimeInForceAllowed
'' Description: Is the given time in force allowed for the given account and symbol?
'' Inputs:      Account, Symbol, Time In Force
'' Returns:     True if Allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TimeInForceAllowed(ByVal vAccountIDorNumber As Variant, ByVal vSymbolIDOrSymbol As Variant, ByVal nTimeInForce As eTT_TimeInForce) As Boolean
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim strSymbol As String             ' Symbol passed in
    Dim BrokerObj As cBroker            ' Broker object
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = True
    nBroker = AccountTypeForID(GetAccountID(vAccountIDorNumber))
    strSymbol = GetSymbol(vSymbolIDOrSymbol)

    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                bReturn = g.CtgPfg.TimeInForceAllowed(nTimeInForce, strSymbol)
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                bReturn = g.FintecPfg.TimeInForceAllowed(nTimeInForce, strSymbol)
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                bReturn = g.LindWaldock.TimeInForceAllowed(nTimeInForce, strSymbol)
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                bReturn = g.ManExpress.TimeInForceAllowed(nTimeInForce, strSymbol)
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                bReturn = g.PFG.TimeInForceAllowed(nTimeInForce, strSymbol)
'            End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                bReturn = g.Transact.TimeInForceAllowed(nTimeInForce, strSymbol)
            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                bReturn = BrokerObj.TimeInForceAllowed(nTimeInForce, strSymbol)
            End If
    
    End Select
    
    TimeInForceAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.TimeInForceAllowed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendOrder
'' Description: Send the order to the appropriate place
'' Inputs:      Order to Send, Previous Order ID, Force a Send, Force Submit OTO's?
'' Returns:     Order ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SendOrder(Order As cPtOrder, Optional ByVal strPrevGenesisOrderID As String = "", Optional ByVal bForceSend As Boolean = False, Optional ByVal bForceSubmitOtos As Boolean = False) As Long
On Error GoTo ErrSection:

gdStartProfile 901

    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim Bars As New cGdBars             ' Temporary bars object
    Dim lReturn As Long                 ' Return value from a function
    Dim OldOrder As cPtOrder            ' Old order
    Dim OtherOrder As cPtOrder          ' Other order in an OCO situation
    Dim BrokerObj As cBroker            ' Broker object
    
    nBroker = AccountTypeForID(Order.AccountID)
    SetBarProperties Bars, Order.SymbolOrSymbolID
    
    ' If no order date has been filled in yet, fill it in now...
    If Order.OrderDate = 0# Then
        Order.OrderDate = Order.BrokerDate(CurrentTime("", Order.Symbol))
        Order.Save
        OrderCallback Order
    End If
    
    ' If this was a previously submitted order that got parked...
    ' 01/25/2011 DAJ: No need to go into this block if the order has never been
    ' sent to the broker (i.e. there is not BrokerID)...
    If (Order.Status = eTT_OrderStatus_Parked) And (Len(Order.BrokerID) > 0) Then
        ' If this is one side of an OCO, load up the other order so that we can
        ' change the cancel order ID on it later...
        If Order.CancelOrderID <> 0 Then
            Set OtherOrder = New cPtOrder
            If OtherOrder.Load(Order.CancelOrderID) = False Then
                Set OtherOrder = Nothing
            End If
        Else
            Set OtherOrder = Nothing
        End If
        
        ' Make a copy of the old order, mark it as cancelled, and update the collections...
        Set OldOrder = Order.MakeCopy
        OldOrder.ChangeOrderStatus eTT_OrderStatus_Cancelled
        
        ' Clear out and save the new order...
        Order.OrderID = 0&
        Order.BrokerID = ""
        Order.ExchangeID = ""
        ' DAJ 03/16/2012: Don't set the previous broker ID anymore because we don't want
        ' any fills associated with the previous order to be counted as a fill for this
        ' order like they are for amended partially filled orders...
        Order.PreviousBrokerID = "" 'OldOrder.BrokerID
        Order.GenesisOrderID = NextGenesisOrderID(AccountNumberForID(Order.AccountID))
        Order.OrderDate = Order.BrokerDate(CurrentTime("", Order.Symbol))
        Order.Save
        
        SendOrderIdChangeToOptNav Order.AccountID, Str(OldOrder.OrderID), Str(Order.OrderID)
        
        ' If this was one side of an OCO, change the cancel order ID on the other order now
        ' that we have the new ID...
        If Not OtherOrder Is Nothing Then
            OtherOrder.CancelOrderID = Order.OrderID
            OtherOrder.Save
            OrderCallback OtherOrder
            AddOrder OtherOrder
        End If
        
        ' If any parked orders were triggered by the old order, change them to be triggered by
        ' the new order...
        ChangeNegativeOtos Order, OldOrder.OrderID, bForceSubmitOtos
    End If

    ' If the expiration is supposed to be a day order, but is not supplied a date yet, get the
    ' current session date and assign it...
    If Order.Expiration = -1& Then
        Order.Expiration = Bars.SessionDateForTradeTime(ConvertTimeZone(CurrentTime, "", Bars.Prop(eBARS_ExchangeTimeZoneInf))) * -1&
    End If
    
    Select Case nBroker
'            Case eTT_AccountType_CtgPfg
'                If Not g.CtgPfg Is Nothing Then
'                    g.CtgPfg.SendOrder Order
'                End If
'
'            Case eTT_AccountType_FintecPfg
'                If Not g.FintecPfg Is Nothing Then
'                    g.FintecPfg.SendOrder Order
'                End If
'
'            Case eTT_AccountType_LindWaldock
'                If Not g.LindWaldock Is Nothing Then
'                    g.LindWaldock.SendOrder Order, strPrevGenesisOrderID
'                End If
'
'            Case eTT_AccountType_ManExpress
'                If Not g.ManExpress Is Nothing Then
'                    g.ManExpress.SendOrder Order, strPrevGenesisOrderID
'                End If
'
'            Case eTT_AccountType_PFG
'                If Not g.PFG Is Nothing Then
'                    g.PFG.SendOrder Order
'                End If
            
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                g.Transact.SendOrder Order
            End If
        
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                BrokerObj.SendOrder Order
            End If
    
    End Select
    
    AutoJournal Order
    
    SendOrder = Order.OrderID

gdStopProfile 901

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.SendOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateFillSummaryForAutoExit
'' Description: Create position and fill summary entries for the auto exit
'' Inputs:      Auto Exit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateFillSummaryForAutoExit(ByVal AutoExit As cActiveExit)
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the auto exit account
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    nBroker = AccountTypeForID(AutoExit.AccountID)
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.CreateFillSummaryForAutoExit AutoExit
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.CreateFillSummaryForAutoExit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateFillSummaryForAutoTrade
'' Description: Create position and fill summary entries for the auto exit
'' Inputs:      Auto Exit
'' Returns:     Fill Summary ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateFillSummaryForAutoTrade(ByVal AutoTrade As cAutoTradeItem) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim nBroker As eTT_AccountType      ' Broker for the auto exit account
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    lReturn = 0&
    nBroker = AccountTypeForID(AutoTrade.AccountID)
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        lReturn = BInfo.CreateFillSummaryForAutoTrade(AutoTrade)
    End If
    
    CreateFillSummaryForAutoTrade = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CreateFillSummaryForAutoTrade"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFillSummaryForAutoTrade
'' Description: Remove position and fill summary entries for the auto exit
'' Inputs:      Auto Exit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveFillSummaryForAutoTrade(ByVal AutoTrade As cAutoTradeItem)
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the auto exit account
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    nBroker = AccountTypeForID(AutoTrade.AccountID)
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.RemoveFillSummaryForAutoTrade AutoTrade
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.RemoveFillSummaryForAutoTrade"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SnapshotAccountsForBroker
'' Description: Return the snapshot accounts collection for the given broker
'' Inputs:      Broker
'' Returns:     Accounts Collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SnapshotAccountsForBroker(ByVal nBroker As eTT_AccountType) As cPtAccounts
On Error GoTo ErrSection:

    Dim Accounts As cPtAccounts         ' Accounts collection to return
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Accounts = BInfo.Accounts.MakeCopy
    End If

    Set SnapshotAccountsForBroker = Accounts

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.SnapshotAccountsForBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SnapshotAccounts
'' Description: Return the snapshot accounts for all enabled brokers
'' Inputs:      None
'' Returns:     Accounts Collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SnapshotAccounts() As cPtAccounts
On Error GoTo ErrSection:

    Dim Accounts As cPtAccounts         ' Accounts collection to return
    Dim nBroker As Long                 ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim BrokerAccounts As cPtAccounts   ' Collection of broker accounts
    
    Set Accounts = New cPtAccounts
    
    For nBroker = 1 To kNumBrokers - 1
        If IsBrokerUser(nBroker) Then
            Set BrokerAccounts = SnapshotAccountsForBroker(nBroker)
            For lIndex = 1 To BrokerAccounts.Count
                Accounts.Add BrokerAccounts(lIndex).MakeCopy
            Next lIndex
        End If
    Next nBroker
    
    Set SnapshotAccounts = Accounts

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.SnapshotAccounts"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SnapshotOrdersForBroker
'' Description: Return the snapshot orders collection for the given broker
'' Inputs:      Broker
'' Returns:     Orders Collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SnapshotOrdersForBroker(ByVal nBroker As eTT_AccountType) As cPtOrders
On Error GoTo ErrSection:

    Dim Orders As cPtOrders             ' Orders collection to return
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Orders = BInfo.Orders.MakeCopy
    End If

    Set SnapshotOrdersForBroker = Orders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.SnapshotOrdersForBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SnapshotFillsForBroker
'' Description: Return the snapshot fills collection for the given broker
'' Inputs:      Broker
'' Returns:     Fills Collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SnapshotFillsForBroker(ByVal nBroker As eTT_AccountType) As cPtFills
On Error GoTo ErrSection:

    Dim Fills As cPtFills               ' Fills collection to return
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Fills = BInfo.Fills.MakeCopy
    End If

    Set SnapshotFillsForBroker = Fills

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.SnapshotFillsForBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateAccount
'' Description: Update the given account in the appropriate broker info
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateAccount(Account As cPtAccount)
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    Dim BrokerObj As cBroker            ' Broker object
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the account found in the account types array?
    
    If IsLiveAccount(Account.AccountType) Then
        Set BInfo = BrokerInfo(Account.AccountType)
        If Not BInfo Is Nothing Then
            BInfo.AddAccountObject Account
        End If
    Else
        Set BrokerObj = Broker(Account.AccountType)
        If Not BrokerObj Is Nothing Then
            BrokerObj.UpdateAccount Account
        End If
    End If
    
    bFound = False
    For lIndex = 0 To m.astrAccountTypes.Size - 1
        If Parse(m.astrAccountTypes(lIndex), vbTab, 1) = Str(Account.AccountID) Then
            m.astrAccountTypes(lIndex) = Str(Account.AccountID) & vbTab & Account.AccountNumber & vbTab & Str(Account.AccountType) & vbTab & Account.Name
            bFound = True
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        m.astrAccountTypes.Add Str(Account.AccountID) & vbTab & Account.AccountNumber & vbTab & Str(Account.AccountType) & vbTab & Account.Name
    End If
    
    If Not g.TradingItems Is Nothing Then
        g.TradingItems.AccountUpdated Account.AccountID
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.UpdateAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingOrdersForAutoTrade
'' Description: Are there working orders for the given automated trading item?
'' Inputs:      Automated Trading Item
'' Returns:     Number of working orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingOrdersForAutoTrade(AutoTrade As cAutoTradeItem) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim nBroker As eTT_AccountType      ' Broker for the auto trade item
    Dim BInfo As cBrokerInfo            ' Temporary broker info object

    lReturn = 0&
    nBroker = AccountTypeForID(AutoTrade.AccountID)

    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        lReturn = BInfo.HasWorkingOrdersForAutoTrade(AutoTrade.AutoTradeItemID)
    End If
    
    HasWorkingOrdersForAutoTrade = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.HasWorkingOrdersForAutoTrade"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderPriceHitForAutoTrade
'' Description: Are there working orders for the given automated trading item?
'' Inputs:      Automated Trading Item, Price
'' Returns:     True if Working Orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderPriceHitForAutoTrade(AutoTrade As cAutoTradeItem, ByVal dCurrentPrice As Double) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim nBroker As eTT_AccountType      ' Broker for the auto trade item
    Dim BInfo As cBrokerInfo            ' Temporary broker info object

    bReturn = False
    nBroker = AccountTypeForID(AutoTrade.AccountID)

    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        bReturn = BInfo.OrderPriceHitForAutoTrade(AutoTrade.AutoTradeItemID, dCurrentPrice)
    End If
    
    OrderPriceHitForAutoTrade = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.OrderPriceHitForAutoTrade"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrimaryOrdersForSymbol
'' Description: Make a collection of the primary open orders for a symbol
'' Inputs:      Account, Symbol, Auto Trade ID, Price
'' Returns:     Collection of primary orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrimaryOrdersForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrID As Variant, ByVal lAtID As Long, Optional ByVal dPrice As Double = kNullData) As cGdTree
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol to find
    Dim strAccount As String            ' Account number
    Dim nBroker As eTT_AccountType      ' Account type
    Dim Orders As New cGdTree           ' Collection to return from the function
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    vSymbolOrID = ConvertToTradeSymbol(vSymbolOrID, Int(CurrentTime("", "", True)))
    strSymbol = GetSymbol(vSymbolOrID)
    strAccount = GetAccountNumber(vAccountNumberOrID)
    nBroker = AccountTypeForNumber(strAccount)

    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Orders = BInfo.PrimaryOrdersForSymbol(strAccount, strSymbol, lAtID, dPrice)
    End If
    
    Set PrimaryOrdersForSymbol = Orders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.PrimaryOrdersForSymbol"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TriggeredOrdersForOrder
'' Description: Make a collection of the orders to be triggered by an order
'' Inputs:      Triggering Order
'' Returns:     Collection of Triggered Orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TriggeredOrdersForOrder(ByVal Order As cPtOrder) As cGdTree
On Error GoTo ErrSection:

    Dim Orders As New cGdTree           ' Collection of orders to return
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    Set BInfo = BrokerInfo(Order.Broker)
    If Not BInfo Is Nothing Then
        Set Orders = BInfo.TriggeredOrdersForOrderID(Order.OrderID)
    End If

    Set TriggeredOrdersForOrder = Orders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.TriggeredOrdersForOrder"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TriggeredOrdersForOrderID
'' Description: Make a collection of the orders to be triggered by an order ID
'' Inputs:      Triggering Order ID
'' Returns:     Collection of Triggered Orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TriggeredOrdersForOrderID(ByVal lOrderID As Long) As cGdTree
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order object
    Dim Orders As New cGdTree           ' Collection of orders to return
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    If Order.Load(lOrderID) Then
        Set BInfo = BrokerInfo(Order.Broker)
        If Not BInfo Is Nothing Then
            Set Orders = BInfo.TriggeredOrdersForOrderID(Order.OrderID)
        End If
    End If

    Set TriggeredOrdersForOrderID = Orders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.TriggeredOrdersForOrderID"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrdersForOrder
'' Description: Make a collection of the orders to be cancelled by an order
'' Inputs:      Order
'' Returns:     Collection of Cancel Orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelOrdersForOrder(ByVal Order As cPtOrder) As cGdTree
On Error GoTo ErrSection:

    Dim Orders As New cGdTree           ' Collection of orders to return
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    Set BInfo = BrokerInfo(Order.Broker)
    If Not BInfo Is Nothing Then
        Set Orders = BInfo.CancelOrdersForOrderID(Order.OrderID)
    End If

    Set CancelOrdersForOrder = Orders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CancelOrdersForOrder"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrdersForOrderID
'' Description: Make a collection of the orders to be cancelled by an order ID
'' Inputs:      Order ID
'' Returns:     Collection of Cancel Orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelOrdersForOrderID(ByVal lOrderID As Long) As cGdTree
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order object
    Dim Orders As New cGdTree           ' Collection of orders to return
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    If Order.Load(lOrderID) Then
        Set BInfo = BrokerInfo(Order.Broker)
        If Not BInfo Is Nothing Then
            Set Orders = BInfo.CancelOrdersForOrderID(Order.OrderID)
        End If
    End If

    Set CancelOrdersForOrderID = Orders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CancelOrdersForOrderID"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshFill
'' Description: Tell the appropriate broker info object to refresh the given fill
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshFill(Fill As cPtFill)
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the fill
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    nBroker = AccountTypeForID(Fill.AccountID)
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.RefreshFill Fill
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.RefreshFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshOrder
'' Description: Tell the appropriate broker info object to refresh the given order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshOrder(Order As cPtOrder)
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the fill
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    nBroker = AccountTypeForID(Order.AccountID)
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.RefreshOrder Order
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.RefreshOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOrder
'' Description: Update the given order in the appropriate broker info
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddOrder(Order As cPtOrder)
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the fill
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    nBroker = AccountTypeForID(Order.AccountID)
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.AddOrder Order, False
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.AddOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFill
'' Description: Update the given fill in the appropriate broker info
'' Inputs:      Fill, Recalculate?, Previous AT ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddFill(Fill As cPtFill, Optional ByVal bRecalculateSnapshot As Boolean = True, Optional ByVal lPrevAutoTradeID = kNullData)
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the fill
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    nBroker = AccountTypeForID(Fill.AccountID)
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.AddFill Fill, False, bRecalculateSnapshot, , lPrevAutoTradeID
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.AddFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFill
'' Description: Remove the given fill from the appropriate broker info
'' Inputs:      Fill, Recalculate?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveFill(Fill As cPtFill, Optional ByVal bRecalculateSnapshot As Boolean = True)
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the fill
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    nBroker = AccountTypeForID(Fill.AccountID)
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.RemoveFill Fill, bRecalculateSnapshot
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.RemoveFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrdersForAccount
'' Description: Get the orders for the given account
'' Inputs:      Account, Working Only?
'' Returns:     Orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrdersForAccount(ByVal vAccountIDorNumber As Variant, Optional ByVal bWorkingOnly As Boolean = False) As cPtOrders
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account number
    Dim nBroker As eTT_AccountType      ' Broker for the fill
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    Dim Orders As cPtOrders             ' Collection to return
    
    strAccount = GetAccountNumber(vAccountIDorNumber)
    nBroker = AccountTypeForNumber(strAccount)
    Set Orders = New cPtOrders
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Orders = BInfo.OrdersForAccount(strAccount, bWorkingOnly)
    End If
    
    Set OrdersForAccount = Orders
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.OrdersForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillsForAccount
'' Description: Get the fills for the given account
'' Inputs:      Account
'' Returns:     Fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillsForAccount(ByVal vAccountIDorNumber As Variant) As cPtFills
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account number
    Dim nBroker As eTT_AccountType      ' Broker for the fill
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    Dim Fills As cPtFills               ' Collection to return
    
    strAccount = GetAccountNumber(vAccountIDorNumber)
    nBroker = AccountTypeForNumber(strAccount)
    Set Fills = New cPtFills
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Fills = BInfo.FillsForAccount(strAccount)
    End If
    
    Set FillsForAccount = Fills
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.FillsForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSummariesForAccount
'' Description: Get the fill summaries for the given account
'' Inputs:      Account
'' Returns:     Fill Summaries
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillSummariesForAccount(ByVal vAccountIDorNumber As Variant) As cAccountPositions
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID
    Dim nBroker As eTT_AccountType      ' Broker for the fill
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    Dim FillSumms As cAccountPositions  ' Collection of fill summaries to return
    
    lAccountID = GetAccountID(vAccountIDorNumber)
    nBroker = AccountTypeForID(lAccountID)
    Set FillSumms = New cAccountPositions
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set FillSumms = BInfo.FillSummariesForAccount(lAccountID)
    End If
    
    Set FillSummariesForAccount = FillSumms
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.FillSummariesForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OpenPositionsForAccount
'' Description: Retrieve the open positions for the given account
'' Inputs:      Account, Only Total?
'' Returns:     Open positions for the account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenPositionsForAccount(ByVal vAccountIDorNumber As Variant, Optional ByVal bOnlyTotal As Boolean = False) As cAccountPositions
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID
    Dim nBroker As eTT_AccountType      ' Broker for the fill
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    Dim FillSumms As cAccountPositions  ' Collection of fill summaries to return
    
    lAccountID = GetAccountID(vAccountIDorNumber)
    nBroker = AccountTypeForID(lAccountID)
    Set FillSumms = New cAccountPositions
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set FillSumms = BInfo.OpenPositionsForAccount(lAccountID, bOnlyTotal)
    End If
    
    Set OpenPositionsForAccount = FillSumms
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.OpenPositionsForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RebuildFillSummaryForSymbol
'' Description: Rebuild Fill Summaries for the given information
'' Inputs:      Account, Symbol, Auto Trade ID, Rebuild History?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RebuildFillSummaryForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long, Optional ByVal bRebuildHistory As Boolean = False)
On Error GoTo ErrSection:
    
    Dim lAccountID As Long              ' Account ID
    Dim nBroker As eTT_AccountType      ' Broker for the fill
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    lAccountID = GetAccountID(vAccountNumberOrID)
    nBroker = AccountTypeForID(lAccountID)

    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.RebuildFillSummaryForSymbol vAccountNumberOrID, vSymbolOrSymbolID, lAtID, bRebuildHistory
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.RebuildFillSummaryForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RebuildFillSummaryForAccount
'' Description: Rebuild Fill Summaries for the given account
'' Inputs:      Account, Rebuild History?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RebuildFillSummaryForAccount(ByVal vAccountNumberOrID As Variant, Optional ByVal bRebuildHistory As Boolean = False)
On Error GoTo ErrSection:
    
    Dim lAccountID As Long              ' Account ID
    Dim nBroker As eTT_AccountType      ' Broker for the fill
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    lAccountID = GetAccountID(vAccountNumberOrID)
    nBroker = AccountTypeForID(lAccountID)

    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.RebuildFillSummaryForAccount vAccountNumberOrID, bRebuildHistory
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.RebuildFillSummaryForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateIsSnapshot
'' Description: Determine if the given date is within the snapshot date range
'' Inputs:      Session Date, Broker
'' Returns:     True if Snapshot, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateIsSnapshot(ByVal lSessionDate As Long, ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        bReturn = BInfo.DateIsSnapshot(lSessionDate)
    End If
    
    DateIsSnapshot = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.DateIsSnapshot"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateIsSnapshotForOrder
'' Description: Determine if the order session date is within the snapshot date range
'' Inputs:      Order
'' Returns:     True if Snapshot, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateIsSnapshotForOrder(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    
    Set BInfo = BrokerInfo(Order.Broker)
    If Not BInfo Is Nothing Then
        bReturn = BInfo.DateIsSnapshotForAccount(Order.SessionDate, Order.AccountID)
    End If
    
    DateIsSnapshotForOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.DateIsSnapshotForOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateIsSnapshotForFill
'' Description: Determine if the fill session date is within the snapshot date range
'' Inputs:      Fill
'' Returns:     True if Snapshot, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateIsSnapshotForFill(ByVal Fill As cPtFill) As Boolean
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    
    Set BInfo = BrokerInfo(Fill.Broker)
    If Not BInfo Is Nothing Then
        bReturn = BInfo.DateIsSnapshotForAccount(Fill.SessionDate, Fill.AccountID)
    End If
    
    DateIsSnapshotForFill = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.DateIsSnapshotForFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdatePrice
'' Description: Update the price for open profit for the given fill summary
'' Inputs:      Broker, Fill Summary Key, New Price
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdatePrice(ByVal nBroker As eTT_AccountType, ByVal strFillSummKey As String, ByVal dPrice As Double)
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.UpdatePrice strFillSummKey, dPrice
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.UpdatePrice"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Account
'' Description: Return the account for the given informaiton
'' Inputs:      Account Number or ID
'' Returns:     Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Account(ByVal vAccountNumberOrID As Variant) As cPtAccount
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the account passed in
    Dim BInfo As cBrokerInfo            ' Temporary broker info object
    Dim Acct As cPtAccount              ' Account to return
    
    nBroker = AccountTypeForID(GetAccountID(vAccountNumberOrID))
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Acct = BInfo.Account(vAccountNumberOrID)
    End If
    
    Set Account = Acct

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.Account"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAccountInfo
'' Description: Refresh account information for the account passed in
'' Inputs:      Account Number or ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshAccountInfo(ByVal vAccountNumberOrID As Variant)
On Error GoTo ErrSection:

'    Dim nBroker As eTT_AccountType      ' Broker for the account passed in
'
'    nBroker = AccountTypeForID(GetAccountID(vAccountNumberOrID))
'    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                g.CtgPfg.GetAccountInfo GetAccountNumber(vAccountNumberOrID)
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                g.FintecPfg.GetAccountInfo GetAccountNumber(vAccountNumberOrID)
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                g.LindWaldock.GetAccountInfo GetAccountNumber(vAccountNumberOrID)
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                g.ManExpress.GetAccountInfo GetAccountNumber(vAccountNumberOrID)
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                g.PFG.GetAccountInfo GetAccountNumber(vAccountNumberOrID)
'            End If
'
'    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.RefreshAccountInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowAccountInfoForm
'' Description: Show the account information form for the account passed in
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowAccountInfoForm(ByVal strAccount As String)
On Error GoTo ErrSection:

' 09/02/2014 DAJ: Found that this was only used for PFG and Xpress which have both been
' removed from the project.  By getting rid of this code, we can also remove frmAccountInfo
' and cPfgAccountInfo...
#If 0 Then
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the form found?
    Dim frm As frmAccountInfo           ' Account information form
    
    bFound = False
    For lIndex = 0 To Forms.Count - 1
        If TypeOf Forms(lIndex) Is frmAccountInfo Then
            If Forms(lIndex).Account = strAccount Then
                Forms(lIndex).SetFocus
                bFound = True
                Exit For
            End If
        End If
    Next lIndex
    
    If bFound = False Then
        Set frm = New frmAccountInfo
        frm.ShowMe strAccount
    End If
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ShowAccountInfoForm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateAccountInfoForm
'' Description: Update the account information form for the account passed in
'' Inputs:      Account, Account Info
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateAccountInfoForm(ByVal strAccount As String, ByVal strAccountInfo As String)
On Error GoTo ErrSection:

' 09/02/2014 DAJ: Found that this was only used for PFG and Xpress which have both been
' removed from the project.  By getting rid of this code, we can also remove frmAccountInfo
' and cPfgAccountInfo...
#If 0 Then
    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 0 To Forms.Count - 1
        If TypeOf Forms(lIndex) Is frmAccountInfo Then
            If Forms(lIndex).Account = strAccount Then
                Forms(lIndex).AccountInfo = strAccountInfo
                Exit For
            End If
        End If
    Next lIndex
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.UpdateAccountInfoForm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TickValuePerMove
'' Description: Determine the tick value per move for the given information
'' Inputs:      Account, Symbol, Bars
'' Returns:     Tick Value Per Move
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TickValuePerMove(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, Optional Bars As cGdBars = Nothing) As Double
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the account passed in
    Dim dReturn As Double               ' Return value for the function
    
    If Bars Is Nothing Then
        Set Bars = New cGdBars
        SetBarProperties Bars, vSymbolOrSymbolID
    End If
    
    nBroker = AccountTypeForID(GetAccountID(vAccountNumberOrID))
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                dReturn = g.CtgPfg.TickValuePerMove(vSymbolOrSymbolID, Bars)
'            Else
'                dReturn = Bars.TickValuePerMove
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                dReturn = g.FintecPfg.TickValuePerMove(vSymbolOrSymbolID, Bars)
'            Else
'                dReturn = Bars.TickValuePerMove
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                dReturn = g.PFG.TickValuePerMove(vSymbolOrSymbolID, Bars)
'            Else
'                dReturn = Bars.TickValuePerMove
'            End If
            
        Case Else
            dReturn = Bars.TickValuePerMove
    
    End Select
    
    TickValuePerMove = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.TickValuePerMove"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetContractInfo
'' Description: Get the contract information for the given symbol and broker
'' Inputs:      Broker, Symbol, Show Message?
'' Returns:     True if requested, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetContractInfo(ByVal nBroker As eTT_AccountType, ByVal strSymbol As String, Optional ByVal bShowMessage As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strExchangeList As String       ' Exchange list
    Dim strDefaultExchange As String    ' Default exchange
    Dim BrokerObj As cBroker            ' Broker object
    
    bReturn = False
    
    If (IsIbBroker(nBroker) = True) Or (IsRithmicBroker(nBroker) = True) Then
        Set BrokerObj = Broker(nBroker)
        If Not BrokerObj Is Nothing Then
            If BrokerObj.ConnectionStatus = eGDConnectionStatus_Connected Then
                BrokerObj.GetContractInfo strSymbol, bShowMessage
                bReturn = True
            Else
                Select Case nBroker
                    Case eTT_AccountType_Ideal
                        g.Ideal.ExchangeInfoFromIni strSymbol, strExchangeList, strDefaultExchange
                        UpdateBrokerContractInfo nBroker, strSymbol, strExchangeList, strDefaultExchange
                        
                    Case eTT_AccountType_IntBrokers
                        g.IntBroker.ExchangeInfoFromIni strSymbol, strExchangeList, strDefaultExchange
                        UpdateBrokerContractInfo nBroker, strSymbol, strExchangeList, strDefaultExchange
                
                End Select
            End If
        End If
    End If

    GetContractInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.GetContractInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBrokerContractInfo
'' Description: Update broker contract information on appropriate forms
'' Inputs:      Broker, Symbol, Exchange List, Default Exchange
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateBrokerContractInfo(ByVal nBroker As eTT_AccountType, ByVal strSymbol As String, ByVal strExchangeList As String, ByVal strDefaultExchange As String)
On Error GoTo ErrSection:

    Dim strExchangeInfo As String       ' Exchange information to pass to forms
    Dim lIndex As Long                  ' Index into a for loop
    Dim strTradeSymbol As String        ' Trade Symbol for the symbol on the form
    Dim frm As Form                     ' Form out of the collection

    strExchangeInfo = strDefaultExchange & "|" & strExchangeList

    DebugLog "Begin cBrokerDispatch.UpdateBrokerContractInfo (Count = " & Str(Forms.Count) & ")"
    For lIndex = Forms.Count - 1 To 0 Step -1
        If g.bUnloading Then
            Exit For
        
        ' DAJ 09/06/2013: SetExchanges calls InfBox to clear any InfBoxes that may be visible.  This
        ' also caused DoEvents to be called which can ( and in some cases does ) allow other forms
        ' to go away.  This obviously changes the forms count which can cause "Subscript out of
        ' Range" errors on the next time through the loop...
        ElseIf lIndex < Forms.Count Then
            Set frm = Forms(lIndex)
            
            DebugLog vbTab & "Name = '" & frm.Name & "'; Caption = '" & frm.Caption & "' (Count = " & Str(Forms.Count) & ")"
            If IsFrmChart(frm) Then
                strTradeSymbol = ConvertToTradeSymbol(GetSymbol(frm.SymbolOrSymbolID), CurrentTime)
                
                If strTradeSymbol = strSymbol Then
                    If (AccountTypeForID(frm.TradeAccountID) = nBroker) Then
                        frm.SetExchanges strExchangeList, strDefaultExchange
                    End If
                End If
            ElseIf TypeOf frm Is frmTickDistribution Then
                strTradeSymbol = ConvertToTradeSymbol(GetSymbol(frm.SymbolOrSymbolID), CurrentTime)
                
                If strTradeSymbol = strSymbol Then
                    If (AccountTypeForID(frm.TradeAccountID) = nBroker) Then
                        frm.SetExchanges strExchangeList, strDefaultExchange
                    End If
                End If
            ElseIf TypeOf frm Is frmTTEditOrder Then
                strTradeSymbol = ConvertToTradeSymbol(GetSymbol(frm.Symbol), CurrentTime)
                
                If (frm.Broker = nBroker) And (strTradeSymbol = strSymbol) Then
                    frm.SetExchanges strExchangeList, strDefaultExchange
                End If
            End If
        End If
    Next lIndex
    DebugLog "End cBrokerDispatch.UpdateBrokerContractInfo (Count = " & Str(Forms.Count) & ")"
    
ErrExit:
    Exit Sub
    
ErrSection:
    If frm Is Nothing Then
        DebugLog vbTab & vbTab & "Error: '" & Err.Description & "'; Form = Nothing (Count = " & Str(Forms.Count) & ")"
    Else
        DebugLog vbTab & vbTab & "Error: '" & Err.Description & "'; Name = '" & frm.Name & "'; Caption = '" & frm.Caption & "' (Count = " & Str(Forms.Count) & ")"
    End If
    RaiseError "cBrokerDispatch.UpdateBrokerContractInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolsOnFormsForBroker
'' Description: Get list of symbols on appropriate forms for given broker
'' Inputs:      Broker
'' Returns:     Array of Symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SymbolsOnFormsForBroker(ByVal nBroker As eTT_AccountType) As cGdArray
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrReturn As cGdArray          ' Array to return for the function
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    Dim Group As cActiveTsOrderGroup    ' Active TradeSense Order Group object
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings

    For lIndex = Forms.Count - 1 To 0 Step -1
        If g.bUnloading Then
            Exit For
        Else
            If IsFrmChart(Forms(lIndex)) Then
                If (AccountTypeForID(Forms(lIndex).TradeAccountID) = nBroker) Then
                    astrReturn.Add GetSymbol(Forms(lIndex).SymbolOrSymbolID)
                End If
            ElseIf TypeOf Forms(lIndex) Is frmTickDistribution Then
                If (AccountTypeForID(Forms(lIndex).TradeAccountID) = nBroker) Then
                    astrReturn.Add GetSymbol(Forms(lIndex).SymbolOrSymbolID)
                End If
            ElseIf TypeOf Forms(lIndex) Is frmTTEditOrder Then
                If (Forms(lIndex).Broker = nBroker) Then
                    astrReturn.Add Forms(lIndex).Symbol
                End If
            End If
        End If
    Next lIndex
    
    For lIndex = 1 To g.TradingItems.Count
        Set TradeItem = g.TradingItems(lIndex)
        If TradeItem.Broker = nBroker Then
            astrReturn.Add GetSymbol(TradeItem.PositionSymbolOrID)
        End If
    Next lIndex
    
    For lIndex = 1 To g.TsoGroups.Count
        Set Group = g.TsoGroups(lIndex)
        If Group.Broker = nBroker Then
            astrReturn.Add Group.TradeSymbol
        End If
    Next lIndex
    
    astrReturn.Sort eGdSort_DeleteDuplicates
    
    Set SymbolsOnFormsForBroker = astrReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.SymbolsOnFormsForBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExchangeInfoFromIni
'' Description: Attempt to retrieve the exchange information for the given
''              symbol from the INI file.  If not found in the INI file, return
''              the Interactive Brokers exchange for the given symbol
'' Inputs:      Genesis Symbol, Exchange List (returned), Default Exchange (returned)
'' Returns:     True if found in INI file, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExchangeInfoFromIni(ByVal nBroker As eTT_AccountType, ByVal strGenesisSymbol As String, strExchangeList As String, strDefaultExchange As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    ' Set default return values...
    bReturn = False
    strExchangeList = ""
    strDefaultExchange = ""

    ' Ask the appropriate broker for the exchange information from its INI file...
    Select Case nBroker
        Case eTT_AccountType_Ideal
            If Not g.Ideal Is Nothing Then
                bReturn = g.Ideal.ExchangeInfoFromIni(strGenesisSymbol, strExchangeList, strDefaultExchange)
            End If
        
        Case eTT_AccountType_IntBrokers
            If Not g.IntBroker Is Nothing Then
                bReturn = g.IntBroker.ExchangeInfoFromIni(strGenesisSymbol, strExchangeList, strDefaultExchange)
            End If
            
    End Select
    
    ExchangeInfoFromIni = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ExchangeInfoFromIni"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowExchangeControls
'' Description: Determine whether to show exchange controls on the calling form
'' Inputs:      Account, Symbol
'' Returns:     True if show exchange controls, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowExchangeControls(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim nBroker As eTT_AccountType      ' Broker for the account passed in
    
    bReturn = False
    nBroker = AccountTypeForID(GetAccountID(vAccountNumberOrID))
    
    If (nBroker = eTT_AccountType_IntBrokers) Or (nBroker = eTT_AccountType_Ideal) Then
        bReturn = True
    End If
    
    ShowExchangeControls = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ShowExchangeControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionForSymbol
'' Description: Retrieve the position for the given account and symbol
'' Inputs:      Account, Symbol
'' Returns:     Position for that account and symbol
''
'' Fields:      Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant) As String
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Temporary broker object
    Dim strAccountNumber As String      ' Account number
    Dim strSymbol As String             ' Symbol
    Dim nBroker As eTT_AccountType      ' Broker
    Dim strReturn As String             ' Return value for the function
    
    strAccountNumber = GetAccountNumber(vAccountNumberOrID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    nBroker = AccountTypeForNumber(strAccountNumber)
    strReturn = ""
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        strReturn = BInfo.PositionForSymbol(strAccountNumber, strSymbol)
    End If
    
    PositionForSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.PositionForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SnapshotFillsForSymbol
'' Description: Retrieve the snapshot fills for the given account and symbol
'' Inputs:      Account, Symbol
'' Returns:     Snapshot Fills for that account and symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SnapshotFillsForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant) As cPtFills
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Temporary broker object
    Dim strAccountNumber As String      ' Account number
    Dim strSymbol As String             ' Symbol
    Dim nBroker As eTT_AccountType      ' Broker
    Dim Fills As cPtFills               ' Fills to return from the function
    
    strAccountNumber = GetAccountNumber(vAccountNumberOrID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    nBroker = AccountTypeForNumber(strAccountNumber)
    Set Fills = Nothing
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Fills = BInfo.SnapshotFillsForSymbol(strAccountNumber, strSymbol)
    End If
    
    Set SnapshotFillsForSymbol = Fills

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.SnapshotFillsForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OpenPositionsForBaseSymbol
'' Description: Return any open positions in the given account for the given
''              base symbol
'' Inputs:      Account, Symbol, Exclude Symbol?
'' Returns:     Collection of open positions
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenPositionsForBaseSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long, Optional ByVal bExcludeThisSymbol As Boolean = False) As cAccountPositions
On Error GoTo ErrSection:

    Dim FillSumms As cAccountPositions  ' Collection of fill summaries to return
    Dim lAccountID As Long              ' Account ID for account passed in
    Dim strSymbol As String             ' Symbol for symbol passed in
    Dim lSymbolID As Long               ' Symbol ID for symbol passed in
    Dim vTradeSymbolOrID As Variant     ' Trading symbol or symbol ID
    Dim strBaseSymbol As String         ' Base symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim OpenPos As cAccountPositions    ' Collection of open positions to return
    
    Set OpenPos = New cAccountPositions
    
    lAccountID = GetAccountID(vAccountNumberOrID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    If lSymbolID = 0& Then
        vTradeSymbolOrID = ConvertToTradeSymbol(strSymbol, Int(CurrentTime("", "", True)))
    Else
        vTradeSymbolOrID = ConvertToTradeSymbol(lSymbolID, Int(CurrentTime("", "", True)))
    End If
    strBaseSymbol = BaseSymbolForSymbol(vTradeSymbolOrID)
    
    Set FillSumms = FillSummariesForAccount(vAccountNumberOrID)
    If Not FillSumms Is Nothing Then
        For lIndex = 1 To FillSumms.Count
            If (FillSumms(lIndex).AccountID = lAccountID) And (BaseSymbolForSymbol(FillSumms(lIndex).SymbolOrSymbolID) = strBaseSymbol) And (FillSumms(lIndex).AutoTradeItemID = lAtID) Then
                If SecurityType(FillSumms(lIndex).SymbolOrSymbolID, True) = SecurityType(vSymbolOrSymbolID, True) Then
                    If (FillSumms(lIndex).SymbolOrSymbolID <> vTradeSymbolOrID) Or (bExcludeThisSymbol = False) Then
                        If FillSumms(lIndex).CurrentPositionSnapshot <> 0& Then
                            OpenPos.Add FillSumms(lIndex)
                        End If
                    End If
                End If
            End If
        Next lIndex
    End If

    Set OpenPositionsForBaseSymbol = OpenPos
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.OpenPositionsForBaseSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    WorkingOrdersForBaseSymbol
'' Description: Return any working orders in the given account for the given
''              base symbol
'' Inputs:      Account, Symbol, Exclude Symbol?
'' Returns:     Collection of Working Orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WorkingOrdersForBaseSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long, Optional ByVal bExcludeThisSymbol As Boolean = False) As cPtOrders
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID for account passed in
    Dim strSymbol As String             ' Symbol for symbol passed in
    Dim lSymbolID As Long               ' Symbol ID for symbol passed in
    Dim vTradeSymbolOrID As Variant     ' Trading symbol or symbol ID
    Dim strBaseSymbol As String         ' Base symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim Orders As cPtOrders             ' Collection of working orders for the account
    Dim RetOrders As cPtOrders          ' Collection of orders to return
    
    Set RetOrders = New cPtOrders
    
    lAccountID = GetAccountID(vAccountNumberOrID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    If lSymbolID = 0& Then
        vTradeSymbolOrID = ConvertToTradeSymbol(strSymbol, Int(CurrentTime("", "", True)))
    Else
        vTradeSymbolOrID = ConvertToTradeSymbol(lSymbolID, Int(CurrentTime("", "", True)))
    End If
    strBaseSymbol = BaseSymbolForSymbol(vTradeSymbolOrID)
    
    Set Orders = OrdersForAccount(vAccountNumberOrID, True)
    If Not Orders Is Nothing Then
        For lIndex = 1 To Orders.Count
            If (Orders(lIndex).AccountID = lAccountID) And (BaseSymbolForSymbol(Orders(lIndex).SymbolOrSymbolID) = strBaseSymbol) And (Orders(lIndex).AutoTradeItemID = lAtID) Then
                If SecurityType(Orders(lIndex).SymbolOrSymbolID, True) = SecurityType(vSymbolOrSymbolID, True) Then
                    If (Orders(lIndex).SymbolOrSymbolID <> vTradeSymbolOrID) Or (bExcludeThisSymbol = False) Then
                        RetOrders.Add Orders(lIndex)
                    End If
                End If
            End If
        Next lIndex
    End If

    Set WorkingOrdersForBaseSymbol = RetOrders
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.WorkingOrdersForBaseSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderForGenesisID
'' Description: Retrieve the order with the given Genesis Order ID and account
'' Inputs:      Account, Genesis Order ID
'' Returns:     Order if found, Nothing if not
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderForGenesisID(ByVal vAccountNumberOrID As Variant, ByVal strGenesisOrderID As String) As cPtOrder
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to return for the function
    Dim nBroker As eTT_AccountType      ' Broker for the order
    Dim BInfo As cBrokerInfo            ' Temporary broker object
    
    Set Order = Nothing
    nBroker = AccountTypeForID(GetAccountID(vAccountNumberOrID))
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Order = BInfo.OrderForGenesisID(vAccountNumberOrID, strGenesisOrderID)
    End If
    
    Set OrderForGenesisID = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.OrderForGenesisID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderForAccessID
'' Description: Retrieve the order with the given Access ID and account
'' Inputs:      Account, Access ID
'' Returns:     Order if found, Nothing if not
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderForAccessID(ByVal vAccountNumberOrID As Variant, ByVal lAccessID As Long) As cPtOrder
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to return for the function
    Dim nBroker As eTT_AccountType      ' Broker for the order
    Dim BInfo As cBrokerInfo            ' Temporary broker object
    
    Set Order = Nothing
    nBroker = AccountTypeForID(GetAccountID(vAccountNumberOrID))
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Order = BInfo.OrderForAccessID(vAccountNumberOrID, lAccessID)
    End If
    
    Set OrderForAccessID = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.OrderForAccessID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderForBrokerID
'' Description: Retrieve the order with the given Broker Order ID and account
'' Inputs:      Account, Broker ID, Genesis Order ID
'' Returns:     Order if found, Nothing if not
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderForBrokerID(ByVal vAccountNumberOrID As Variant, ByVal strBrokerID As String, Optional ByVal strGenesisOrderID As String = "") As cPtOrder
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to return for the function
    Dim nBroker As eTT_AccountType      ' Broker for the order
    Dim BInfo As cBrokerInfo            ' Temporary broker object
    
    Set Order = Nothing
    nBroker = AccountTypeForID(GetAccountID(vAccountNumberOrID))
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set Order = BInfo.Order(strBrokerID, strGenesisOrderID)
    End If
    
    Set OrderForBrokerID = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.OrderForBrokerID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisIdForOrderId
'' Description: Determine the Genesis Order ID for the given database ID
'' Inputs:      Order ID
'' Returns:     Genesis ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisIdForOrderId(ByVal lOrderID As Long) As String
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strReturn As String             ' Return value for the function
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [OrderID]=" & Str(lOrderID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        strReturn = rs!GenesisOrderID
    End If
    
    GenesisIdForOrderId = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.GenesisIdForOrderId"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTradeLine
'' Description: Get the trade line for the given account, account position,
''              and trade number
'' Inputs:      Account, Account Position ID, Trade Number
'' Returns:     Trade Line (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetTradeLine(ByVal lAccountID As Long, ByVal lAccountPositionID As Long, ByVal lTradeNumber As Long) As cTradeLine
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the given account
    Dim BInfo As cBrokerInfo            ' Broker info object
    Dim TradeLine As cTradeLine         ' Trade line object to return

    nBroker = AccountTypeForID(lAccountID)
    Set TradeLine = Nothing
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        Set TradeLine = BInfo.GetTradeLine(lAccountPositionID, lTradeNumber)
    End If
    
    Set GetTradeLine = TradeLine
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.GetTradeLine"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshTradeLine
'' Description: Refresh the given trade line in the correct fill summary
'' Inputs:      Trade Line
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshTradeLine(ByVal TradeLine As cTradeLine)
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Temporary broker object

    Set BInfo = BrokerInfo(TradeLine.Broker)
    If Not BInfo Is Nothing Then
        BInfo.RefreshTradeLine TradeLine
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.RefreshTradeLine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingOcoOrders
'' Description: Are there working OCO orders for the given broker?
'' Inputs:      Broker
'' Returns:     True if Working OCO Orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingOcoOrders(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Temporary broker object
    Dim bReturn As Boolean              ' Return value from the function

    bReturn = False
    
    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        bReturn = BInfo.HasWorkingOcoOrders
    End If
    
    HasWorkingOcoOrders = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.HasWorkingOcoOrders"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelWorkingOcoOrders
'' Description: Cancel all working OCO orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelWorkingOcoOrders(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Temporary broker object

    Set BInfo = BrokerInfo(nBroker)
    If Not BInfo Is Nothing Then
        BInfo.CancelWorkingOcoOrders
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.CancelWorkingOcoOrders"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerAllowsOCO
'' Description: Does the given broker allow OCO links held at the broker?
'' Inputs:      Broker
'' Returns:     True if Broker Held OCO allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerAllowsOCO(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim BrokerObj As cBroker            ' Broker object
    
    bReturn = False
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            bReturn = False
'
'        Case eTT_AccountType_FintecPfg
'            bReturn = False
'
'        Case eTT_AccountType_LindWaldock
'            bReturn = False
'
'        Case eTT_AccountType_ManExpress
'            bReturn = False
'
'        Case eTT_AccountType_PFG
'            bReturn = True
            
        Case eTT_AccountType_TransAct
            bReturn = False
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                bReturn = BrokerObj.AllowBrokerOco
            End If
            
    End Select
    
    BrokerAllowsOCO = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.BrokerAllowsOCO"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutomaticCommission
'' Description: Determine automatic commission for account/symbol
'' Inputs:      Account, Symbol, Quantity
'' Returns:     Default Commission
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AutomaticCommission(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, Optional ByVal lQuantity As Long = 1&) As Double
On Error GoTo ErrSection:

    Dim Acct As cPtAccount              ' Account object
    Dim dReturn As Double               ' Return value for the function
    
    dReturn = 0#
    Set Acct = Account(vAccountNumberOrID)
    If Not Acct Is Nothing Then
        If SecurityType(vSymbolOrSymbolID, True) = "F" Then
            dReturn = Acct.Comms * lQuantity
        ElseIf SecurityType(vSymbolOrSymbolID, True) = "S" Then
            dReturn = Acct.StockFees
        End If
    End If
    
    AutomaticCommission = dReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AutomaticComission"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EntriesForSymbol
'' Description: Get the open entries for given symbol, account, and at id
'' Inputs:      Account, Symbol, Auto Trade Item ID
'' Returns:     Fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EntriesForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long) As cGdTree
On Error GoTo ErrSection:

    Dim Fills As cGdTree                ' Return value for the function
    Dim AcctPos As cAccountPosition     ' Fill summary for the given symbol
    
    Set Fills = Nothing
    Set AcctPos = FillSummary(vAccountNumberOrID, vSymbolOrSymbolID, lAtID)
    If Not AcctPos Is Nothing Then
        Set Fills = AcctPos.Entries
    End If
    
    Set EntriesForSymbol = Fills

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.EntriesForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OnlyAccountForBroker
'' Description: Determine if the given account ID is the only account for the given broker
'' Inputs:      Account Type, Account ID
'' Returns:     True if only account for broker, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OnlyAccountForBroker(ByVal nAccountType As eTT_AccountType, ByVal lAccountID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim Accounts As cPtAccounts         ' Collection of accounts
    Dim lIndex As Long                  ' Index into a for loop

    bReturn = True
    Set Accounts = g.Broker.SnapshotAccountsForBroker(nAccountType)
    If Not Accounts Is Nothing Then
        For lIndex = 1 To Accounts.Count
            If Accounts(lIndex).AccountID <> lAccountID Then
                bReturn = False
                Exit For
            End If
        Next lIndex
    End If
    
    OnlyAccountForBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.OnlyAccountForBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteAccount
'' Description: Attempt to delete the given account
'' Inputs:      Account
'' Returns:     True if deleted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DeleteAccount(ByVal Account As cPtAccount) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strMessage As String            ' Message to display to the user

    bReturn = False
    strMessage = ""
    
    If Not Account Is Nothing Then
        If (g.Broker.IsLiveAccount(Account.AccountType) = True) And (Account.ConnectionStatus = eGDConnectionStatus_Connected) Then
            strMessage = "You cannot delete an account that is currently connected"
        ElseIf Not TradeTrackerForm(Account.AccountID) Is Nothing Then
            strMessage = "You cannot delete an account while it is open in the Trade Tracker window"
        ElseIf (Account.AccountType = eTT_AccountType_SimStream) And (OnlyAccountForBroker(Account.AccountType, Account.AccountID) = True) Then
            strMessage = "You cannot delete the last simulated account"
        End If
        
        If Len(strMessage) > 0 Then
            InfBox strMessage, "!", , "Delete Account Error"
            g.Broker.BrokerDebug Account.AccountType, "Message displayed to the user: '" & strMessage & "'"
        Else
            If InfBox("Deleting '" & Account.Name & "' will also delete any orders and fills that exist for '" & Account.Name & "'.|Are you sure you want to continue?", "?", "Yes|+-No", "Delete Account Confirmation") = "Y" Then
                g.Broker.RemoveAccount Account
                bReturn = True
            End If
        End If
    End If
    
    DeleteAccount = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.DeleteAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveAccount
'' Description: Remove an account from appropriate places
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveAccount(Account As cPtAccount)
On Error GoTo ErrSection:

    Dim BInfo As cBrokerInfo            ' Broker Info object
    Dim lIndex As Long                  ' Index into a for loop

    ' Remove account from the cache array...
    For lIndex = 0 To m.astrAccountTypes.Size - 1
        If Parse(m.astrAccountTypes(lIndex), vbTab, 1) = Str(Account.AccountID) Then
            m.astrAccountTypes.Remove lIndex
            Exit For
        End If
    Next lIndex

    ' Remove any auto-exits linked to the account...
    If Not g.OrderStrategies Is Nothing Then
        g.OrderStrategies.DeleteForAccount Account.AccountID
    End If

    ' Remove any automated trading items linked to the account...
    If Not g.TradingItems Is Nothing Then
        g.TradingItems.DeleteForAccount Account.AccountID
    End If

    ' Remove any TradeSense order groups linked to the account...
    If Not g.TsoGroups Is Nothing Then
        g.TsoGroups.DeleteForAccount Account.AccountID, "User deleting the account"
    End If
    
    ' Remove the account from the database...
    Account.Delete

    ' Remove the account from the appropriate broker info object...
    Set BInfo = BrokerInfo(Account.AccountType)
    If Not BInfo Is Nothing Then
        BInfo.RemoveAccount Account
    End If
    
    ' Remove the account from the dropdown combo boxes...
    RefreshAccount Account.AccountID
    
    ' Send the remove account notification to Option Navigator...
    SendAccountToOptionNav Account, False, , True
    
    ' If the account was connected, call for a refresh...
    If Account.ConnectionStatus = eGDConnectionStatus_Connected Then
        g.Broker.Refresh Account.AccountType, True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.RemoveAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateAccountCache
'' Description: Update the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateAccountCache(Account As cPtAccount)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    ' Update account in the cache array...
    For lIndex = 0 To m.astrAccountTypes.Size - 1
        If Parse(m.astrAccountTypes(lIndex), vbTab, 1) = Str(Account.AccountID) Then
            m.astrAccountTypes(lIndex) = Str(Account.AccountID) & vbTab & Account.AccountNumber & vbTab & Str(Account.AccountType) & vbTab & Account.Name
            Exit For
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.UpdateAccountCache"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsRithmicBroker
'' Description: Does the given broker use Rithmic for back end?
'' Inputs:      Broker
'' Returns:     True if Rithmic Broker, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsRithmicBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean
    
    Select Case nBroker
        Case eTT_AccountType_Rithmic, eTT_AccountType_Optimus, eTT_AccountType_OpVest, eTT_AccountType_Vision, eTT_AccountType_ZenFire, eTT_AccountType_ZanerRithmic, eTT_AccountType_ZanerZenFire, eTT_AccountType_AlpariZenFire
            bReturn = True
            
        Case Else
            bReturn = False
            
    End Select
    
    IsRithmicBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsRithmicBroker"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsRjoBroker
'' Description: Is the given broker an RJ O'Brien broker?
'' Inputs:      Broker
'' Returns:     True if RJ O'Brien Broker, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsRjoBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean
    
    Select Case nBroker
        Case eTT_AccountType_CtgCqg, eTT_AccountType_CtgPats, eTT_AccountType_RjoCqg, eTT_AccountType_RjoPats, eTT_AccountType_RjoHkPats
            bReturn = True
            
        Case Else
            bReturn = False
            
    End Select
    
    IsRjoBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsRjoBroker"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsCqgBroker
'' Description: Is the given broker an CQG broker?
'' Inputs:      Broker
'' Returns:     True if CQG Broker, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsCqgBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean
    
    Select Case nBroker
        Case eTT_AccountType_CQG, eTT_AccountType_AmpCqg, eTT_AccountType_CtgCqg, eTT_AccountType_FptCqg, eTT_AccountType_KnightCqg, eTT_AccountType_RjoCqg, eTT_AccountType_RobbinsCqg, eTT_AccountType_VisionCqg, eTT_AccountType_ZanerCqg
            bReturn = True
            
        Case Else
            bReturn = False
            
    End Select
    
    IsCqgBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsCqgBroker"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsPfgBroker
'' Description: Is the given broker a PFG broker?
'' Inputs:      Broker
'' Returns:     True if PFG Broker, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsPfgBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean
    
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg, eTT_AccountType_FintecPfg, eTT_AccountType_PFG
'            bReturn = True
            
        Case Else
            bReturn = False
            
    End Select
    
    IsPfgBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsPfgBroker"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsIbBroker
'' Description: Is the given broker an Interactive Brokers broker?
'' Inputs:      Broker
'' Returns:     True if Interactive Brokers Broker, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsIbBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean
    
    Select Case nBroker
        Case eTT_AccountType_Ideal, eTT_AccountType_IntBrokers
            bReturn = True
            
        Case Else
            bReturn = False
            
    End Select
    
    IsIbBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsIbBroker"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsCurrenexBroker
'' Description: Is the given broker a Currenex broker?
'' Inputs:      Broker
'' Returns:     True if Currenex Broker, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsCurrenexBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean
    
    Select Case nBroker
        Case eTT_AccountType_Currenex, eTT_AccountType_AlpariCurrenex, eTT_AccountType_FxddCurrenex, eTT_AccountType_KnightCurrenex, eTT_AccountType_VanKarCurrenex, eTT_AccountType_ZanerCurrenex
            bReturn = True
            
        Case Else
            bReturn = False
            
    End Select
    
    IsCurrenexBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsCurrenexBroker"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsRcgPatsBroker
'' Description: Is the given broker a Rosenthal Collins (PATS) broker?
'' Inputs:      Broker
'' Returns:     True if RCG (PATS) Broker, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsRcgPatsBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean
    
    Select Case nBroker
        Case eTT_AccountType_RcgPats, 16 ' , eTT_AccountType_Rosenthal
            bReturn = True
            
        Case Else
            bReturn = False
            
    End Select
    
    IsRcgPatsBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsRcgPatsBroker"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsPatsBroker
'' Description: Is the given broker a PATS broker?
'' Inputs:      Broker
'' Returns:     True if PATS Broker, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsPatsBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean
    
    Select Case nBroker
        Case eTT_AccountType_AlpariPats, eTT_AccountType_BornPats, eTT_AccountType_CtgPats, eTT_AccountType_DemoPats, eTT_AccountType_PATS, eTT_AccountType_RcgPats, eTT_AccountType_RjoHkPats, eTT_AccountType_RjoPats, eTT_AccountType_ZanerPats
            bReturn = True
            
        Case Else
            bReturn = False
            
    End Select
    
    IsPatsBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsPatsBroker"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOecBroker
'' Description: Is the given broker an Open E-Cry broker?
'' Inputs:      Broker
'' Returns:     True if OEC Broker, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsOecBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean
    
    Select Case nBroker
        Case eTT_AccountType_Oec, eTT_AccountType_FptOec
            bReturn = True
            
        Case Else
            bReturn = False
            
    End Select
    
    IsOecBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsOecBroker"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartingGenesisIdForOptNav
'' Description: Return the starting Genesis ID for Option Navigator for the given broker
'' Inputs:      Broker
'' Returns:     Starting Genesis ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StartingGenesisIdForOptNav(ByVal nBroker As eTT_AccountType) As String
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker object
    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                strReturn = g.CtgPfg.StartingGenesisIdForOptNav
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                strReturn = g.FintecPfg.StartingGenesisIdForOptNav
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                strReturn = g.LindWaldock.StartingGenesisIdForOptNav
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                strReturn = g.ManExpress.StartingGenesisIdForOptNav
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                strReturn = g.PFG.StartingGenesisIdForOptNav
'            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                strReturn = BrokerObj.StartingGenesisIdForOptNav
            End If
    
    End Select
    
    StartingGenesisIdForOptNav = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.StartingGenesisIdForOptNav"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowAccountOnline
'' Description: Allow the user to view their account through the broker website
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowAccountOnline(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

'    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                g.CtgPfg.ShowAccountOnline
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                g.FintecPfg.ShowAccountOnline
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                g.LindWaldock.ShowAccountOnline
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                g.ManExpress.ShowAccountOnline
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                g.PFG.ShowAccountOnline
'            End If
'
'    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ShowAccountOnline"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeBrokerPassword
'' Description: Allow the user to change the password for the broker if supported
'' Inputs:      Broker Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeBrokerPassword(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Select Case nBroker
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                g.LindWaldock.ChangePassword
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                g.ManExpress.ChangePassword
'            End If

    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ChangeBrokerPassword"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTradingMenu
'' Description: Load up the trading menu for the main toolbar
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadTradingMenu()
On Error GoTo ErrSection:

    Dim bSaveRedraw As Boolean          ' Current state of the toolbar's redraw
    Dim lIndex As Long                  ' Index into a for loop

    With frmMain.tbToolbar
        bSaveRedraw = .Redraw
        .Redraw = False
        
        With .Tools("ID_Trading").Menu
            If .Tools.Count = 7 Then
                For lIndex = 1 To kNumBrokers - 1
                    .Tools.Add "ID_Trading_" & Str(lIndex) & "_-1", ssTypeMenu
                    .Tools(lIndex + 7).Name = BrokerName(lIndex)
                Next lIndex
                
                .Tools.Add "separator", ssTypeSeparator, 8
            End If
        
            If .Tools.Count > 8 Then
                For lIndex = 1 To kNumBrokers - 1
                    frmMain.tbToolbar.Tools("ID_Trading_" & Str(lIndex) & "_-1").Visible = IsBrokerUser(lIndex)
                Next lIndex
            End If
        End With
        
        .Redraw = bSaveRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.LoadTradingMenu"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleTradingMenu
'' Description: Handle the trading menu for the main toolbar
'' Inputs:      Tool ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleTradingMenu(ByVal strToolID As String)
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker
    Dim nCommand As eGDTradingMenu      ' Trading menu command
    
    If Len(strToolID) > 11 Then
        If UCase(Left(strToolID, 11)) = "ID_TRADING_" Then
            nBroker = CLng(Val(Parse(strToolID, "_", 3)))
            nCommand = CLng(Val(Parse(strToolID, "_", 4)))
            
            Select Case nCommand
                Case -1&:
                    LoadTradingBrokerMenu nBroker, strToolID
                Case eGDTradingMenu_Connect
                    Connect nBroker
                Case eGDTradingMenu_Disconnect
                    Disconnect nBroker, "User Disconnecting from Main Menu"
                Case eGDTradingMenu_SwitchAccounts
                    SwitchBrokerAccount nBroker
                Case eGDTradingMenu_SwitchAccountsMode
                Case eGDTradingMenu_ConnectInfo
                    ShowBrokerConnectionInfo nBroker
                Case eGDTradingMenu_ChangePassword
                    ChangeBrokerPassword nBroker
                Case eGDTradingMenu_Refresh
                    Refresh nBroker, True
                Case eGDTradingMenu_ViewActivity
                    ShowActivityView nBroker
                Case eGDTradingMenu_BrokerView
                    ShowBrokerView nBroker
                Case eGDTradingMenu_ViewOnline
                    ShowAccountOnline nBroker
                Case eGDTradingMenu_VerifyPositions
                    VerifyPositions nBroker
                Case eGDTradingMenu_AccountDetails
            End Select
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.HandleTradingMenu"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTradingBrokerMenu
'' Description: Load up the trading menu for the given broker
'' Inputs:      Broker Type, Tool ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadTradingBrokerMenu(ByVal nBroker As eTT_AccountType, ByVal strToolID As String)
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker object for the given broker
    Dim nStatus As eGDConnectionStatus  ' Connection status for the broker
    Dim lIndex As eGDTradingMenu        ' Index into a for loop
    Dim bSaveRedraw As Boolean          ' Current state of the toolbar's redraw
    Dim strToolPrefix As String         ' Tool ID prefix
    Dim strID As String                 ' ID for the new tool
    
    Set BrokerObj = Broker(nBroker)
    nStatus = ConnectionStatusForBroker(nBroker)
    
    With frmMain.tbToolbar
        bSaveRedraw = .Redraw
        .Redraw = False
        
        strToolPrefix = "ID_Trading_" & Str(nBroker) & "_"
        
        With .Tools(strToolID).Menu
            If .Tools.Count = 0 Then
                For lIndex = 0 To eGDTradingMenu_NumItems - 1
                    .Tools.Add strToolPrefix & Str(lIndex), ssTypeButton
                    .Tools(lIndex + 1).Name = TradingMenuItemName(nBroker, lIndex)
                Next lIndex
            End If
        End With
    
        For lIndex = 0 To eGDTradingMenu_NumItems - 1
            strID = strToolPrefix & Str(lIndex)
            .Tools(strID).TagVariant = nBroker
            .Tools(strID).Enabled = EnableTradingMenuItem(nBroker, lIndex)
            
            If (lIndex = eGDTradingMenu_SwitchAccountsMode) Or (lIndex = eGDTradingMenu_AccountDetails) Then
                .Tools(strID).Visible = False
            Else
                If BrokerObj Is Nothing Then
                    .Tools(strID).Visible = ShowTradingMenuItem(nBroker, lIndex)
                Else
                    .Tools(strID).Visible = BrokerObj.ShowTradeMenuItem(lIndex)
                End If
            End If
        Next lIndex
        
        .Redraw = bSaveRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.LoadTradingBrokerMenu"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradingMenuItemName
'' Description: Determine the name for the given trading menu item for the broker
'' Inputs:      Broker Type, Menu Item
'' Returns:     Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TradingMenuItemName(ByVal nBroker As eTT_AccountType, ByVal nMenuItem As eGDTradingMenu) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strBrokerName As String         ' Broker name
    
    strBrokerName = BrokerName(nBroker)
    
    Select Case nMenuItem
        Case eGDTradingMenu_Connect
            strReturn = "Connect to " & strBrokerName
        Case eGDTradingMenu_Disconnect
            strReturn = "Disconnect from " & strBrokerName
        Case eGDTradingMenu_SwitchAccounts
            strReturn = "Switch " & strBrokerName & " Account"
        Case eGDTradingMenu_SwitchAccountsMode
            strReturn = "Switch " & strBrokerName & " Account Mode"
            If nBroker = eTT_AccountType_TransAct Then
                Select Case g.Transact.LoginMode
                    Case eGDTransActLoginMode_Live
                        strReturn = "Switch to Simulated Live Mode"
                    Case eGDTransActLoginMode_SimLive
                        strReturn = "Switch to Live Mode"
                End Select
            End If
        Case eGDTradingMenu_ConnectInfo
            strReturn = strBrokerName & " Connection Information"
        Case eGDTradingMenu_ChangePassword
            strReturn = "Change " & strBrokerName & " Password"
        Case eGDTradingMenu_Refresh
            strReturn = "Refresh " & strBrokerName & " Account"
        Case eGDTradingMenu_ViewActivity
            strReturn = "View " & strBrokerName & " Activity"
        Case eGDTradingMenu_BrokerView
            If (nBroker = eTT_AccountType_CtgPats) Or (nBroker = eTT_AccountType_RjoPats) Or (nBroker = eTT_AccountType_TT) Then
                strReturn = "Show " & strBrokerName & " Data"
            Else
                strReturn = strBrokerName & " Broker View"
            End If
        Case eGDTradingMenu_ViewOnline
            strReturn = "View " & strBrokerName & " Account Online"
        Case eGDTradingMenu_VerifyPositions
            strReturn = "Verify " & strBrokerName & " Positions"
        Case eGDTradingMenu_AccountDetails
            strReturn = "View " & strBrokerName & " Account Details"
    End Select
    
    TradingMenuItemName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.TradingMenuItemName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableTradingMenuItem
'' Description: Determine whether to enable the given menu item for the broker
'' Inputs:      Broker Type, Menu Item
'' Returns:     True if Enable, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnableTradingMenuItem(ByVal nBroker As eTT_AccountType, ByVal nMenuItem As eGDTradingMenu) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim nStatus As eGDConnectionStatus  ' Connection status for the given broker
    
    bReturn = False
    nStatus = ConnectionStatusForBroker(nBroker)
    
    Select Case nMenuItem
        Case eGDTradingMenu_Connect
            bReturn = (nStatus <> eGDConnectionStatus_Connected)
        Case eGDTradingMenu_Disconnect
            bReturn = (nStatus <> eGDConnectionStatus_Disconnected)
        Case eGDTradingMenu_SwitchAccounts
            bReturn = (nStatus = eGDConnectionStatus_Connected)
        Case eGDTradingMenu_SwitchAccountsMode
            bReturn = (nStatus = eGDConnectionStatus_Connected)
        Case eGDTradingMenu_ConnectInfo
            bReturn = True
        Case eGDTradingMenu_ChangePassword
            bReturn = (nStatus = eGDConnectionStatus_Connected)
        Case eGDTradingMenu_Refresh
            bReturn = (nStatus = eGDConnectionStatus_Connected)
        Case eGDTradingMenu_ViewActivity
            bReturn = True
        Case eGDTradingMenu_BrokerView
            bReturn = (nStatus = eGDConnectionStatus_Connected)
        Case eGDTradingMenu_ViewOnline
            bReturn = True
        Case eGDTradingMenu_VerifyPositions
            bReturn = (nStatus = eGDConnectionStatus_Connected)
        Case eGDTradingMenu_AccountDetails
            bReturn = (nStatus = eGDConnectionStatus_Connected)
    End Select
    
    EnableTradingMenuItem = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.EnableTradingMenuItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowTradingMenuItem
'' Description: Determine whether to show the given menu item for the broker
'' Inputs:      Broker Type, Menu Item
'' Returns:     True if Show, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowTradingMenuItem(ByVal nBroker As eTT_AccountType, ByVal nMenuItem As eGDTradingMenu) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim BrokerObj As cBroker            ' Broker object for the given broker
    
    bReturn = True
    
    Select Case nBroker
'        Case eTT_AccountType_LindWaldock, eTT_AccountType_ManExpress
'            Select Case nMenuItem
'                Case eGDTradingMenu_Connect: bReturn = True
'                Case eGDTradingMenu_Disconnect: bReturn = True
'                Case eGDTradingMenu_SwitchAccounts: bReturn = True
'                Case eGDTradingMenu_SwitchAccountsMode: bReturn = False
'                Case eGDTradingMenu_ConnectInfo: bReturn = True
'                Case eGDTradingMenu_ChangePassword: bReturn = True
'                Case eGDTradingMenu_Refresh: bReturn = True
'                Case eGDTradingMenu_ViewActivity: bReturn = True
'                Case eGDTradingMenu_BrokerView: bReturn = False
'                Case eGDTradingMenu_ViewOnline: bReturn = True
'                Case eGDTradingMenu_VerifyPositions: bReturn = True
'                Case eGDTradingMenu_AccountDetails: bReturn = True
'            End Select
        
'        Case eTT_AccountType_CtgPfg, eTT_AccountType_FintecPfg, eTT_AccountType_PFG
'            Select Case nMenuItem
'                Case eGDTradingMenu_Connect: bReturn = True
'                Case eGDTradingMenu_Disconnect: bReturn = True
'                Case eGDTradingMenu_SwitchAccounts: bReturn = True
'                Case eGDTradingMenu_SwitchAccountsMode: bReturn = False
'                Case eGDTradingMenu_ConnectInfo: bReturn = True
'                Case eGDTradingMenu_ChangePassword: bReturn = False
'                Case eGDTradingMenu_Refresh: bReturn = True
'                Case eGDTradingMenu_ViewActivity: bReturn = True
'                Case eGDTradingMenu_BrokerView: bReturn = False
'                Case eGDTradingMenu_ViewOnline: bReturn = True
'                Case eGDTradingMenu_VerifyPositions: bReturn = True
'                Case eGDTradingMenu_AccountDetails: bReturn = True
'            End Select
        
        Case eTT_AccountType_TransAct
            Select Case nMenuItem
                Case eGDTradingMenu_Connect: bReturn = True
                Case eGDTradingMenu_Disconnect: bReturn = True
                Case eGDTradingMenu_SwitchAccounts: bReturn = True
                Case eGDTradingMenu_SwitchAccountsMode:
                    Select Case g.Transact.LoginMode
                        Case eGDTransActLoginMode_Demo
                            bReturn = False
                        Case eGDTransActLoginMode_Live
                            bReturn = True
                        Case eGDTransActLoginMode_SimLive
                            bReturn = True
                    End Select
                Case eGDTradingMenu_ConnectInfo: bReturn = True
                Case eGDTradingMenu_ChangePassword: bReturn = False
                Case eGDTradingMenu_Refresh: bReturn = True
                Case eGDTradingMenu_ViewActivity: bReturn = True
                Case eGDTradingMenu_BrokerView: bReturn = False
                Case eGDTradingMenu_ViewOnline: bReturn = False
                Case eGDTradingMenu_VerifyPositions: bReturn = True
                Case eGDTradingMenu_AccountDetails: bReturn = False
            End Select
        
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                bReturn = BrokerObj.ShowTradeMenuItem(nMenuItem)
            End If
            
    End Select
    
    ShowTradingMenuItem = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ShowTradingMenuItem"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolInformation
'' Description: Get the symbol information for the given symbol
'' Inputs:      Broker, Genesis Symbol
'' Returns:     Symbol Information
''
'' Fields:      Order Type Mask, TIF Mask
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SymbolInformation(ByVal nBroker As eTT_AccountType, ByVal strGenesisSymbol As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim BrokerObj As cBroker            ' Broker object
    
    strReturn = ""
    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                strReturn = g.CtgPfg.SymbolInformation(strGenesisSymbol)
'            End If
'
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                strReturn = g.FintecPfg.SymbolInformation(strGenesisSymbol)
'            End If
'
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                strReturn = g.LindWaldock.SymbolInformation(strGenesisSymbol)
'            End If
'
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                strReturn = g.ManExpress.SymbolInformation(strGenesisSymbol)
'            End If
'
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                strReturn = g.PFG.SymbolInformation(strGenesisSymbol)
'            End If
            
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                strReturn = BrokerObj.SymbolInformation(strGenesisSymbol)
            End If
            
    End Select
    
    SymbolInformation = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.SymbolInformation"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckTradeServerOrders
'' Description: Do a call to check on the orders on the Trade Server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckTradeServerOrders()
On Error GoTo ErrSection:

    If ProcessIsBusy = False Then
        g.SimTradeTs.GetOrders
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.CheckTradeServerOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableCheckStatusMenu
'' Description: Determine whether to enable the Check Status menu item
'' Inputs:      None
'' Returns:     True if Enable, False if Disable
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnableCheckStatusMenu() As Boolean
On Error GoTo ErrSection:

    EnableCheckStatusMenu = g.SimTradeTs.Broker.BrokerInfo.HasWorkingOrders(False)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.EnableCheckStatusMenu"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenExpiredPositions
'' Description: Flatten any simulated positions where the symbol has expired
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FlattenExpiredPositions()
On Error GoTo ErrSection:

    g.SimTradeStream.FlattenExpiredPositions
    g.SimTradeTs.FlattenExpiredPositions
        
    Refresh eTT_AccountType_SimStream
    Refresh eTT_AccountType_SimBroker

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.FlattenExpiredPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenExpiredPosition
'' Description: Flattent the expired position for the given account and symbol
'' Inputs:      Account, Symbol, Expiration Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FlattenExpiredPosition(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, Optional ByVal lExpirationDate As Long = kNullData)
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Determine the account type for the given account
    Dim strSymbol As String             ' Symbol
    
    nBroker = AccountTypeForID(GetAccountID(vAccountNumberOrID))
    strSymbol = GetSymbol(vSymbolOrSymbolID)

    Select Case nBroker
        Case eTT_AccountType_SimStream
            g.SimTradeStream.FlattenExpiredPosition vAccountNumberOrID, strSymbol, -1&, lExpirationDate
            
        Case eTT_AccountType_SimReplay
            g.SimTradeReplay.FlattenExpiredPosition vAccountNumberOrID, strSymbol, -1&, lExpirationDate
            
        Case eTT_AccountType_SimBroker
            g.SimTradeTs.FlattenExpiredPosition vAccountNumberOrID, strSymbol, -1&, lExpirationDate
    
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSimTradeTs.FlattenExpiredPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Allow the user to cancel the given order
'' Inputs:      Order to Cancel, Manual Cancel?, Confirm Cancel?, Called from ID,
''              Cancelling All?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelOrder(Order As cPtOrder, ByVal bManualCancel As Boolean, Optional ByVal bConfirm As Boolean = True, Optional ByVal lCalledFromID As Long = 0&, Optional ByVal bCancellingAll As Boolean = False)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strReturn As String             ' Return from infbox
    Dim nStatus As eTT_OrderStatus      ' Current order status
    Dim OtherOrder As cPtOrder          ' Other order from an OCO situation
    Dim BrokerObj As cBroker            ' Broker object

    strReturn = "Y"
    If (bConfirm = True) And (ConfirmManual = True) Then
        strReturn = InfBox("Are you sure that you want to cancel|" & Order.OrderText & "?|", "?", "+Yes|-No", "Order Cancel Confirmation")
        BrokerDebug Order.Broker, vbTab & "Cancel Confirmation: " & strReturn
        
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
            InfBox "You cannot cancel a market order", "!", , "Order Cancel Error"
            
        ElseIf NotSent(Order.Status) Then
            nStatus = Order.Status
            Order.ChangeOrderStatus eTT_OrderStatus_Cancelled
        
        ElseIf OrderIsPending(Order) Then
'            If (Order.Broker = eTT_AccountType_CtgPfg) And (Len(Order.BrokerID) > 0) Then
'                BrokerDebug Order.Broker, "Calling for single order refresh for " & Order.BrokerID & " because cancel called on pending order"
'                g.CtgPfg.GetSingleOrder Order.BrokerID
'            ElseIf (Order.Broker = eTT_AccountType_FintecPfg) And (Len(Order.BrokerID) > 0) Then
'                BrokerDebug Order.Broker, "Calling for single order refresh for " & Order.BrokerID & " because cancel called on pending order"
'                g.FintecPfg.GetSingleOrder Order.BrokerID
'            ElseIf (Order.Broker = eTT_AccountType_PFG) And (Len(Order.BrokerID) > 0) Then
'                BrokerDebug Order.Broker, "Calling for single order refresh for " & Order.BrokerID & " because cancel called on pending order"
'                g.PFG.GetSingleOrder Order.BrokerID
'            Else
                nStatus = Order.Status
                Order.ChangeOrderStatus eTT_OrderStatus_Cancelled
                GetOrders Order.Broker, Order.AccountID
'            End If
                
        ElseIf ConnectionStatusForAccount(Order.AccountID) <> eGDConnectionStatus_Connected Then
            Order.Message = "Not currently connected to " & BrokerName(Order.Broker) & " account " & AccountNameForID(Order.AccountID)
            ShowNotConnectedError Order.AccountID, Order.Broker, "CancelOrder", True
                
        Else
            If bCancellingAll = False Then
                ' DAJ 08/20/2009: If the user decides to cancel one side of an Order-Cancel-Order,
                ' ask them if they wisk to cancel the other side as well...
                If (bManualCancel = True) And (Order.AutoTradeItemID = 0&) And (Order.IsAutoExit = False) Then
                    If (Order.CancelOrderID <> 0) And (Order.CancelOrderID <> lCalledFromID) Then
                        Set OtherOrder = New cPtOrder
                        If OtherOrder.Load(Order.CancelOrderID) Then
                            If IsOpenOrder(OtherOrder.Status) Then
                                If InfBox("You have chosen to cancel one side of an Order-Cancel-Order.  Would you like to cancel the other one as well?", "?", "+Yes|-No", "Order Cancel Order") = "Y" Then
                                    CancelOrder OtherOrder, False, False, Order.OrderID
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
'                        Err.Raise vbObjectError + 1000, , BrokerName(eTT_AccountType_LindWaldock) & " not Initialized"
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
'                            InfBox "You cannot cancel this order because you are not currently connected to PFG", "!", , "PFG Order"
'                        End If
'                    Else
'                        Order.ChangeOrderStatus nStatus
'                        Err.Raise vbObjectError + 1000, , "PFG not Initialized"
'                    End If
                    
                Case eTT_AccountType_TransAct
                    If Not g.Transact Is Nothing Then
                        If ConnectionStatusForAccount(Order.AccountID) = eGDConnectionStatus_Connected Then
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
                    Set BrokerObj = Broker(Order.Broker)
                    If Not BrokerObj Is Nothing Then
                        BrokerObj.CancelOrder Order
                    End If
                
            End Select
            
            ' If this wasn't an auto exit or automated trading order and the user wants the journal
            ' to pop up, do so...
            If AutoJournalPopUp Then
                If (g.OrderStrategies.OrderExistsInStrategy(Order) = False) And (Order.AutoTradeItemID = 0&) Then
                    frmTTSummary.ShowJournalForOrder Order
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.CancelOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Cancel the order and mark it as parked
'' Inputs:      Order, Manual Park?, Called From ID, Ask user if OCO?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ParkOrder(Order As cPtOrder, ByVal bManualPark As Boolean, Optional ByVal lCalledFromID As Long = 0&, Optional ByVal bAskUserOnOCO As Boolean = True)
On Error GoTo ErrSection:

    Dim nStatus As eTT_OrderStatus      ' Current order status
    Dim OtherOrder As cPtOrder          ' Other side of an OCO
    Dim strReturn As String             ' Return from an InfBox
    Dim BrokerObj As cBroker            ' Broker object
    Dim nBroker As eTT_AccountType      ' Broker type for the account number

    If (Order.OrderType = eTT_OrderType_Market) And (IsLiveAccount(Order.Broker) = True) Then
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
                                    ParkOrder OtherOrder, False, Order.OrderID
                                    
                                Case "B"
                                    Order.CancelOrderID = 0
                                
                                Case "C"
                                    Exit Sub
                            End Select
                        Else
                            ParkOrder OtherOrder, False, Order.OrderID
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
'                                        ParkOrder OtherOrder, False, Order.OrderID
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
'                                ParkOrder OtherOrder, False, Order.OrderID
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
'                        Err.Raise vbObjectError + 1000, , "CTG3 not Initialized"
'                    End If
'
'                Case eTT_AccountType_LindWaldock
'                    If Not g.LindWaldock Is Nothing Then
'                        g.LindWaldock.ParkOrder Order
'                    Else
'                        Order.ChangeOrderStatus nStatus
'                        Err.Raise vbObjectError + 1000, , BrokerName(eTT_AccountType_LindWaldock) & " not Initialized"
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
                    Set BrokerObj = Broker(nBroker)
                    If Not BrokerObj Is Nothing Then
                        BrokerObj.ParkOrder Order
                    Else
                        Order.ChangeOrderStatus nStatus
                        Err.Raise vbObjectError + 1000, , BrokerName(nBroker) & " not Initialized"
                    End If
                            
            End Select
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ParkOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateNewFill
'' Description: Create a new fill
'' Inputs:      Fill, Symbol, Account, Auto Trade ID
'' Returns:     True if user OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateNewFill(Fill As cPtFill, ByVal strSymbol As String, ByVal lAccountID As Long, ByVal lAutoTradeID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If frmTTEditFill.ShowMe(Fill, strSymbol, lAccountID, , lAutoTradeID) Then
        InfBox "Please wait while the historical figures for this account get recalculated", , , "Recalculating History for " & AccountNameForID(Fill.AccountID), True
        
        BrokerDebug AccountTypeForID(Fill.AccountID), "Fill " & Fill.BrokerID & ": '" & mTradeTracker.FillDisplay(Fill) & "' has been manually created", True
        
        If DateIsSnapshot(Fill.SessionDate, Fill.Broker) Then
            Fill.IsSnapshot = True
            Fill.Save
            
            AddFill Fill
        
            ' Call for positions for a simulated account to get everything back in sync...
            If (IsLiveAccount(Fill.Broker) = False) Then
                If (ConnectionStatusForBroker(Fill.Broker) = eGDConnectionStatus_Connected) Then
                    Select Case Fill.Broker
                        Case eTT_AccountType_SimBroker
                            g.SimTradeTs.GetPositionForSymbol Fill.AccountID, Fill.SymbolOrSymbolID
                        Case eTT_AccountType_SimStream
                            g.SimTradeStream.GetPositionForSymbol Fill.AccountID, Fill.SymbolOrSymbolID
                    End Select
                End If
            End If
        Else
            Fill.IsSnapshot = False
            Fill.Save
            
            RefreshManualFill Fill, , (Fill.Fees > 0)
        End If
        
        ' If the fill is a member of an automated trading item, make sure to have that automated
        ' trading item refresh its position to account for this fill...
        If Fill.AutoTradingItemID > 0 Then
            g.TradingItems.RefreshPosition Fill.AutoTradingItemID
        End If
        
        bReturn = True
    End If
    
    CreateNewFill = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CreateNewFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ModifyFill
'' Description: Allow the user to modify an existing fill
'' Inputs:      Fill
'' Returns:     True if user OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ModifyFill(Fill As cPtFill) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lOldAutoTradeID As Long         ' Original automated trading item ID
    Dim bDateIsSnapshot As Boolean      ' Is the fill date considered snapshot?
    Dim dOldFees As Double              ' Previous value for the fill fees
    Dim dOldFillDate As Date            ' Previous value for the fill date
    Dim Order As cPtOrder               ' Order object

    bReturn = False
    
    lOldAutoTradeID = Fill.AutoTradingItemID
    dOldFees = Fill.Fees
    dOldFillDate = Fill.FillDate
    
    If frmTTEditFill.ShowMe(Fill, Fill.Symbol, Fill.AccountID) Then
        InfBox "Please wait while the historical figures for this account get recalculated", , , "Recalculating History for " & AccountNameForID(Fill.AccountID), True
        
        bDateIsSnapshot = DateIsSnapshot(Fill.SessionDate, Fill.Broker)
        
        Fill.IsSnapshot = bDateIsSnapshot
        Fill.Save
        BrokerDebug Fill.Broker, "Fill " & Fill.BrokerID & ": '" & mTradeTracker.FillDisplay(Fill) & "' has been manually edited", True
        
        ' DAJ 07/10/2014: If the fill date changed on a simulated fill that has an order such that the
        ' snapshot flag changed on the fill, make sure that the snapshot flag gets changed appropriately
        ' on the order as well.  Tim ran into a case where he changed a snapshot fill to history, but the
        ' order stayed snapshot which caused the fill to be counted as both history and snapshot...
        If (Fill.FillDate <> dOldFillDate) And (Fill.OrderID > 0) And (IsLiveAccount(Fill.Broker) = False) Then
            Set Order = New cPtOrder
            If Order.Load(Fill.OrderID) Then
                If Order.IsSnapshot <> Fill.IsSnapshot Then
                    If Order.IsSnapshot = True Then
                        BrokerDebug Fill.Broker, "Order '" & Order.OrderText(True, True, True) & "' has been changed from snapshot to history because fill just changed"
                    Else
                        BrokerDebug Fill.Broker, "Order '" & Order.OrderText(True, True, True) & "' has been changed from history to snapshot because fill just changed"
                    End If
                    
                    Order.IsSnapshot = Fill.IsSnapshot
                    Order.Save
                End If
            End If
        End If
        
        If bDateIsSnapshot Then
            If lOldAutoTradeID <> Fill.AutoTradingItemID Then
                AddFill Fill, , lOldAutoTradeID
            Else
                AddFill Fill
            End If
        
            ' Call for positions for a simulated account to get everything back in sync...
            If (IsLiveAccount(Fill.Broker) = False) Then
                If (ConnectionStatusForBroker(Fill.Broker) = eGDConnectionStatus_Connected) Then
                    Select Case Fill.Broker
                        Case eTT_AccountType_SimBroker
                            g.SimTradeTs.GetPositionForSymbol Fill.AccountID, Fill.SymbolOrSymbolID
                        Case eTT_AccountType_SimStream
                            g.SimTradeStream.GetPositionForSymbol Fill.AccountID, Fill.SymbolOrSymbolID
                    End Select
                End If
            End If
        ElseIf lOldAutoTradeID <> Fill.AutoTradingItemID Then
            RefreshManualFill Fill, lOldAutoTradeID, (Fill.Fees <> dOldFees)
        Else
            RefreshManualFill Fill, , (Fill.Fees <> dOldFees)
        End If
        
        ' If the fill is and/or was a member of an automated trading item, make sure to have
        ' that automated trading item refresh its position to account for this fill...
        If Fill.AutoTradingItemID > 0 Then
            g.TradingItems.RefreshPosition Fill.AutoTradingItemID
        End If
        If lOldAutoTradeID > 0 Then
            g.TradingItems.RefreshPosition lOldAutoTradeID
        End If
        
        g.CattleBridge.Broker_Fill g.TnCattle.GenesisFillToTurnkey(Fill)
        
        bReturn = True
    End If
    
    ModifyFill = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ModifyFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteFill
'' Description: Allow the user to delete a fill
'' Inputs:      Fill, Source
'' Returns:     True if user OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DeleteFill(Fill As cPtFill, ByVal strSource As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If InfBox("You are about to delete fill||" & Fill.BrokerID & "|( " & mTradeTracker.FillDisplay(Fill, True, True, False, False, False, False, False) & " )||Are you sure you want to continue?|", "?", "+Yes|-No", "Confirmation") = "Y" Then
        InfBox "Please wait while the historical figures for this account get recalculated", , , "Recalculating History for " & AccountNameForID(Fill.AccountID), True
        
        Fill.Delete "Manual from " & strSource
        
        If DateIsSnapshot(Fill.SessionDate, Fill.Broker) Then
            RemoveFill Fill
        
            ' Call for positions for a simulated account to get everything back in sync...
            If (IsLiveAccount(Fill.Broker) = False) Then
                If (ConnectionStatusForBroker(Fill.Broker) = eGDConnectionStatus_Connected) Then
                    Select Case Fill.Broker
                        Case eTT_AccountType_SimBroker
                            g.SimTradeTs.GetPositionForSymbol Fill.AccountID, Fill.SymbolOrSymbolID
                        Case eTT_AccountType_SimStream
                            g.SimTradeStream.GetPositionForSymbol Fill.AccountID, Fill.SymbolOrSymbolID
                    End Select
                End If
            End If
        Else
            RefreshManualFill Fill, , (Fill.Fees > 0)
        End If
        
        ' If the fill was a member of an automated trading item, make sure to have that automated
        ' trading item refresh its position to account for this fill...
        If Fill.AutoTradingItemID > 0 Then
            g.TradingItems.RefreshPosition Fill.AutoTradingItemID
        End If
        
        bReturn = True
    End If
    
    DeleteFill = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.DeleteFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshManualFill
'' Description: Reload, recalculate, and redisplay as appropriate
'' Inputs:      Fill, Old Automated Trading Item ID, Recalculate Fees?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshManualFill(Fill As cPtFill, Optional ByVal lOldAutoTradeID As Long = -2&, Optional ByVal bRecalcFees As Boolean = True)
On Error GoTo ErrSection:

    Dim FillAccount As cPtAccount       ' Account for the given fill
    
    If bRecalcFees Then
        Set FillAccount = Account(Fill.AccountID)
        If Not FillAccount Is Nothing Then
            FillAccount.RecalculateFees
            FillAccount.Save
        End If
    End If
    
    ' Recalculate the appropriate account position history (this will also cause
    ' this form to reload all information)...
    RebuildFillSummaryForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, Fill.AutoTradingItemID, True
    If lOldAutoTradeID >= 0& Then
        RebuildFillSummaryForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, lOldAutoTradeID, True
    End If
    
    ' If the fill was created in a simulated account and it is connected, then
    ' call for a refresh to get everything in sync...
    If (IsLiveAccount(Fill.Broker) = False) Then
        If (ConnectionStatusForBroker(Fill.Broker) = eGDConnectionStatus_Connected) Then
            Select Case Fill.Broker
                Case eTT_AccountType_SimBroker
                    g.SimTradeTs.GetPositionForSymbol Fill.AccountID, Fill.SymbolOrSymbolID
                Case eTT_AccountType_SimStream
                    g.SimTradeStream.GetPositionForSymbol Fill.AccountID, Fill.SymbolOrSymbolID
            End Select
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.RefreshManualFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle a message coming in from the broker
'' Inputs:      Broker, Message Number, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nBroker As eTT_AccountType, ByVal lMsgNumber As Long, ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim BrokerObj As cBroker            ' Broker object

    Select Case nBroker
'        Case eTT_AccountType_CtgPfg
'            If Not g.CtgPfg Is Nothing Then
'                g.CtgPfg.HandleMessage lMsgNumber, strMessage
'            End If
'        Case eTT_AccountType_FintecPfg
'            If Not g.FintecPfg Is Nothing Then
'                g.FintecPfg.HandleMessage lMsgNumber, strMessage
'            End If
        Case eTT_AccountType_Ideal
            If Not g.Ideal Is Nothing Then
                g.Ideal.HandleMessage lMsgNumber, strMessage
            End If
        Case eTT_AccountType_IntBrokers
            If Not g.IntBroker Is Nothing Then
                g.IntBroker.HandleMessage lMsgNumber, strMessage
            End If
'        Case eTT_AccountType_LindWaldock
'            If Not g.LindWaldock Is Nothing Then
'                g.LindWaldock.HandleMessage lMsgNumber, strMessage
'            End If
'        Case eTT_AccountType_ManExpress
'            If Not g.ManExpress Is Nothing Then
'                g.ManExpress.HandleMessage lMsgNumber, strMessage
'            End If
        Case eTT_AccountType_Optimus
            If Not g.Optimus Is Nothing Then
                g.Optimus.HandleMessage lMsgNumber, strMessage
            End If
        Case eTT_AccountType_OpVest
            If Not g.OpVest Is Nothing Then
                g.OpVest.HandleMessage lMsgNumber, strMessage
            End If
'        Case eTT_AccountType_PFG
'            If Not g.PFG Is Nothing Then
'                g.PFG.HandleMessage lMsgNumber, strMessage
'            End If
        Case eTT_AccountType_Rithmic
            If Not g.Rithmic Is Nothing Then
                g.Rithmic.HandleMessage lMsgNumber, strMessage
            End If
        Case eTT_AccountType_TransAct
            If Not g.Transact Is Nothing Then
                g.Transact.HandleMessage lMsgNumber, strMessage
            End If
        Case eTT_AccountType_Vision
            If Not g.Vision Is Nothing Then
                g.Vision.HandleMessage lMsgNumber, strMessage
            End If
        Case eTT_AccountType_ZenFire
            If Not g.ZenFire Is Nothing Then
                g.ZenFire.HandleMessage lMsgNumber, strMessage
            End If
        Case Else
            Set BrokerObj = Broker(nBroker)
            If Not BrokerObj Is Nothing Then
                BrokerObj.HandleMessage lMsgNumber, strMessage
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.HandleMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoJournal
'' Description: If auto journal is turned on and not an automated order, show journal
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AutoJournal(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    ' If this wasn't an auto exit or automated trading order and the user wants the journal
    ' to pop up, do so...
    If m.bAutoJournalPopUp Then
        ' 10/07/2011 DAJ: At least for now, we don't want to pop up the journal window
        ' for orders coming from Option Navigator because the modal dialog creates some
        ' issues...
        ' 01/13/2011 DAJ: We now want the dialog to come up when coming from Option Nav
        ' if they have saved off an image for us...
        If (Order.GroupID = 0) Or (Len(Order.OptionNavImageFile) > 0) Then
            If (g.OrderStrategies.OrderExistsInStrategy(Order) = False) And (Order.AutoTradeItemID = 0&) And (g.TsoGroups.OrderExistsInGroup(Order) = False) Then
                frmTTSummary.ShowJournalForOrder Order
            End If
        End If
        
    ' If they are not auto journalling, but Option Nav created an image file, delete it...
    ElseIf Len(Order.OptionNavImageFile) > 0 Then
        If FileExist(Order.OptionNavImageFile) Then
            KillFile Order.OptionNavImageFile
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.AutoJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetExpirationDateForSymbol
'' Description: Set the expiration date for the given symbol
'' Inputs:      Symbol, Expiration Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetExpirationDateForSymbol(ByVal vSymbolOrSymbolID As Variant, ByVal lExpirationDate As Long)
On Error GoTo ErrSection:

    Dim vTradeSymbolOrID As Variant     ' Trade symbol or symbol ID
    Dim lIndex As Long                  ' Index into a for loop
    Dim BInfo As cBrokerInfo            ' Broker Info object
    
    vTradeSymbolOrID = TradeSymbolOrID(vSymbolOrSymbolID)
    
    For lIndex = 1 To kNumBrokers - 1
        If IsBrokerUser(lIndex) Then
            Set BInfo = BrokerInfo(lIndex)
            If Not BInfo Is Nothing Then
                BInfo.SetExpirationDateForSymbol vTradeSymbolOrID, lExpirationDate
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.SetExpirationDateForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsBroker
'' Description: Determine if the user is actually a broker
'' Inputs:      Account Type
'' Returns:     True if Broker, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    IsBrokerUser nBroker, bReturn
    
    IsBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.IsBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllAccounts
'' Description: Get a collection of accounts the user can see
'' Inputs:      None
'' Returns:     Accounts
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllAccounts() As cPtAccounts
On Error GoTo ErrSection:

    Dim ReturnAccounts As cPtAccounts   ' Collection of accounts to return from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim BInfo As cBrokerInfo            ' Broker Info object
    Dim BrokerAccounts As cPtAccounts   ' Collection of broker accounts
    
    Set ReturnAccounts = New cPtAccounts
    For lIndex = 1 To kNumBrokers - 1
        If IsBrokerUser(lIndex) Then
            Set BInfo = BrokerInfo(lIndex)
            If Not BInfo Is Nothing Then
                Set BrokerAccounts = BInfo.Accounts
                For lIndex2 = 1 To BrokerAccounts.Count
                    ReturnAccounts.Add BrokerAccounts(lIndex2)
                Next lIndex2
            End If
        End If
    Next lIndex
    
    Set AllAccounts = ReturnAccounts

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AllAccounts"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllOrders
'' Description: Get a collection of working orders the user can see
'' Inputs:      Only Working?
'' Returns:     All Orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllOrders(ByVal bOnlyWorking As Boolean) As cPtOrders
On Error GoTo ErrSection:

    Dim ReturnOrders As cPtOrders       ' Collection of orders to return from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim BInfo As cBrokerInfo            ' Broker Info object
    Dim BrokerOrders As cPtOrders       ' Collection of working orders
    
    Set ReturnOrders = New cPtOrders
    For lIndex = 1 To kNumBrokers - 1
        If IsBrokerUser(lIndex) Then
            Set BInfo = BrokerInfo(lIndex)
            If Not BInfo Is Nothing Then
                Set BrokerOrders = BInfo.Orders
                For lIndex2 = 1 To BrokerOrders.Count
                    If (IsOpenOrder(BrokerOrders(lIndex2).Status) = True) Or (bOnlyWorking = False) Then
                        ReturnOrders.Add BrokerOrders(lIndex2)
                    End If
                Next lIndex2
            End If
        End If
    Next lIndex
    
    Set AllOrders = ReturnOrders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AllOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllFills
'' Description: Get a collection of fills the user can see
'' Inputs:      None
'' Returns:     Fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllFills() As cPtFills
On Error GoTo ErrSection:

    Dim ReturnFills As cPtFills         ' Collection of fills to return from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim BInfo As cBrokerInfo            ' Broker Info object
    Dim BrokerFills As cPtFills         ' Collection of broker fills
    
    Set ReturnFills = New cPtFills
    For lIndex = 1 To kNumBrokers - 1
        If IsBrokerUser(lIndex) Then
            Set BInfo = BrokerInfo(lIndex)
            If Not BInfo Is Nothing Then
                Set BrokerFills = BInfo.Fills
                For lIndex2 = 1 To BrokerFills.Count
                    ReturnFills.Add BrokerFills(lIndex2)
                Next lIndex2
            End If
        End If
    Next lIndex
    
    Set AllFills = ReturnFills

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.AllFills"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveSymbolInfo
'' Description: Save the symbol information for the given symbol
'' Inputs:      Symbol or ID, Expiration Date, Lot Size
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveSymbolInfo(ByVal vSymbolOrSymbolID As Variant, ByVal dExpirationDate As Double, ByVal dLotSize As Double)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lSymbolID As Long               ' Symbol ID for the symbol passed in
    Dim strSymbol As String             ' Symbol for the symbol passed in
    
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    
    If lSymbolID = 0& Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblSymbolInfo] WHERE [Symbol]='" & strSymbol & "';", dbOpenDynaset)
    Else
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblSymbolInfo] WHERE [SymbolID]=" & Str(lSymbolID) & ";", dbOpenDynaset)
    End If
    
    If rs.BOF And rs.EOF Then
        rs.AddNew
    Else
        rs.Edit
    End If
    
    rs!SymbolID = lSymbolID
    rs!Symbol = strSymbol
    rs!ExpirationDate = dExpirationDate
    rs!LotSize = dLotSize
    
    rs.Update

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.SaveSymbolInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSymbolInfo
'' Description: Load the symbol information for the given symbol
'' Inputs:      Symbol or ID, Expiration Date, Lot Size
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadSymbolInfo(ByVal vSymbolOrSymbolID As Variant, dExpirationDate As Double, dLotSize As Double) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    Dim lSymbolID As Long               ' Symbol ID for the symbol passed in
    Dim strSymbol As String             ' Symbol for the symbol passed in
    
    bReturn = False
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    
    If lSymbolID = 0& Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblSymbolInfo] WHERE [Symbol]='" & strSymbol & "';", dbOpenDynaset)
    Else
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblSymbolInfo] WHERE [SymbolID]=" & Str(lSymbolID) & ";", dbOpenDynaset)
    End If
    
    If Not (rs.BOF And rs.EOF) Then
        dExpirationDate = rs!ExpirationDate
        dLotSize = rs!LotSize
        bReturn = True
    End If
    
    LoadSymbolInfo = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.LoadSymbolInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CarriedFillsForBroker
'' Description: Get the carried fills for the given broker
'' Inputs:      Broker
'' Returns:     Carried Fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CarriedFillsForBroker(ByVal nBroker As eTT_AccountType) As cGdTree
On Error GoTo ErrSection:

    Dim Fills As cGdTree                ' Fills to return
    Dim BrokerObj As cBroker            ' Broker object

    If nBroker = eTT_AccountType_TransAct Then
        Set Fills = New cGdTree
    Else
        Set BrokerObj = Broker(nBroker)
        If Not BrokerObj Is Nothing Then
            Set Fills = BrokerObj.CarriedFills.MakeCopy
        Else
            Set Fills = New cGdTree
        End If
    End If
    
    Set CarriedFillsForBroker = Fills
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CarriedFillsForBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CarriedFillsForAccount
'' Description: Get the carried fills for the given account
'' Inputs:      Account
'' Returns:     Carried Fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CarriedFillsForAccount(ByVal vAccountIDorNumber As Variant) As cGdTree
On Error GoTo ErrSection:

    Dim lAccountID As String            ' Account ID
    Dim nBroker As eTT_AccountType      ' Broker for the fill
    Dim Fills As cGdTree                ' Fills to return
    Dim BrokerObj As cBroker            ' Broker object
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object

    lAccountID = GetAccountID(vAccountIDorNumber)
    nBroker = AccountTypeForID(lAccountID)
    
    Set Fills = CarriedFillsForBroker(nBroker)
    For lIndex = Fills.Count To 1 Step -1
        Set Fill = Fills(lIndex)
        If Fill.AccountID <> lAccountID Then
            Fills.Remove lIndex
        End If
    Next lIndex
    
    Set CarriedFillsForAccount = Fills
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.CarriedFillsForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowUnknownSymbolError
'' Description: Show an error that the symbol cannot be converted through the broker
'' Inputs:      Symbol, Broker, Source, Dump to Log?, Broker object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowUnknownSymbolError(ByVal strGenesisSymbol As String, ByVal nBroker As eTT_AccountType, ByVal strSource As String, Optional ByVal bDumpToLog As Boolean = False, Optional BrokerObj As cBroker = Nothing)
On Error GoTo ErrSection:

    Dim strBrokerName As String         ' Broker name for the given broker
    Dim strDisplayMessage As String     ' Message to display to the user
    Dim strLogMessage As String         ' Message to send to the log

    If BrokerObj Is Nothing Then
        strBrokerName = BrokerName(nBroker)
    Else
        strBrokerName = BrokerObj.BrokerName
    End If
    
    strDisplayMessage = UnknownSymbolError(strGenesisSymbol, nBroker, strBrokerName, BrokerObj)

    If bDumpToLog Then
        strLogMessage = "Error shown to user: '" & Trim(Replace(strDisplayMessage, "|", " ")) & "' Source = '" & strSource & "'"
        If BrokerObj Is Nothing Then
            BrokerDebug nBroker, strLogMessage
        Else
            BrokerObj.DumpDebug strLogMessage
        End If
    End If
    InfBox strDisplayMessage, "!", , strBrokerName & " Order Error"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ShowUnknownSymbolError"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnknownSymbolError
'' Description: Build an error that the symbol cannot be converted through the broker
'' Inputs:      Symbol, Broker, Broker object
'' Returns:     Error
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UnknownSymbolError(ByVal strGenesisSymbol As String, ByVal nBroker As eTT_AccountType, ByVal strBrokerName As String, Optional BrokerObj As cBroker = Nothing) As String
On Error GoTo ErrSection:

    UnknownSymbolError = "Trade Navigator is not set up to trade||" & strGenesisSymbol & "||through the " & strBrokerName & " servers|"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.UnknownSymbolError"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowNotEnabledForSymbolError
'' Description: Show an error that the symbol is not enabled through the broker
'' Inputs:      Symbol, Broker, Broker Base Symbol, Broker Exchange, Source,
''              Dump to Log?, Broker object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowNotEnabledForSymbolError(ByVal strGenesisSymbol As String, ByVal nBroker As eTT_AccountType, ByVal strBrokerBase As String, ByVal strBrokerExchange As String, ByVal strSource As String, Optional ByVal bDumpToLog As Boolean = False, Optional BrokerObj As cBroker = Nothing)
On Error GoTo ErrSection:

    Dim strBrokerName As String         ' Broker name for the given broker
    Dim strDisplayMessage As String     ' Message to display to the user
    Dim strLogMessage As String         ' Message to send to the log
    Dim astrEnabledSymbols As cGdArray  ' Enabled symbols
    Dim strButtons As String            ' Buttons to display
    Dim lIndex As Long                  ' Index into a for loop

    If BrokerObj Is Nothing Then
        Set astrEnabledSymbols = EnabledSymbols(nBroker)
        strBrokerName = BrokerName(nBroker)
    Else
        Set astrEnabledSymbols = BrokerObj.EnabledSymbols
        strBrokerName = BrokerObj.BrokerName
    End If
        
    strDisplayMessage = NotEnabledForSymbolError(strGenesisSymbol, nBroker, strBrokerBase, strBrokerExchange, strBrokerName, BrokerObj)

    If bDumpToLog Then
        strLogMessage = "Error shown to user: '" & Trim(Replace(strDisplayMessage, "|", " ")) & "' Source = '" & strSource & "'"
        If BrokerObj Is Nothing Then
            BrokerDebug nBroker, strLogMessage
        Else
            BrokerObj.DumpDebug strLogMessage
        End If
    End If
    
    If astrEnabledSymbols Is Nothing Then
        strButtons = ""
    Else
        strButtons = "+-OK|Symbols"
        
        For lIndex = astrEnabledSymbols.Size - 1 To 0 Step -1
            If (astrEnabledSymbols(lIndex) = "!") Or (astrEnabledSymbols(lIndex) = "@") Or (Left(astrEnabledSymbols(lIndex), 2) = "O:") Then
                astrEnabledSymbols.Remove lIndex
            End If
        Next lIndex
    End If
    
    If InfBox(strDisplayMessage, "!", strButtons, strBrokerName & " Order Error") = "S" Then
        If Not astrEnabledSymbols Is Nothing Then
            InfBox "You are enabled to trade the following futures|through the " & strBrokerName & " servers:||" & astrEnabledSymbols.JoinFields(", ") & "|", , , "Enabled Symbols"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ShowNotEnabledForSymbolError"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NotEnabledForSymbolError
'' Description: Build an error that the symbol is not enabled through the broker
'' Inputs:      Symbol, Broker, Broker Base Symbol, Broker Exchange, Broker object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NotEnabledForSymbolError(ByVal strGenesisSymbol As String, ByVal nBroker As eTT_AccountType, ByVal strBrokerBase As String, ByVal strBrokerExchange As String, ByVal strBrokerName As String, Optional BrokerObj As cBroker = Nothing) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If Len(strBrokerExchange) > 0 Then
        If IsCqgBroker(nBroker) Or IsIbBroker(nBroker) Then
            strBrokerBase = "'" & strBrokerBase & "' on the '" & strBrokerExchange & "' exchange"
        Else
            strBrokerBase = "'" & strBrokerBase & "'|on the '" & strBrokerExchange & "' exchange"
        End If
    Else
        strBrokerBase = "'" & strBrokerBase & "'"
    End If
    
    If IsCqgBroker(nBroker) Or IsIbBroker(nBroker) Then
        strReturn = "Your " & strBrokerName & " account is not enabled for real-time market data for the symbol:||" & strGenesisSymbol & "||Trade Navigator will not allow you to trade this symbol until you call your broker to configure your account to be enabled for real-time market data for the symbol " & strBrokerBase & "|"
    Else
        strReturn = "Your " & strBrokerName & " account is not set up to trade||" & strGenesisSymbol & "||Please call your broker to configure your account to be able to trade the symbol " & strBrokerBase & "|"
    End If
    
    NotEnabledForSymbolError = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.NotEnabledForSymbolError"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowNotConnectedError
'' Description: Show an error that the user is not connected to the account
'' Inputs:      Account, Broker, Source, Dump to Log?, Broker object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowNotConnectedError(ByVal lAccountID As Long, ByVal nBroker As eTT_AccountType, ByVal strSource As String, Optional ByVal bDumpToLog As Boolean = False, Optional BrokerObj As cBroker = Nothing)
On Error GoTo ErrSection:

    Dim strDisplayMessage As String     ' Message to display to the user
    Dim strBrokerName As String         ' Display name for the broker passed in
    Dim strLogMessage As String         ' Message to send to the log

    If BrokerObj Is Nothing Then
        strBrokerName = BrokerName(nBroker)
    Else
        strBrokerName = BrokerObj.BrokerName
    End If
    strDisplayMessage = "You are not currently connected to " & strBrokerName & " account '" & g.Broker.AccountNameForID(lAccountID) & "'"

    If bDumpToLog Then
        strLogMessage = "Error shown to user: '" & Trim(Replace(strDisplayMessage, "|", " ")) & "' Source = '" & strSource & "'"
        If BrokerObj Is Nothing Then
            BrokerDebug nBroker, strLogMessage
        Else
            BrokerObj.DumpDebug strLogMessage
        End If
    End If
    
    InfBox strDisplayMessage, "!", , strBrokerName & " Order Error"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ShowNotConnectedError"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoesLiteralForexQuantities
'' Description: Does the given broker do literal forex quantities?
'' Inputs:      Broker Type
'' Returns:     True if does literal forex quantities, false otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DoesLiteralForexQuantities(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    DoesLiteralForexQuantities = IsCurrenexBroker(nBroker) Or IsIbBroker(nBroker)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.DoesLiteralForexQuantities"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrdersAsOco
'' Description: Submit the given orders as an Order-Cancels-Order
'' Inputs:      Order1, Order2, Triggered-By Order ID, Hold at Broker?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SubmitOrdersAsOco(Order1 As cPtOrder, Order2 As cPtOrder, Optional ByVal lTriggeredByOrderID As Long = 0&, Optional ByVal bHoldAtBroker As Boolean = False)
On Error GoTo ErrSection:

    Dim strAccountNumber As String      ' Account number for the order
    Dim BrokerObj As cBroker            ' Broker object

    ' If Order1 doesn't have a GenesisOrderID yet, assign one...
    If Len(Order1.GenesisOrderID) = 0 Then
        strAccountNumber = g.Broker.AccountNumberForID(Order1.AccountID)
        Order1.GenesisOrderID = NextGenesisOrderID(strAccountNumber, Order1.Broker)
    End If

    ' If Order2 doesn't have a GenesisOrderID yet, assign one...
    If Len(Order2.GenesisOrderID) = 0 Then
        strAccountNumber = g.Broker.AccountNumberForID(Order2.AccountID)
        Order2.GenesisOrderID = NextGenesisOrderID(strAccountNumber, Order2.Broker)
    End If

    ' If Order1 doesn't have an Access ID yet, we need to save it so that we have one...
    If Order1.OrderID <= 0& Then
        Order1.Save
    End If
    
    ' If trying to hold the OCO at the broker, need to verify that we can...
    If bHoldAtBroker Then
        If Order1.Broker <> Order2.Broker Then
            bHoldAtBroker = False
        Else
            bHoldAtBroker = BrokerAllowsOCO(Order1.Broker)
        End If
    End If
    
    If (bHoldAtBroker = True) And (lTriggeredByOrderID = 0&) Then
        Set BrokerObj = Broker(Order1.Broker)
        If Not BrokerObj Is Nothing Then
            BrokerObj.SubmitOrdersAsOco Order1, Order2
        End If
    Else
        Order2.CancelOrderID = Order1.OrderID
        Order2.TriggerOrderID = lTriggeredByOrderID
        Order2.ChangeOrderStatus eTT_OrderStatus_Parked
        
        Order1.CancelOrderID = Order2.OrderID
        Order1.TriggerOrderID = lTriggeredByOrderID
        Order1.ChangeOrderStatus eTT_OrderStatus_Parked
        
        ' SubmitOrder will take care of submitting the other side, but make sure we don't ask the user...
        SubmitOrder Order1, , , , False
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.SubmitOrdersAsOco"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    KillStandalones
'' Description: Kill any standalone processes that may still be running
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub KillStandalones()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strProcessName As String        ' Name of the process to kill
    Dim BrokerObj As cBroker            ' Broker object

    For lIndex = 1 To kNumBrokers - 1
        If IsBrokerUser(lIndex) Then
            strProcessName = ""
            
            Select Case lIndex
'                Case eTT_AccountType_CtgPfg
'                    If Not g.CtgPfg Is Nothing Then
'                        strProcessName = g.CtgPfg.ProcessName
'                    End If
'
'                Case eTT_AccountType_FintecPfg
'                    If Not g.FintecPfg Is Nothing Then
'                        strProcessName = g.FintecPfg.ProcessName
'                    End If
'
'                Case eTT_AccountType_LindWaldock
'                    If Not g.LindWaldock Is Nothing Then
'                        strProcessName = g.LindWaldock.ProcessName
'                    End If
'
'                Case eTT_AccountType_ManExpress
'                    If Not g.ManExpress Is Nothing Then
'                        strProcessName = g.ManExpress.ProcessName
'                    End If
'
'                Case eTT_AccountType_PFG
'                    If Not g.PFG Is Nothing Then
'                        strProcessName = g.PFG.ProcessName
'                    End If
                    
                Case eTT_AccountType_TransAct
                    If Not g.Transact Is Nothing Then
                        strProcessName = g.Transact.ProcessName
                    End If
                    
                Case Else
                    Set BrokerObj = Broker(lIndex)
                    If Not BrokerObj Is Nothing Then
                        strProcessName = BrokerObj.ProcessName
                    End If
                    
            End Select
            
            If Len(strProcessName) > 0 Then
                If KillProcess(strProcessName) > 0 Then
                    StartupLog strProcessName & " killed"
                End If
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.KillStandalones"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MaxOrderLegs
'' Description: Determine the maximum number of order legs allowed on an
''              order based on the given broker
'' Inputs:      Broker Type
'' Returns:     Maximum number of order legs
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MaxOrderLegs(ByVal nBroker As eTT_AccountType) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = 1&
'    If (nBroker = eTT_AccountType_LindWaldock) Or (nBroker = eTT_AccountType_ManExpress) Then
'        lReturn = 3&
    If IsIbBroker(nBroker) Then
        lReturn = 4&
    ElseIf nBroker = eTT_AccountType_SimStream Then
        lReturn = 4&
    ElseIf nBroker = eTT_AccountType_Tradier Then
        lReturn = 4&
    End If
    
    MaxOrderLegs = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.MaxOrderLegs"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SecTypesForBroker
'' Description: Determine the security types allowed for the given broker
'' Inputs:      Broker
'' Returns:     Security Types
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SecTypesForBroker(ByVal nBroker As eTT_AccountType) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = "F,FO"
    If IsOecBroker(nBroker) Then
        strReturn = "F,FO,OECFX"
    ElseIf IsIbBroker(nBroker) Then
        strReturn = "F,FO,S,SO,IBFX"
    ElseIf IsCurrenexBroker(nBroker) Then
        strReturn = "CNXFX"
    ElseIf IsRithmicBroker(nBroker) Then
        strReturn = "F,FO,S,SO"
    ElseIf nBroker = eTT_AccountType_SimStream Then
        strReturn = "F,FO,S,SO,FX,CNXFX,IBFX,OECFX"
    ElseIf nBroker = eTT_AccountType_SimBroker Then
        strReturn = "F,FO,S,SO,FX"
    End If
    
    SecTypesForBroker = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.SecTypesForBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeSecType
'' Description: Determine the security type for the given symbol or symbol id
'' Inputs:      Symbol or Symbol ID
'' Returns:     Security Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TradeSecType(ByVal vSymbolOrSymbolID As Variant) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strSymbol As String             ' Symbol from the value passed in
    
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    strReturn = SecurityType(strSymbol, True)
    If strReturn = "I" Then
        If IsForex(strSymbol) Then
            strReturn = "FX"
            If InStr(strSymbol, "@") <> 0 Then
                strReturn = UCase(Parse(strSymbol, "@", 2) & strReturn)
            End If
        End If
    End If
    
    TradeSecType = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.TradeSecType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MinimumOrderQuantity
'' Description: Determine the minimum order quantity allowed for the given
''              account and symbol
'' Inputs:      Account, Symbol
'' Returns:     Minimum order quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MinimumOrderQuantity(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant) As Long
On Error GoTo ErrSection:

    MinimumOrderQuantity = GetProvidedBrokerProperty(vAccountNumberOrID, vSymbolOrSymbolID, "MinQty", 1&)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.MinimumOrderQuantity"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DefaultOrderQuantity
'' Description: Determine the default order quantity allowed for the given
''              account and symbol
'' Inputs:      Account, Symbol
'' Returns:     Default order quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DefaultOrderQuantity(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant) As Long
On Error GoTo ErrSection:

    Dim strSecType As String            ' Security type for the given symbol
    Dim lDefault As Long                ' Default value for the default quantity
    
    lDefault = 1&
    strSecType = TradeSecType(vSymbolOrSymbolID)
    
    Select Case UCase(strSecType)
        Case "S"
            lDefault = 100&
        Case "CNXFX"
            lDefault = 10000&
        Case "IBFX"
            lDefault = 10000&
    End Select

    DefaultOrderQuantity = GetProvidedBrokerProperty(vAccountNumberOrID, vSymbolOrSymbolID, "DefaultQty", lDefault)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.DefaultOrderQuantity"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MinimumLotSize
'' Description: Determine the minimum lot size allowed for the given account
''              and symbol
'' Inputs:      Account, Symbol
'' Returns:     Minimum lot size
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MinimumLotSize(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant) As Long
On Error GoTo ErrSection:

    MinimumLotSize = GetProvidedBrokerProperty(vAccountNumberOrID, vSymbolOrSymbolID, "MinLotSize", 1&)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.MinimumLotSize"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidQuantity
'' Description: Determine if the given quantity is valid for the given information
'' Inputs:      Account, Symbol, Quantity
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidQuantity(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lQuantity As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lMinQuantity As Long            ' Minimum quantity
    Dim lMinLotSize As Long             ' Minimum lot size

    bReturn = True
    lMinQuantity = MinimumOrderQuantity(vAccountNumberOrID, vSymbolOrSymbolID)
    lMinLotSize = MinimumLotSize(vAccountNumberOrID, vSymbolOrSymbolID)

    If (lQuantity < lMinQuantity) Then
        bReturn = False
    ElseIf ((lQuantity Mod lMinLotSize) <> 0) Then
        bReturn = False
    End If

    ValidQuantity = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ValidQuantity"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidQuantityForOrder
'' Description: Determine if the order quantity is valid
'' Inputs:      Order
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidQuantityForOrder(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lMinQuantity As Long            ' Minimum quantity
    Dim lMinLotSize As Long             ' Minimum lot size
    Dim lCurrentPos As Long             ' Current position for the symbol/account

    bReturn = True
    lMinQuantity = MinimumOrderQuantity(Order.AccountID, Order.SymbolOrSymbolID)
    lMinLotSize = MinimumLotSize(Order.AccountID, Order.SymbolOrSymbolID)
    lCurrentPos = CurrentPosition(Order.AccountID, Order.Symbol, -1&)
    
    If ((Order.Buy = True) And (lCurrentPos >= 0)) Or ((Order.Buy = False) And (lCurrentPos <= 0)) Or (Order.Quantity <> lCurrentPos) Then
        If (Order.Quantity < lMinQuantity) Then
            bReturn = False
        ElseIf ((Order.Quantity Mod lMinLotSize) <> 0) Then
            bReturn = False
        End If
    End If

    ValidQuantityForOrder = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ValidQuantityForOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitQuantityEditor
'' Description: Initialize a price editor control for quantity
'' Inputs:      Price Editor, Scroll Bar, Text Box, Account, Symbol, Quantity,
''              Allow Zero?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'RH changed Textbox
'Public Sub InitQuantityEditor(Qty As cPriceEditor, sbQty As gdScrollBar, txtQty As TextBox, ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, Optional ByVal lQuantity As Long = kNullData, Optional ByVal bAllowZero As Boolean = False)
Public Sub InitQuantityEditor(Qty As cPriceEditor, sbQty As gdScrollBar, txtQty As ctlUniTextBoxXP, ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, Optional ByVal lQuantity As Long = kNullData, Optional ByVal bAllowZero As Boolean = False)
On Error GoTo ErrSection:

    Dim lMinimumQuantity As Long        ' Minimum quantity
    Dim lDefaultQuantity As Long        ' Default quantity
    Dim lMinimumLotSize As Long         ' Minimum lot size
    
    If bAllowZero Then
        lMinimumQuantity = 0&
    Else
        lMinimumQuantity = MinimumOrderQuantity(vAccountNumberOrID, vSymbolOrSymbolID)
    End If
    If lQuantity = kNullData Then
        lDefaultQuantity = DefaultOrderQuantity(vAccountNumberOrID, vSymbolOrSymbolID)
    Else
        lDefaultQuantity = lQuantity
    End If
    lMinimumLotSize = MinimumLotSize(vAccountNumberOrID, vSymbolOrSymbolID)
    
    Qty.Init sbQty, txtQty, Nothing, lDefaultQuantity, lMinimumQuantity, , , , lMinimumLotSize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.InitQuantityEditor"
    Resume ErrExit
    Resume
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetQuantityPresets
'' Description: Get the quantity presets for the given information
'' Inputs:      Account, Symbol, First preset, Second preset, Third preset
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetQuantityPresets(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, lPreset1 As Long, lPreset2 As Long, lPreset3 As Long)
On Error GoTo ErrSection:

    lPreset1 = PresetQuantity(vAccountNumberOrID, vSymbolOrSymbolID, 1&)
    lPreset2 = PresetQuantity(vAccountNumberOrID, vSymbolOrSymbolID, 2&)
    lPreset3 = PresetQuantity(vAccountNumberOrID, vSymbolOrSymbolID, 3&)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.GetQuantityPresets"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetQuantityPresets
'' Description: Set the quantity presets for the given information
'' Inputs:      Account, Symbol, First preset, Second preset, Third preset
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetQuantityPresets(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lPreset1 As Long, ByVal lPreset2 As Long, ByVal lPreset3 As Long)
On Error GoTo ErrSection:

    Dim bChanged As Boolean             ' Did any of the values change?
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSecType As String            ' Security type for the symbol passed in

    bChanged = False

    If PresetQuantity(vAccountNumberOrID, vSymbolOrSymbolID, 1&) <> lPreset1 Then
        PresetQuantity(vAccountNumberOrID, vSymbolOrSymbolID, 1&) = lPreset1
        bChanged = True
    End If
    If PresetQuantity(vAccountNumberOrID, vSymbolOrSymbolID, 2&) <> lPreset2 Then
        PresetQuantity(vAccountNumberOrID, vSymbolOrSymbolID, 2&) = lPreset2
        bChanged = True
    End If
    If PresetQuantity(vAccountNumberOrID, vSymbolOrSymbolID, 3&) <> lPreset3 Then
        PresetQuantity(vAccountNumberOrID, vSymbolOrSymbolID, 3&) = lPreset3
        bChanged = True
    End If
    
    If bChanged Then
        strSecType = TradeSecType(vSymbolOrSymbolID)
        For lIndex = 0 To Forms.Count - 1
            If (TypeOf Forms(lIndex) Is frmChart) Or (TypeOf Forms(lIndex) Is frmChart2) Or (TypeOf Forms(lIndex) Is frmTickDistribution) Then
                If TradeSecType(Forms(lIndex).SymbolOrSymbolID) = strSecType Then
                    Forms(lIndex).SetQuantityPresetButtons
                End If
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.SetQuantityPresets"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ContractSize
'' Description: Determine the override for the contract size if applicable
'' Inputs:      Genesis Symbol, Account
'' Returns:     Contract Size override
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ContractSize(ByVal strGenesisSymbol As String, ByVal vAccountNumberOrID As Variant) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim strAccountNumber As String      ' Account number
    Dim nBroker As eTT_AccountType      ' Account type
    Dim BrokerObj As cBroker            ' Broker object
    
    lReturn = kNullData
    strAccountNumber = GetAccountNumber(vAccountNumberOrID)
    nBroker = AccountTypeForNumber(strAccountNumber)
    
    Set BrokerObj = Broker(nBroker)
    If Not BrokerObj Is Nothing Then
        lReturn = BrokerObj.ContractSize(strGenesisSymbol, strAccountNumber)
    End If
    
    ContractSize = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.ContractSize"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReplayLogFile
'' Description: Replay the given log file for the given broker
'' Inputs:      Log File, Broker Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReplayLogFile(ByVal strLogFile As String, ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' File broken out into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim strReceived As String           ' Message Received string
    Dim iLen As Integer                 ' Length of string
    Dim iPos As Integer                 ' Position of string in another string
    Dim iPos2 As Integer                ' Position of string in another string
    Dim strBroker As String             ' Broker name
    Dim strMessageType As String        ' Message type
    Dim lMsgNumber As Long              ' Message number
    Dim strMessage As String            ' Message to send
    Dim BrokerObj As cBroker            ' Broker object
    
    strBroker = g.Broker.BrokerName(nBroker)
    strReceived = strBroker & " Message Received"
    iLen = Len(strReceived)
    
    Set astrFile = New cGdArray
    If astrFile.FromFile(strLogFile) Then
        For lIndex = 0 To astrFile.Size - 1
            iPos = InStr(astrFile(lIndex), strReceived)
            If iPos > 0 Then
                iPos = iPos + iLen + 2
                iPos2 = InStr(iPos, astrFile(lIndex), ")")
                If iPos2 > 0 Then
                    strMessageType = Mid(astrFile(lIndex), iPos, iPos2 - iPos)
                    strMessage = Mid(astrFile(lIndex), iPos2 + 3)
                    
                    Select Case nBroker
                        Case eTT_AccountType_Ideal
                            If Not g.Ideal Is Nothing Then
                                lMsgNumber = g.Ideal.MessageTypeFromString(strMessageType)
                                
                                frmTest2.AddList strMessageType & " ( " & Str(lMsgNumber) & " ): " & strMessage
                                If lMsgNumber <> -1& Then
                                    g.Ideal.HandleMessage lMsgNumber, strMessage
                                End If
                            End If
                        
                        Case eTT_AccountType_IntBrokers
                            If Not g.IntBroker Is Nothing Then
                                lMsgNumber = g.IntBroker.MessageTypeFromString(strMessageType)
                                
                                frmTest2.AddList strMessageType & " ( " & Str(lMsgNumber) & " ): " & strMessage
                                If lMsgNumber <> -1& Then
                                    g.IntBroker.HandleMessage lMsgNumber, strMessage
                                End If
                            End If
                        
                        Case eTT_AccountType_Optimus
                            If Not g.Optimus Is Nothing Then
                                lMsgNumber = g.Optimus.StringToMessageType(strMessageType)
                                
                                frmTest2.AddList strMessageType & " ( " & Str(lMsgNumber) & " ): " & strMessage
                                If lMsgNumber <> -1& Then
                                    g.Optimus.HandleMessage lMsgNumber, strMessage
                                End If
                            End If
                        
                        Case eTT_AccountType_OpVest
                            If Not g.OpVest Is Nothing Then
                                lMsgNumber = g.OpVest.StringToMessageType(strMessageType)
                                
                                frmTest2.AddList strMessageType & " ( " & Str(lMsgNumber) & " ): " & strMessage
                                If lMsgNumber <> -1& Then
                                    g.OpVest.HandleMessage lMsgNumber, strMessage
                                End If
                            End If
                        
                        Case eTT_AccountType_Rithmic
                            If Not g.Rithmic Is Nothing Then
                                lMsgNumber = g.Rithmic.StringToMessageType(strMessageType)
                                
                                frmTest2.AddList strMessageType & " ( " & Str(lMsgNumber) & " ): " & strMessage
                                If lMsgNumber <> -1& Then
                                    g.Rithmic.HandleMessage lMsgNumber, strMessage
                                End If
                            End If
                        
                        Case eTT_AccountType_TransAct
                            If Not g.Transact Is Nothing Then
                                lMsgNumber = g.Transact.StringToMessageType(strMessageType)
                                
                                frmTest2.AddList strMessageType & " ( " & Str(lMsgNumber) & " ): " & strMessage
                                If lMsgNumber <> -1& Then
                                    g.Transact.HandleMessage lMsgNumber, strMessage
                                End If
                            End If
                        
                        Case eTT_AccountType_Vision
                            If Not g.Vision Is Nothing Then
                                lMsgNumber = g.Vision.StringToMessageType(strMessageType)
                                
                                frmTest2.AddList strMessageType & " ( " & Str(lMsgNumber) & " ): " & strMessage
                                If lMsgNumber <> -1& Then
                                    g.Vision.HandleMessage lMsgNumber, strMessage
                                End If
                            End If
                        
                        Case eTT_AccountType_ZenFire
                            If Not g.ZenFire Is Nothing Then
                                lMsgNumber = g.ZenFire.StringToMessageType(strMessageType)
                                
                                frmTest2.AddList strMessageType & " ( " & Str(lMsgNumber) & " ): " & strMessage
                                If lMsgNumber <> -1& Then
                                    g.ZenFire.HandleMessage lMsgNumber, strMessage
                                End If
                            End If
                        
                        Case Else
                            Set BrokerObj = Broker(nBroker)
                            If Not BrokerObj Is Nothing Then
                                lMsgNumber = mTradeTracker.BrokerMessageTypeFromString(strMessageType)
                                
                                frmTest2.AddList strMessageType & " ( " & Str(lMsgNumber) & " ): " & strMessage
                                If lMsgNumber <> -1& Then
                                    BrokerObj.HandleMessage lMsgNumber, strMessage
                                End If
                            End If
                    End Select
                End If
            End If
        Next lIndex
    End If
    
    frmTest2.AddList "Done"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.ReplayLogFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TempBrokerAccount
'' Description: Create a temporary broker account for purposes of connection
'' Inputs:      Broker Type, Broker User
'' Returns:     Temporary Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TempBrokerAccount(ByVal nBroker As eTT_AccountType, ByVal bBrokerUser As Boolean)
On Error GoTo ErrSection:

    If FormIsLoaded("frmTTSummary") Then
        frmTTSummary.TempBrokerAccount nBroker, bBrokerUser
    End If
    If FormIsLoaded("frmAccounts") Then
        frmAccounts.TempBrokerAccount nBroker, bBrokerUser
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.TempBrokerAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeSymbolOrID
'' Description: Determine the trade symbol or ID for the symbol passed in
'' Inputs:      Symbol or ID
'' Returns:     Trade Symbol or ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TradeSymbolOrID(ByVal vSymbolOrSymbolID As Variant) As Variant
On Error GoTo ErrSection:

    Dim vTradeSymbolOrID As Variant     ' Trade symbol or symbol ID
    Dim strSymbol As String             ' Symbol for the symbol or ID passed in
    Dim lSymbolID As Long               ' Symbol ID for the symbol or ID passed in
    
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    
    If lSymbolID = 0 Then
        vTradeSymbolOrID = ConvertToTradeSymbol(strSymbol, Int(CurrentTime("", "", True)))
    Else
        vTradeSymbolOrID = ConvertToTradeSymbol(lSymbolID, Int(CurrentTime("", "", True)))
    End If
    
    TradeSymbolOrID = vTradeSymbolOrID

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.TradeSymbolOrID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSfeAllowedArray
'' Description: Do a one-time fill of the SfeAllowed array
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillSfeAllowedArray()
On Error GoTo ErrSection:

    Dim astrExchanges As New cGdArray   ' List of exchanges
    Dim lLastDate As Long               ' Last successful connection to broker
    
    If IsBrokerUser(eTT_AccountType_Ideal) Then
        If FileExist(AddSlash(App.Path) & "Provided\LkExch.IDL") Then
            astrExchanges.Serialize AddSlash(App.Path) & "Provided\LkExch.IDL", False
            If astrExchanges.BinarySearch("SNFE") Then
                lLastDate = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", AddSlash(App.Path) & "Ideal.INI"))))
                SfeAllowed(eTT_AccountType_Ideal) = lLastDate
            End If
        End If
    End If
    
    If IsBrokerUser(eTT_AccountType_IntBrokers) Then
        If FileExist(AddSlash(App.Path) & "Provided\LkExch.IB") Then
            astrExchanges.Serialize AddSlash(App.Path) & "Provided\LkExch.IB", False
            If astrExchanges.BinarySearch("SNFE") Then
                lLastDate = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", AddSlash(App.Path) & "IntBrokers2.INI"))))
                SfeAllowed(eTT_AccountType_IntBrokers) = lLastDate
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.FillSfeAllowedArray"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetProvidedBrokerProperty
'' Description: Get a broker specific property out of the Provided.INI
'' Inputs:      Account, Symbol, Section, Default
'' Returns:     Property value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetProvidedBrokerProperty(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal strSection As String, ByVal lDefaultValue As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim nBroker As eTT_AccountType      ' Account type for the given account
    Dim strSecType As String            ' Security type for the given symbol
    
    nBroker = AccountTypeForID(GetAccountID(vAccountNumberOrID))
    strSecType = TradeSecType(vSymbolOrSymbolID)
    
    lReturn = GetIniFileProperty(strSecType & "_" & Str(nBroker), kNullData, strSection, AddSlash(App.Path) & "Provided\Provided.INI")
    If lReturn = kNullData Then
        lReturn = GetIniFileProperty(strSecType, lDefaultValue, strSection, AddSlash(App.Path) & "Provided\Provided.INI")
    End If
    
    GetProvidedBrokerProperty = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerDispatch.GetProvidedBrokerProperty"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Dim bConfirmManual As Boolean       ' Old confirmation flag

    Set m.astrAccountTypes = New cGdArray
    m.astrAccountTypes.Create eGDARRAY_Strings
    
    Set m.alSfeAllowed = New cGdArray
    If FileExist(AddSlash(App.Path) & kSfeDat) Then
        m.alSfeAllowed.Serialize AddSlash(App.Path) & kSfeDat, False
    Else
        m.alSfeAllowed.Create eGDARRAY_Longs, kNumBrokers, kNullData
        FillSfeAllowedArray
    End If
    DebugLog "SfeAllowed = " & m.alSfeAllowed.JoinFields(",")
    
    m.dLastStatusChange = 0
    m.bPositionsToVerify = False
    
    m.bShowCents = GetIniFileProperty("ShowCents", True, "TTSummary", g.strIniFile)
    m.strGridFont = GetIniFileProperty("TTSummary", "", "Fonts", g.strIniFile)
    m.bDontAllowStopMove = GetIniFileProperty("DontAllowStopMove", True, "Trading", g.strIniFile)
    bConfirmManual = GetIniFileProperty("ConfirmOrders", True, "General", g.strIniFile)
    m.bConfirmManual = GetIniFileProperty("ConfirmManual", bConfirmManual, "Trading", g.strIniFile)
    m.bConfirmTriggered = GetIniFileProperty("ConfirmTriggered", True, "Trading", g.strIniFile)
    m.bConfirmTradeSense = GetIniFileProperty("ConfirmTradeSense", False, "Trading", g.strIniFile)
    m.bWarnStopWrongSide = GetIniFileProperty("WarnStopWrongSide", True, "Trading", g.strIniFile)
    m.bWarnLimitWrongSide = GetIniFileProperty("WarnLimitWrongSide", True, "Trading", g.strIniFile)
    m.lNumTicksStopBuffer = GetIniFileProperty("NumTicksStopBuffer", 2&, "Trading", g.strIniFile)
    m.bAutoJournalPopUp = GetIniFileProperty("AutoJournalPopUp", False, "Trading", g.strIniFile)
    m.bAutoJournalAutomated = GetIniFileProperty("AutoJournalAutomated", False, "Trading", g.strIniFile)
    m.nOptionFillMethod = GetIniFileProperty("OptionFillMethod", eGDOptionFill_BidOrAsk, "Trading", g.strIniFile)
    m.nOptionOpenEquity = GetIniFileProperty("OptionOpenEquity", eGDOptionOpenEquity_UseBidAsk, "Trading", g.strIniFile)
    
    m.strWebAccounts = GetIniFileProperty("Accounts", "", "WebPortfolio", g.strIniFile)
    m.lWebFontSize = GetIniFileProperty("FontSize", 4&, "WebPortfolio", g.strIniFile)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Destroy the class members when the object is destroyed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.astrAccountTypes = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerDispatch.Class_Terminate"
    
End Sub
