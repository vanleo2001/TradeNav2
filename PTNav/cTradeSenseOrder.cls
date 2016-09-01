VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeSenseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeSenseOrder.cls
'' Description: Class that handles a Trade Sense order
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 06/03/2010   DAJ         Changes for new Trade Sense Order Groups
'' 06/17/2010   DAJ         Changed filenames to ID instead of name
'' 06/24/2010   DAJ         Added the Tooltip function
'' 06/28/2010   DAJ         Added Use flag for turning off orders
'' 07/15/2010   DAJ         Added capabilities for inputs
'' 08/10/2010   DAJ         Fix all negative OCO's, not just the first one
'' 08/11/2010   DAJ         Return new bar flag from UpdateBars, Add input values on Tool Tip
'' 08/12/2010   DAJ         Check for streaming availability
'' 08/12/2010   DAJ         Fix issue with parameter replacement
'' 08/19/2010   DAJ         Added capability for TradeSense order expire at end of session
'' 08/23/2010   DAJ         Added required module flag for TradeSense orders/groups
'' 09/13/2010   DAJ         Added checks to make sure that expression object is valid
'' 09/30/2010   DAJ         Added optional argument to AllRtDataAvailable function
'' 10/14/2010   DAJ         Convert expiration time to NY for compare with feed time
'' 10/25/2010   DAJ         Don't do anything in CalcExpiration if SetBarProperties fails
'' 10/29/2010   DAJ         Don't do anything in CalcExpiration if Expr.SessionDate is Null
'' 11/04/2010   DAJ         Fix expiration bug when activate between sessions
'' 12/17/2010   DAJ         Added UpdateBarsSince argument to FillCallback
'' 01/07/2011   DAJ         Changed the activation calls
'' 04/29/2011   DAJ         Added CurrentPrices call
'' 05/18/2011   DAJ         Added custom start/stop time for Market1
'' 06/08/2011   DAJ         Send IntraBar flag over to RunExpression on Init
'' 06/09/2011   DAJ         Implemented auto-breakout bar period
'' 09/08/2011   DAJ         Moved auto breakout calculation down to cRunExpression
'' 10/07/2011   DAJ         Exposed the RunExpressions so it could be called explicitly from outside
'' 09/26/2012   DAJ         Added New Session flag to UpdateBars
'' 10/03/2012   DAJ         Lot size for forex symbols in TradeSense order groups
'' 10/03/2012   DAJ         Added LastGoodBar mode for cRunExpression and cLateCalculating
'' 11/28/2012   DAJ         Submit Market if Stop on wrong side of market flag
'' 05/10/2013   DAJ         Round order prices to the nearest tick
'' 01/31/2014   DAJ         Log the current OHLC when the session changes
'' 06/26/2014   DAJ         Roll position on a TradeSense order group
'' 02/05/2015   DAJ         Use "WithLimitPrice" for LimitPrice when applicable
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kExtension = "TSO"

Private Enum eGDRtMode
    eGDRtMode_Off = 0
    eGDRtMode_On
    eGDRtMode_Suspended
End Enum

Private Type mPrivate
    strID As String                     ' Unique identifier used for the filename
    strName As String                   ' Name of the Trade Sense order
    strConditionFormatted As String     ' Order Condition in formatted text
    strConditionCoded As String         ' Order Condition in coded text
    strPriceFormatted As String         ' Price Condition in formatted text
    strPriceCoded As String             ' Price Condition in coded text
    strWithLimitFormatted As String     ' With Limit Price in formatted text
    strWithLimitCoded As String         ' With Limit Price in coded text
    bBuy As Boolean                     ' Is this a buy order?
    bEntry As Boolean                   ' Is this an entry order?
    nOrderType As eTT_OrderType         ' Type of the order
    lExitPercent As Long                ' Percentage of position to exit
    bCustom As Boolean                  ' Is this a custom order?
    bExpireDay As Boolean               ' Is this a day only order?
    dExpiration As Double               ' Date/Time order will expire
    strRequiredMod As String            ' Required module for this to show up
    bMarketIfWrongSide As Boolean       ' Submit a market order if on wrong side of the market?
    
    strBarPeriod As String              ' Bar period
    'strAutoBreakoutPeriod As String     ' Bar period for auto-breakout bars
    bOverride As Boolean                ' Override the number bars required?
    lNumBarsAuto As Long                ' Number of bars required calculated
    lNumBarsOver As Long                ' Number of bars required override
    bIntraBar As Boolean                ' Needs to be checked intra-bar?
    nRtMode As eGDRtMode                ' Real-time checking mode
    bRtDataAvailable As Boolean         ' Real-time data available?
    lCurrentSession As Long             ' Current session date
    
    lOrderNumber As Long                ' Order number for groups
    bUse As Boolean                     ' Should this order be used?
    lOTO As Long                        ' Order-trigger-order ID for groups
    alOCOs As cGdArray                  ' Array of OCO orders
    nBroker As eTT_AccountType          ' Broker for this order
    strLogText As String                ' Text for dumping to a log
    
    Rule As cRule                       ' Rule object for creating RTF
    
    strSymbol As String                 ' Symbol for the Bars
    Expr As cRunExpression              ' Run expression object
    Bars As cGdBars                     ' Bars object
    
    bConditionTrue As Boolean           ' Is the condition true?
    dPrice As Double                    ' Price for the order
    dWithLimitPrice As Double           ' With Limit Price for the order
    
    bAllowInputs As Boolean             ' Allow inputs in the expressions?
    Inputs As cTradeSenseOrderInputs    ' Collection of expression inputs
End Type
Private m As mPrivate

Public Property Get ID() As String
    ID = m.strID
End Property

Public Property Get Name() As String
    Name = m.strName
End Property
Public Property Let Name(ByVal strName As String)
    m.strName = strName
End Property

Public Property Get ConditionFormatted() As String
    ConditionFormatted = m.strConditionFormatted
End Property
Public Property Let ConditionFormatted(ByVal strConditionFormatted As String)
    m.strConditionFormatted = strConditionFormatted
End Property

Public Property Get ConditionCoded() As String
    ConditionCoded = m.strConditionCoded
End Property
Public Property Let ConditionCoded(ByVal strConditionCoded As String)
    m.strConditionCoded = strConditionCoded
End Property

Public Property Get ConditionRTF() As String
    ConditionRTF = m.Rule.GetRTF(m.strConditionFormatted)
End Property

Public Property Get ConditionEnglish(Optional ByVal bReplaceLateCalc As Boolean = True) As String
    ConditionEnglish = m.Expr.EnglishExpression(0, bReplaceLateCalc)
End Property

Public Property Get PriceFormatted() As String
    PriceFormatted = m.strPriceFormatted
End Property
Public Property Let PriceFormatted(ByVal strPriceFormatted As String)
    m.strPriceFormatted = strPriceFormatted
End Property

Public Property Get PriceCoded() As String
    PriceCoded = m.strPriceCoded
End Property
Public Property Let PriceCoded(ByVal strPriceCoded As String)
    m.strPriceCoded = strPriceCoded
End Property

Public Property Get PriceRTF() As String
    PriceRTF = m.Rule.GetRTF(m.strPriceFormatted)
End Property

Public Property Get PriceEnglish(Optional ByVal bReplaceLateCalc As Boolean = True) As String
    PriceEnglish = m.Expr.EnglishExpression(1, bReplaceLateCalc)
End Property

Public Property Get WithLimitFormatted() As String
    WithLimitFormatted = m.strWithLimitFormatted
End Property
Public Property Let WithLimitFormatted(ByVal strWithLimitFormatted As String)
    m.strWithLimitFormatted = strWithLimitFormatted
End Property

Public Property Get WithLimitCoded() As String
    WithLimitCoded = m.strWithLimitCoded
End Property
Public Property Let WithLimitCoded(ByVal strWithLimitCoded As String)
    m.strWithLimitCoded = strWithLimitCoded
End Property

Public Property Get WithLimitRTF() As String
    WithLimitRTF = m.Rule.GetRTF(m.strWithLimitFormatted)
End Property

Public Property Get WithLimitEnglish(Optional ByVal bReplaceLateCalc As Boolean = True) As String
    WithLimitEnglish = m.Expr.EnglishExpression(2, bReplaceLateCalc)
End Property

Public Property Get Buy() As Boolean
    Buy = m.bBuy
End Property
Public Property Let Buy(ByVal bBuy As Boolean)
    m.bBuy = bBuy
End Property

Public Property Get Entry() As Boolean
    Entry = m.bEntry
End Property
Public Property Let Entry(ByVal bEntry As Boolean)
    m.bEntry = bEntry
End Property

Public Property Get OrderType() As eTT_OrderType
    OrderType = m.nOrderType
End Property
Public Property Let OrderType(ByVal nOrderType As eTT_OrderType)
    m.nOrderType = nOrderType
End Property

Public Property Get ExitPercent() As Long
    ExitPercent = m.lExitPercent
End Property
Public Property Let ExitPercent(ByVal lExitPercent As Long)
    m.lExitPercent = lExitPercent
End Property

Public Property Get BarPeriod() As String
    BarPeriod = m.strBarPeriod
End Property
Public Property Let BarPeriod(ByVal strBarPeriod As String)
    m.strBarPeriod = strBarPeriod
End Property

Public Property Get Override() As Boolean
    Override = m.bOverride
End Property
Public Property Let Override(ByVal bOverride As Boolean)
    m.bOverride = bOverride
End Property

Public Property Get NumBarsAuto() As Long
    NumBarsAuto = m.lNumBarsAuto
End Property
Public Property Let NumBarsAuto(ByVal lNumBarsAuto As Long)
    m.lNumBarsAuto = lNumBarsAuto
End Property

Public Property Get NumBarsOverride() As Long
    NumBarsOverride = m.lNumBarsOver
End Property
Public Property Let NumBarsOverride(ByVal lNumBarsOverride As Long)
    m.lNumBarsOver = lNumBarsOverride
End Property

Public Property Get IntraBar() As Boolean
    IntraBar = m.bIntraBar
End Property
Public Property Let IntraBar(ByVal bIntraBar As Boolean)
    m.bIntraBar = bIntraBar
End Property

Public Property Get ExpireDay() As Boolean
    ExpireDay = m.bExpireDay
End Property
Public Property Let ExpireDay(ByVal bExpireDay As Boolean)
    m.bExpireDay = bExpireDay
End Property

Public Property Get RequiredMod() As String
    RequiredMod = m.strRequiredMod
End Property
Public Property Let RequiredMod(ByVal strRequiredMod As String)
    m.strRequiredMod = strRequiredMod
End Property

Public Property Get MarketIfWrongSide() As Boolean
    MarketIfWrongSide = m.bMarketIfWrongSide
End Property
Public Property Let MarketIfWrongSide(ByVal bMarketIfWrongSide As Boolean)
    m.bMarketIfWrongSide = bMarketIfWrongSide
End Property

Public Property Get ConditionTrue() As Boolean
    ConditionTrue = m.bConditionTrue
End Property

Public Property Get Price() As Double
    Price = m.dPrice
End Property

Public Property Get WithLimitPrice() As Double
    WithLimitPrice = m.dWithLimitPrice
End Property

Public Property Get Custom() As Boolean
    Custom = m.bCustom
End Property

Public Property Get OrderNumber() As Long
    OrderNumber = m.lOrderNumber
End Property
Public Property Let OrderNumber(ByVal lOrderNumber As Long)
    m.lOrderNumber = lOrderNumber
End Property

Public Property Get Use() As Boolean
    Use = m.bUse
End Property
Public Property Let Use(ByVal bUse As Boolean)
    m.bUse = bUse
End Property

Public Property Get OCOs() As cGdArray
    Set OCOs = m.alOCOs
End Property
Public Property Let OCOs(ByVal alOCOs As cGdArray)
    Set m.alOCOs = alOCOs
End Property

Public Property Get OTO() As Long
    OTO = m.lOTO
End Property
Public Property Let OTO(ByVal lOTO As Long)
    m.lOTO = lOTO
End Property

Public Property Get AllowInputs() As Boolean
    AllowInputs = m.bAllowInputs
End Property
Public Property Let AllowInputs(ByVal bAllowInputs As Boolean)
    m.bAllowInputs = bAllowInputs
End Property

Public Property Get Inputs() As cTradeSenseOrderInputs
    Set Inputs = m.Inputs
End Property
Public Property Let Inputs(ByVal tsInputs As cTradeSenseOrderInputs)
    Set m.Inputs = tsInputs
End Property

Public Property Get Expiration() As Double
    Expiration = m.dExpiration
End Property

Public Sub ClearID()
    m.strID = ""
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Extract the class members from the given string
'' Inputs:      Order string
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strTradeSenseOrder As String)
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of fields

    astrFields.SplitFields strTradeSenseOrder, vbTab
    m.strConditionFormatted = astrFields(0)
    m.strConditionCoded = astrFields(1)
    m.strPriceFormatted = astrFields(2)
    m.strPriceCoded = astrFields(3)
    m.strWithLimitFormatted = astrFields(4)
    m.strWithLimitCoded = astrFields(5)
    m.bBuy = (astrFields(6) = "1")
    m.bEntry = (astrFields(7) = "1")
    m.nOrderType = CLng(Val(astrFields(8)))
    m.strBarPeriod = astrFields(9)
    m.bOverride = (astrFields(10) = "1")
    m.lNumBarsAuto = CLng(Val(astrFields(11)))
    m.lNumBarsOver = CLng(Val(astrFields(12)))
    If Len(astrFields(13)) > 0 Then
        m.lExitPercent = CLng(Val(astrFields(13)))
    Else
        m.lExitPercent = 100
    End If
    If Len(astrFields(14)) > 0 Then
        m.bIntraBar = (astrFields(14) = "1")
    Else
        m.bIntraBar = False
    End If
    If Len(astrFields(15)) > 0 Then
        m.bExpireDay = (astrFields(15) = "1")
    Else
        m.bExpireDay = False
    End If
    If Len(astrFields(16)) > 0 Then
        m.strRequiredMod = astrFields(16)
    Else
        m.strRequiredMod = ""
    End If
    If Len(astrFields(17)) > 0 Then
        m.bMarketIfWrongSide = (astrFields(17) = "1")
    Else
        m.bMarketIfWrongSide = False
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Build a string from the class members
'' Inputs:      None
'' Returns:     Order string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of fields
    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Len(m.strConditionCoded) > 0 Then
        astrFields.Add m.strConditionFormatted
        astrFields.Add m.strConditionCoded
        astrFields.Add m.strPriceFormatted
        astrFields.Add m.strPriceCoded
        astrFields.Add m.strWithLimitFormatted
        astrFields.Add m.strWithLimitCoded
        astrFields.Add BoolToString(m.bBuy)
        astrFields.Add BoolToString(m.bEntry)
        astrFields.Add Str(m.nOrderType)
        astrFields.Add m.strBarPeriod
        astrFields.Add BoolToString(m.bOverride)
        astrFields.Add Str(m.lNumBarsAuto)
        astrFields.Add Str(m.lNumBarsOver)
        astrFields.Add Str(m.lExitPercent)
        astrFields.Add BoolToString(m.bIntraBar)
        astrFields.Add BoolToString(m.bExpireDay)
        astrFields.Add m.strRequiredMod
        astrFields.Add BoolToString(m.bMarketIfWrongSide)
    
        strReturn = astrFields.JoinFields(vbTab)
    End If
    
    ToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Preview
'' Description: Build the preview string from the members
'' Inputs:      None
'' Returns:     Preview string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Preview() As String
On Error GoTo ErrSection:

    Dim strPreview As String             ' Return value for the function
    Dim strOrderType As String          ' Order type string

    strPreview = ""
    If (Len(m.strConditionFormatted) > 0) Then
        If m.bBuy Then
            strPreview = m.strConditionFormatted & "~80001 ~24005THEN ~24007BUY at "
        Else
            strPreview = m.strConditionFormatted & "~80001 ~24005THEN ~24008SELL at "
        End If
        
        strOrderType = mTradeTracker.OrderType(m.nOrderType)
        
        Select Case m.nOrderType
            Case eTT_OrderType_Market, eTT_OrderType_MarketOnClose
                strPreview = strPreview & "~50" & Format(Len(strOrderType), "000") & UCase(strOrderType)
                
            Case eTT_OrderType_StopWithLimit
                strPreview = strPreview & m.strPriceFormatted & " ~50007 STOP, " & m.strWithLimitFormatted & " ~50006 LIMIT"
        
            Case eTT_OrderType_StopWithLimitCloseOnly
                strPreview = strPreview & m.strPriceFormatted & " ~50007 STOP, " & m.strWithLimitFormatted & " ~500017 LIMIT CLOSE ONLY"
                
            Case Else
                strPreview = strPreview & m.strPriceFormatted & " ~50001 " & "~50" & Format(Len(strOrderType), "000") & UCase(strOrderType)
        End Select
    End If
    
    Preview = strPreview

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.Preview"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PreviewRTF
'' Description: Build the preview string from the members and convert to RTF
'' Inputs:      None
'' Returns:     Preview string in RTF
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PreviewRTF() As String
On Error GoTo ErrSection:

    PreviewRTF = m.Rule.GetRTF(Preview)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.PreviewRTF"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActivateWithAcctPos
'' Description: Activate the Trade Sense order
'' Inputs:      Symbol, Account Position, Broker, Log Text, Custom Start Time,
''              Custom Stop Time, Auto Breakout Period, Lot Size
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ActivateWithAcctPos(ByVal strSymbol As String, ByVal AcctPos As cAccountPosition, Optional ByVal nBroker As eTT_AccountType = -1&, Optional ByVal strLogText As String = "", Optional ByVal Fills As cGdTree = Nothing, Optional ByVal dStartTime As Double = kNullData, Optional ByVal dStopTime As Double = kNullData, Optional ByVal lLotSize As Long = 1&) ', Optional ByVal strAutoBreakoutPeriod = "")
On Error GoTo ErrSection:

    Activate strSymbol, AcctPos, Nothing, nBroker, strLogText, dStartTime, dStopTime, lLotSize ', strAutoBreakoutPeriod
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.ActivateWithAcctPos"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActivateWithFills
'' Description: Activate the Trade Sense order
'' Inputs:      Symbol, Fills, Broker, Log Text, Custom Start Time,
''              Custom Stop Time, Auto Breakout Period, Lot Size
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ActivateWithFills(ByVal strSymbol As String, ByVal Fills As cGdTree, Optional ByVal nBroker As eTT_AccountType = -1&, Optional ByVal strLogText As String = "", Optional ByVal dStartTime As Double = kNullData, Optional ByVal dStopTime As Double = kNullData, Optional ByVal lLotSize As Long = 1&) ', Optional ByVal strAutoBreakoutPeriod = "")
On Error GoTo ErrSection:

    Activate strSymbol, Nothing, Fills, nBroker, strLogText, dStartTime, dStopTime, lLotSize ', strAutoBreakoutPeriod
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.ActivateWithFills"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBars
'' Description: Update the bars
'' Inputs:      New Bar?, Expire the order?, New session?
'' Returns:     True if expressions run, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UpdateBars(Optional bNewBar As Boolean, Optional bExpire As Boolean, Optional bNewSession As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lCurrentSession As Long         ' Current session date

    bReturn = False
    If Not m.Expr Is Nothing Then
        bNewBar = m.Expr.UpdateBars(True, (m.nRtMode = eGDRtMode_On), bNewSession)
        If (bNewBar = True) And (m.nRtMode = eGDRtMode_Suspended) Then
            m.nRtMode = eGDRtMode_On
        End If
        
        lCurrentSession = m.Bars.SessionDateForTradeTime(CurrentTime(m.Bars.Prop(eBARS_ExchangeTimeZoneInf), m.Bars.Prop(eBARS_Symbol), True))
        If m.lCurrentSession = 0& Then
            m.lCurrentSession = lCurrentSession
        ElseIf lCurrentSession <> m.lCurrentSession Then
            m.lCurrentSession = lCurrentSession
            g.Broker.BrokerDebug m.nBroker, m.strLogText & "New Session ( " & DateFormat(m.lCurrentSession) & " ) - " & m.Expr.CurrentPrices
        End If
        
        bExpire = False
        If m.bExpireDay Then
            If m.dExpiration = kNullData Then
                CalcExpiration
            ElseIf g.RealTime.FeedTime(m.strSymbol) >= m.dExpiration Then
                bExpire = True
            End If
        End If
        
        bReturn = bNewBar Or (m.nRtMode = eGDRtMode_On)
        If bReturn = True Then
            SetValues
            
            If ConditionTrue And (m.nRtMode = eGDRtMode_On) Then
                m.nRtMode = eGDRtMode_Suspended
            End If
        End If
    End If
    
    UpdateBars = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.UpdateBars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdatePosition
'' Description: Update the position information for the late calculating functions
'' Inputs:      Account Position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdatePosition(ByVal AcctPos As cAccountPosition)
On Error GoTo ErrSection:
    
    If Not m.Expr Is Nothing Then
        m.Expr.UpdatePosition AcctPos
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.UpdatePosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillCallback
'' Description: Update the position information for the late calculating functions
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FillCallback(ByVal Fill As cPtFill)
On Error GoTo ErrSection:
    
    If Not m.Expr Is Nothing Then
        m.Expr.FillCallback Fill
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.FillCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnglishPreview
'' Description: Build an english preview for the order for logging
'' Inputs:      Replace Late Calculating stuff?
'' Returns:     Order Preview in English
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnglishPreview(Optional ByVal bReplaceLateCalc As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    If m.bBuy Then
        strReturn = ConditionEnglish(bReplaceLateCalc) & " THEN BUY at " & EnglishOrderPrice(bReplaceLateCalc)
    Else
        strReturn = ConditionEnglish(bReplaceLateCalc) & " THEN SELL at " & EnglishOrderPrice(bReplaceLateCalc)
    End If
    
    EnglishPreview = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.EnglishPreview"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnglishOrderPrice
'' Description: Build an english price for the order for logging
'' Inputs:      None
'' Returns:     Order Price in English
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnglishOrderPrice(Optional ByVal bReplaceLateCalc As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    Select Case m.nOrderType
        Case eTT_OrderType_Market
            strReturn = "MARKET"
        Case eTT_OrderType_Stop
            strReturn = PriceEnglish(bReplaceLateCalc) & " STOP"
        Case eTT_OrderType_Limit
            strReturn = PriceEnglish(bReplaceLateCalc) & " LIMIT"
        Case eTT_OrderType_StopWithLimit
            strReturn = PriceEnglish(bReplaceLateCalc) & " STOP with a " & WithLimitEnglish(bReplaceLateCalc) & " LIMIT"
    End Select
    
    EnglishOrderPrice = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.EnglishOrderPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StopPrice
'' Description: Stop price for the order
'' Inputs:      None
'' Returns:     Stop Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StopPrice() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    
    dReturn = 0#
    If (m.nOrderType = eTT_OrderType_Stop) Or (m.nOrderType = eTT_OrderType_StopWithLimit) Then
        If m.bBuy Then
            dReturn = m.Bars.NearestPriceUp(RoundToSigDigits(Price))
        Else
            dReturn = m.Bars.NearestPriceDown(RoundToSigDigits(Price))
        End If
    End If
    
    StopPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.StopPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LimitPrice
'' Description: Limit price for the order
'' Inputs:      None
'' Returns:     Limit Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LimitPrice() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    
    dReturn = 0#
    If m.nOrderType = eTT_OrderType_Limit Then
        If m.bBuy Then
            dReturn = m.Bars.NearestPriceDown(RoundToSigDigits(Price))
        Else
            dReturn = m.Bars.NearestPriceUp(RoundToSigDigits(Price))
        End If
    ElseIf m.nOrderType = eTT_OrderType_StopWithLimit Then
        If m.bBuy Then
            dReturn = m.Bars.NearestPriceUp(RoundToSigDigits(WithLimitPrice))
        Else
            dReturn = m.Bars.NearestPriceDown(RoundToSigDigits(WithLimitPrice))
        End If
    End If
    
    LimitPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.LimitPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Action
'' Description: Build an action string from the Buy/Sell and Entry/Exit flags
'' Inputs:      None
'' Returns:     Action
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Action() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    If m.bBuy Then
        If m.bEntry Then
            strReturn = "Long Entry"
        Else
            strReturn = "Short Exit"
        End If
    Else
        If m.bEntry Then
            strReturn = "Short Entry"
        Else
            strReturn = "Long Exit"
        End If
    End If
    
    Action = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.Action"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Filename
'' Description: Build the filename for this Trade Sense Order
'' Inputs:      Full Path?
'' Returns:     Filename and Path
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FileName(Optional ByVal bFullPath As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If bFullPath Then
        strReturn = AddSlash(App.Path)
    Else
        strReturn = ""
    End If
    
    If m.bCustom Then
        strReturn = strReturn & "Custom\" & m.strID & "." & kExtension
    Else
        strReturn = strReturn & "Provided\" & m.strID & "." & kExtension
    End If
    
    FileName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.Filename"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromFile
'' Description: Loads the Trade Sense order from a file
'' Inputs:      Filename
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromFile(ByVal strFileName As String)
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' File read into an array
    
    m.strID = FileBase(strFileName)
    m.bCustom = (UCase(FilePath(strFileName)) = UCase(AddSlash(App.Path) & "CUSTOM\"))

    If astrFile.FromFile(strFileName) Then
        m.strName = astrFile(0)
        FromString astrFile(1)
        m.bAllowInputs = (Len(astrFile(2)) > 0) And (astrFile(2) <> "0")
        m.Inputs.FromString astrFile(3)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.FromFile"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToFile
'' Description: Save the Trade Sense order to a file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ToFile()
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' Array of information to spit out to a file
    
    If Len(m.strID) = 0 Then
        m.strID = FileBase(GetUniqueCustomFilename(kExtension))
    End If

    astrFile(0) = m.strName
    astrFile(1) = ToString
    astrFile(2) = Str(m.bAllowInputs)
    astrFile(3) = m.Inputs.ToString
    
    astrFile.ToFile FileName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.ToFile"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromGroupString
'' Description: Set the class members from the given group string
'' Inputs:      Group string
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromGroupString(ByVal strGroupString As String)
On Error GoTo ErrSection:

    Dim astrInfo As New cGdArray        ' Array of split out information
    
    astrInfo.SplitFields strGroupString, "|"
    
    FromFile AddSlash(App.Path) & astrInfo(3)
    
    m.lOrderNumber = CLng(Val(astrInfo(0)))
    m.alOCOs.SplitFields astrInfo(1), ","
    m.lOTO = CLng(Val(astrInfo(2)))
    
    If Len(astrInfo(4)) > 0 Then
        m.bUse = (astrInfo(4) <> "0")
    Else
        m.bUse = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.FromGroupString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToGroupString
'' Description: Put the appropriate information together for a group string
'' Inputs:      None
'' Returns:     Group string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToGroupString() As String
On Error GoTo ErrSection:

    Dim astrReturn As New cGdArray      ' Array of information to return
    
    astrReturn.Create eGDARRAY_Strings, 4
    astrReturn(0) = Str(m.lOrderNumber)
    astrReturn(1) = m.alOCOs.JoinFields(",")
    astrReturn(2) = Str(m.lOTO)
    astrReturn(3) = FileName(False)
    astrReturn(4) = BoolToString(m.bUse)
    
    ToGroupString = astrReturn.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.ToGroupString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOCO
'' Description: Is the given order number an OCO with this one?
'' Inputs:      Order Number
'' Returns:     True if OCO, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsOCO(ByVal lOrderNumber As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    
    For lIndex = 0 To m.alOCOs.Size - 1
        If m.alOCOs(lIndex) = lOrderNumber Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    IsOCO = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.IsOCO"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveOCO
'' Description: Remove the OCO with the given order number
'' Inputs:      Order Number
'' Returns:     True if Removed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemoveOCO(ByVal lOrderNumber As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    
    For lIndex = 0 To m.alOCOs.Size - 1
        If m.alOCOs(lIndex) = lOrderNumber Then
            bReturn = True
            m.alOCOs.Remove lIndex
            Exit For
        End If
    Next lIndex
    
    RemoveOCO = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.RemoveOCO"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeOCO
'' Description: Change the OCO with the given order number to the new number
'' Inputs:      Order Number, New Order Number
'' Returns:     True if Removed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ChangeOCO(ByVal lOrderNumber As Long, ByVal lNewOrderNumber As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    
    For lIndex = 0 To m.alOCOs.Size - 1
        If m.alOCOs(lIndex) = lOrderNumber Then
            bReturn = True
            m.alOCOs(lIndex) = lNewOrderNumber
            Exit For
        End If
    Next lIndex
    
    ChangeOCO = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.ChangeOCO"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixNegativeOCO
'' Description: Fix any negative OCO's
'' Inputs:      None
'' Returns:     True if anything fixed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FixNegativeOCO() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    
    For lIndex = 0 To m.alOCOs.Size - 1
        If m.alOCOs(lIndex) < 0 Then
            bReturn = True
            m.alOCOs(lIndex) = m.alOCOs(lIndex) * -1&
        End If
    Next lIndex
    
    FixNegativeOCO = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.FixNegativeOCO"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Tooltip
'' Description: Build text for a tooltip in a grid
'' Inputs:      None
'' Returns:     Tooltip text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToolTip() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrInputs As cGdArray          ' Inputs portion of the tool tip text
    
    strReturn = EnglishFromCoded(m.strConditionFormatted)
    
    If m.bBuy Then
        strReturn = strReturn & " THEN Buy at "
    Else
        strReturn = strReturn & " THEN Sell at "
    End If
    
    Select Case m.nOrderType
        Case eTT_OrderType_Market
            strReturn = strReturn & "MARKET"
        Case eTT_OrderType_Stop
            strReturn = strReturn & EnglishFromCoded(m.strPriceFormatted) & " STOP"
        Case eTT_OrderType_Limit
            strReturn = strReturn & EnglishFromCoded(m.strPriceFormatted) & " LIMIT"
        Case eTT_OrderType_StopWithLimit
            strReturn = strReturn & EnglishFromCoded(m.strPriceFormatted) & " STOP with a " & EnglishFromCoded(m.strWithLimitFormatted) & " LIMIT"
    End Select
    
    strReturn = Replace(strReturn, "Market1", Chr(34) & m.strSymbol & ", " & m.strBarPeriod & Chr(34))
    
    If m.Inputs.Count > 0 Then
        Set astrInputs = New cGdArray
        astrInputs.Create eGDARRAY_Strings
        
        For lIndex = 1 To m.Inputs.Count
            astrInputs.Add m.Inputs(lIndex).Name & " = '" & m.Inputs(lIndex).Value & "'"
        Next lIndex
        
        strReturn = strReturn & " (" & astrInputs.JoinFields(",") & ")"
    End If

    ToolTip = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.Tooltip"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetInputValue
'' Description: Set the input value for the given input
'' Inputs:      Input
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetInputValue(ByVal tsInput As cTradeSenseOrderInput)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the inputs collection

    If Not tsInput Is Nothing Then
        strKey = tsInput.Key(False)
        If m.Inputs.Exists(strKey) Then
            m.Inputs(strKey).Value = tsInput.Value
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.SetInputValue"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllRtDataAvailable
'' Description: Determine if streaming data is available for all symbols
'' Inputs:      None
'' Returns:     True if streaming data available, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllRtDataAvailable(Optional bExprIsNothing As Boolean = False) As Boolean
On Error GoTo ErrSection:

    If (m.bRtDataAvailable = False) And (Not m.Expr Is Nothing) Then
        m.bRtDataAvailable = m.Expr.AllRtDataAvailable
    End If
    
    bExprIsNothing = (m.Expr Is Nothing)
    AllRtDataAvailable = m.bRtDataAvailable
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.AllRtDataAvailable"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SessionDate
'' Description: Get the current session date for the expression bars
'' Inputs:      None
'' Returns:     Session Date (kNullData if invalid)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SessionDate() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    If m.Expr Is Nothing Then
        lReturn = kNullData
    Else
        lReturn = m.Expr.SessionDate
    End If
    
    SessionDate = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.SessionDate"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPrices
'' Description: Return a string of the current prices
'' Inputs:      None
'' Returns:     String version of current prices
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentPrices() As String
On Error GoTo ErrSection:

    CurrentPrices = m.Expr.CurrentPrices

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.CurrentPrices"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunExpressions
'' Description: Force the expressions to be run
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RunExpressions()
On Error GoTo ErrSection:

    If Not m.Expr Is Nothing Then
        m.Expr.RunExpression
        SetValues
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.RunExpressions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReloadData
'' Description: Reload the data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReloadData()
On Error GoTo ErrSection:

    If Not m.Expr Is Nothing Then
        m.Expr.ReloadData
        g.Broker.BrokerDebug m.nBroker, m.strLogText & "Data Reloaded - " & m.Expr.CurrentPrices
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.ReloadData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the exit strategy
'' Inputs:      None
'' Returns:     New Copy
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cTradeSenseOrder
On Error GoTo ErrSection:

    Dim aCopy As New cTradeSenseOrder
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.MakeCopy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private properties for a copy of the exit strategy
'' Inputs:      Copy From
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    m = mCopyFrom
    
    Set m.Rule = mCopyFrom.Rule.MakeCopy
    
    If mCopyFrom.Expr Is Nothing Then
        Set m.Expr = Nothing
    Else
        Set m.Expr = mCopyFrom.Expr.MakeCopy
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Activate
'' Description: Activate the Trade Sense order
'' Inputs:      Symbol, Account Position, Fills, Broker, Log Text, Custom Start Time,
''              Custom Stop Time, Auto Breakout Period, Lot Size
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Activate(ByVal strSymbol As String, ByVal AcctPos As cAccountPosition, ByVal Fills As cGdTree, Optional ByVal nBroker As eTT_AccountType = -1&, Optional ByVal strLogText As String = "", Optional ByVal dStartTime As Double = kNullData, Optional ByVal dStopTime As Double = kNullData, Optional ByVal lLotSize As Long = 1&) ', Optional ByVal strAutoBreakoutPeriod = "")
On Error GoTo ErrSection:

    Dim astrExpressions As cGdArray     ' Array of expressions
    Dim lIndex As Long                  ' Index into a for loop
    
    m.strSymbol = strSymbol
    m.nBroker = nBroker
    m.strLogText = strLogText
    'm.strAutoBreakoutPeriod = strAutoBreakoutPeriod
    SetBarProperties m.Bars, strSymbol
    
    Set astrExpressions = New cGdArray
    astrExpressions.Create eGDARRAY_Strings, 3
    astrExpressions(0) = m.strConditionCoded
    astrExpressions(1) = m.strPriceCoded
    astrExpressions(2) = m.strWithLimitCoded
    
    SetExpressionParms astrExpressions
    Set m.Expr = New cRunExpression
    
    If Fills Is Nothing Then
        m.Expr.InitMultipleWithAcctPos astrExpressions, "TsOrder", m.strSymbol, m.strBarPeriod, NumBarsRequired, AcctPos, dStartTime, dStopTime, m.bIntraBar, lLotSize
    Else
        m.Expr.InitMultipleWithFills astrExpressions, "TsOrder", m.strSymbol, m.strBarPeriod, NumBarsRequired, Fills, dStartTime, dStopTime, m.bIntraBar, lLotSize
    End If
    
    SetValues
    
    If m.bIntraBar Then
        m.nRtMode = eGDRtMode_On
    Else
        m.nRtMode = eGDRtMode_Off
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.Activate"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetValues
'' Description: Set the price values from the expression
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetValues()
On Error GoTo ErrSection:

    Dim dCondValue As Double            ' Last condition value

    If Not m.Expr Is Nothing Then
        dCondValue = m.Expr.LastValue(0)
        m.bConditionTrue = (dCondValue <> kNullData) And (dCondValue <> 0)
        If m.bConditionTrue Then
            m.dPrice = m.Expr.LastValue(1)
            If m.dPrice = kNullData Then
                m.dPrice = 0#
            End If
            
            m.dWithLimitPrice = m.Expr.LastValue(2)
            If m.dWithLimitPrice = kNullData Then
                m.dWithLimitPrice = 0#
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.SetValues"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetExpressionParms
'' Description: Set the parameters in the expression object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetExpressionParms(astrExpressions As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim tsInput As cTradeSenseOrderInput ' TradeSense order input object
    Dim strToReplace As String          ' String to replace
    Dim strReplaceWith As String        ' String to replace with
    Dim strToReplace2 As String         ' String to replace
    Dim strReplaceWith2 As String       ' String to replace with
    Dim lIndex2 As Long                 ' Index into a for loop
    
    For lIndex = 1 To m.Inputs.Count
        Set tsInput = m.Inputs(lIndex)
        
        strToReplace = ""
        strReplaceWith = ""
        
        Select Case tsInput.ParmType
            Case kSN_RetNumeric, kSN_RetNumericConstant
                strToReplace = "~27" & Format(Len(tsInput.Name), "000") & tsInput.Name
                strReplaceWith = "~13" & Format(Len(tsInput.Value), "000") & tsInput.Value
                
                strToReplace2 = "~05" & Format(Len(tsInput.Name), "000") & tsInput.Name
                strReplaceWith2 = "~13" & Format(Len(tsInput.Value), "000") & tsInput.Value
            
            Case kSN_RetTrueFalse, kSN_RetTrueFalseConstant
                strToReplace = "~28" & Format(Len(tsInput.Name), "000") & tsInput.Name
                If UCase(tsInput.Value) = "TRUE" Then
                    strReplaceWith = "~03004True ~16001( ~17001)"
                Else
                    strReplaceWith = "~03005False ~16001( ~17001)"
                End If
                
                strToReplace2 = "~06" & Format(Len(tsInput.Name), "000") & tsInput.Name
                If UCase(tsInput.Value) = "TRUE" Then
                    strReplaceWith2 = "~03004True ~16001( ~17001)"
                Else
                    strReplaceWith2 = "~03005False ~16001( ~17001)"
                End If
        
        End Select
        
        If (Len(strToReplace) > 0) And (Len(strReplaceWith) > 0) Then
            For lIndex2 = 0 To astrExpressions.Size - 1
                astrExpressions(lIndex2) = Replace(astrExpressions(lIndex2), strToReplace, strReplaceWith)
                astrExpressions(lIndex2) = Replace(astrExpressions(lIndex2), strToReplace2, strReplaceWith2)
            Next lIndex2
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.SetExpressionParms"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NumBarsRequired
'' Description: Number of bars required
'' Inputs:      None
'' Returns:     Number of bars
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NumBarsRequired() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    If m.bOverride Then
        lReturn = m.lNumBarsOver
    Else
        lReturn = m.lNumBarsAuto
    End If
    
    NumBarsRequired = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.NumBarsRequired"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcExpiration
'' Description: Calculate the expiration date for the order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcExpiration()
On Error GoTo ErrSection:

    Dim lSessionDate As Long            ' Current session date
    Dim lEndTime As Long                ' Ending time for the symbol
    Dim Bars As cGdBars                 ' Bars object for properties
    Dim dExpiration As Double           ' Expiration date/time
    Dim dCrossOver As Double            ' Crossover date/time
    Dim dFeedTime As Double             ' Feed time

    If (AllRtDataAvailable = True) And (Not m.Expr Is Nothing) Then
        Set Bars = New cGdBars
        If SetBarProperties(Bars, m.strSymbol) Then
            lSessionDate = m.Expr.SessionDate
            If lSessionDate <> kNullData Then
                lEndTime = Bars.Prop(eBARS_DefaultEndTime)
                
                dCrossOver = ConvertTimeZone(CDbl(lSessionDate) + CDbl((lEndTime + 5) / 1440#), Bars.Prop(eBARS_ExchangeTimeZoneInf))
                dExpiration = ConvertTimeZone(CDbl(lSessionDate) + CDbl((lEndTime - 1) / 1440#), Bars.Prop(eBARS_ExchangeTimeZoneInf))
                dFeedTime = g.RealTime.FeedTime(m.strSymbol)
                
                ' If we are more than 5 minutes after the end time of the session, bump the
                ' expiration to the next session close...
                If (dExpiration < dFeedTime) And (dFeedTime > dCrossOver) Then
                    dExpiration = dExpiration + 1
                    Do While Not IsWeekday(dExpiration)
                        dExpiration = dExpiration + 1
                    Loop
                End If
                
                m.dExpiration = dExpiration
                
                g.Broker.BrokerDebug m.nBroker, m.strLogText & "Order " & Str(m.lOrderNumber) & " expiration set to " & DateFormat(m.dExpiration, MM_DD_YY, HH_MM_SS, AMPM_UPPER)
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.CalcExpiration"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BoolToString
'' Description: Convert a boolean value to a string
'' Inputs:      Boolean Value
'' Returns:     String Value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BoolToString(ByVal bValue As Boolean) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    If bValue = True Then
        strReturn = "1"
    Else
        strReturn = "0"
    End If
    
    BoolToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrder.BoolToString"
    
End Function

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

    m.bBuy = True
    m.bEntry = True
    m.nOrderType = eTT_OrderType_Market
    m.strBarPeriod = "Daily"
    m.bCustom = True
    m.bUse = True
    m.lExitPercent = 100
    m.bIntraBar = False
    m.bRtDataAvailable = False
    m.bExpireDay = False
    m.dExpiration = kNullData
    m.strRequiredMod = ""
    m.nBroker = -1&
    m.bMarketIfWrongSide = False
    m.lCurrentSession = 0&
    
    Set m.Rule = New cRule
    Set m.Expr = Nothing
    Set m.Bars = New cGdBars
    
    Set m.alOCOs = New cGdArray
    m.alOCOs.Create eGDARRAY_Longs
    Set m.Inputs = New cTradeSenseOrderInputs

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.Class_Initialize"
    
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

    Set m.Rule = Nothing
    Set m.Expr = Nothing
    Set m.Bars = Nothing
    
    Set m.alOCOs = Nothing
    Set m.Inputs = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrder.Class_Terminate"
    
End Sub
