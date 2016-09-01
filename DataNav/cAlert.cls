VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAlert.cls
'' Description: Class for handling quote list alerts
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/07/2009   DAJ         Changed delimiter for order info from | to ;
'' 09/01/2009   DAJ         Use new Parked order status
'' 01/12/2010   DAJ         Added num bars required for Trade Sense alerts
'' 03/11/2010   DAJ         Use global collections
'' 09/16/2010   DAJ         More information in e-mail order alerts (#5891)
'' 09/22/2010   DAJ         If custom text supplied, show in addition to normal text (#5891)
'' 03/07/2011   DAJ         Broker Disconnect Alerts
'' 10/11/2011   DAJ         When alert is keep active, only alert once per bar (for data alerts)
'' 10/13/2011   DAJ         When sending a message to message history, use custom text if applicable
'' 10/27/2011   DAJ         Send correct symbol to alert message form for an order alert
'' 02/14/2012   DAJ         New status alerts for position mismatch / auto trade disabled
'' 04/03/2013   DAJ         Automated Strategy Baskets
'' 09/10/2013   DAJ         Handle a blank symbol in a secondary market for TS Alert ( e.g. ",2 minute" )
'' 01/14/2014   DAJ         Added 'Order Rejected' alert
'' 12/02/2015   DAJ         Added Account filter for "Any Orders" alert
'' 12/29/2015   DAJ         Fix for English string being incorrect on order alert for an automated trading item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kstrPositionMismatch = "Position Mismatch"
Private Const kstrAutoTradeItemDisabled = "Automated Trading Item Disabled"

Public Enum eGDAlertType
    eGDAlertType_QuoteBoard = 0
    eGDAlertType_AutoTrade
    eGDAlertType_Status
    eGDAlertType_Price
    eGDAlertType_Time
    eGDAlertType_Chart
    eGDAlertType_Annot
    eGDAlertType_TradeSense
    eGDAlertType_Undefined = -1
End Enum

Public Enum eGDAlertAction
    eAA_MessageBox = 0
    eAA_LogToFile
    eAA_ChangeBackColor
    eAA_SendPage
    eAA_SendEmail
    eAA_PlaceOrder
    eAA_PlaySound
    eAA_MsgHistory
    eAA_NumActions
End Enum

Public Enum eGDOrderAlertType
    eGDOrderAlertType_ManualOrder = 0
    eGDOrderAlertType_AutoTrade
    eGDOrderAlertType_AllOrders
End Enum

Public Enum eGDAlertBarOption
    eGDAlertBar_HiLowThisBar = 0    'high/low current bar: DownTo price, UpTo price, annot
    eGDAlertBar1_DownToClose        'close current bar: DownTo price, high/low current bar: UpTo price
    eGDAlertBar2_UpToClose          'close current bar: UpTo price, high/low current bar: DownTo price
    eGDAlertBar3_CloseThisBar       'close current bar: DownTo price, UpTo price, annot
    eGDAlertBar4_HiLowCompleteBar   'high/low last complete bar: annot
    eGDAlertBar5_CloseCompleteBar   'close last complete bar: annot
    eGDAlertBar9_Undefined = -1
End Enum

Public Enum eGDPriceConditionTrue
    eGDConditionTrue_None = 0
    eGDConditionTrue_UpTo
    eGDConditionTrue_DownTo
End Enum

Private Type mPrivate
    nAlertType As eGDAlertType          ' Alert Type
    
    ' Quote Board Alert Condition variables...
    bSymbol As Boolean                  ' Symbol or Tab alert?
    lSymbolID As Long                   ' Symbol ID of the symbol to check
    strSymbol As String                 ' Symbol to check
    strPeriod As String                 ' Period of the symbol to check
    strTabName As String                ' Name of the tab for the alert
    strField As String                  ' Field to check
    strOperator As String               ' Comparison operator to perform
    dValue As Double                    ' Value to compare against
    strCriteriaID As String             ' ID of the Criteria Field
    
    ' Automated Trading/Manual Order Alert Condition variables...
    nOrderAlert As eGDOrderAlertType    ' Type for the alert
    lTradeItemID As Long                ' ID for the automated trading item
    strAutoTradeCondition As String     ' Condition for the automated trading alert
    
    ' Status Alert Condition variables...
    strStatusItem As String             ' Status Item
    nBroker As eTT_AccountType          ' Broker for a status alert
    
    ' Time Alert Condition variables...
    bWeekday As Boolean                 ' Every specified day(s)? (or At a specific time)
    strDayMask As String                ' Mask as to which weekdays are specified
    dTriggerDateTime As Double          ' Date/Time (or Time) to trigger alert
    strTimeZone As String               ' ""=Local, "GMT"=GMT, "NY"=NY
    
    ' Price Alert Condition variables...
    aCharts As cGdArray                 'array of chart objects to show bell alert icon on
    bUseGetsDownTo As Boolean           ' Using the Gets Down To price?
    dGetsDownTo As Double               ' Gets Down To price
    bUseGetsUpTo As Boolean             ' Using the Gets Up To price?
    dGetsUpTo As Double                 ' Gets Up To price
    strPriceCondition As String         ' TradeSense expression to be evaluated
    bShowOnCharts As Boolean            'flag to show or not show on charts
    bFormsChecked As Boolean            'flag of whether or not list of loaded charts have been checked
    
    ' Chart Alert Condition variables...
    Annot As cAnnotation                ' Annotation the alert is tied to
    Ind As cIndicator                   ' Indicator the alert is tied to
    strChartCondition As String         ' Condition for the alert
    dChartAlertId As Double             ' Chart Alert ID
    
    ' TradeSense Alert Condition variables...
    bAutoDetect As Boolean              ' Auto detect the number of bars required?
    lNumBarsCalc As Long                ' Number of bars required -- calculated
    lNumBarsOver As Long                ' Number of bars required -- override
    
    ' Alert Properties...
    bActive As Boolean                  ' Is the alert currently active?
    bDeactivate As Boolean              ' Deactivate after alert is hit?
    bAsTradingUnits As Boolean          ' Display the value in trading units?
    bLastCheckedStatus As Boolean       'Was alert true or false when last checked?
    bRepeatPlay As Boolean              'Play sound repeatedly until user cancel?
    strKey As String                    'key for locating alert in alerts collection
    eAlertBarFlag As eGDAlertBarOption  'valid only for price and annot alerts
    
    ' Alert Actions...
    astrActions As cGdArray             ' Array of actions for the alert
    astrQBSymTrue As cGdArray           'Array of symbolID;Period that were true when last checked (valid only for QB tab alerts)
    
    ' Working variables...
    dNextTrigger As Double              ' Next trigger time for a time alert
    strCodedText As String              ' Coded text for the condition
    dLastCheckedTime As Double          'time the alert last checked (bars time if available, else feed time if realtime active, else NY time
    strLastCheckedTimeZone As String
    ePriceConditionTrue As eGDPriceConditionTrue    'for price alerts that have both UpTo and DownTo price selected
    bEditInProg As Boolean              'user is editing this alert in alert editor (5640)
    alLastBarAlerted As cGdArray        ' Last bar the data was on when user alerted
    astrSymbols As cGdArray             ' List of symbol(s);Period(s) for the alert
    
    Bars As cGdBars                     'to keep from having to search for bars in alerts collection
    aQbCells As cGdArray                'array of box-style QB Cell objects to notify when going active or inactive

    bNeverDeactivate As Boolean         'special flag for Pete

End Type
Private m As mPrivate

Public Property Get AlertType() As eGDAlertType
    AlertType = m.nAlertType
End Property
Public Property Let AlertType(ByVal nAlertType As eGDAlertType)
    m.nAlertType = nAlertType
End Property

Public Property Get IsSymbol() As Boolean
    IsSymbol = m.bSymbol
End Property
Public Property Let IsSymbol(ByVal bIsSymbol As Boolean)
    m.bSymbol = bIsSymbol
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Let Symbol(ByVal strSymbol As String)
    m.strSymbol = strSymbol
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property
Public Property Let SymbolID(ByVal lSymbolID As Long)
    m.lSymbolID = lSymbolID
End Property

Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID = 0 Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If
End Property
Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
End Property

Public Property Get Period() As String
    Period = m.strPeriod
End Property
Public Property Let Period(ByVal strPeriod As String)
    m.strPeriod = strPeriod
End Property

Public Property Get Field() As String
    Field = m.strField
End Property
Public Property Let Field(ByVal strField As String)
    m.strField = strField
End Property

Public Property Get TabName() As String
    TabName = m.strTabName
End Property
Public Property Let TabName(ByVal strTabName As String)
    m.strTabName = strTabName
End Property

Public Property Get Operator() As String
    Operator = m.strOperator
End Property
Public Property Let Operator(ByVal strOperator As String)
    m.strOperator = Parse(strOperator, " ", 1)
End Property

Public Property Get Value() As Double
    Value = m.dValue
End Property
Public Property Let Value(ByVal dValue As Double)
    m.dValue = dValue
End Property

Public Property Get Active() As Boolean
    Active = m.bActive
End Property

Public Property Let Active(ByVal bActive As Boolean)
On Error GoTo ErrSection:
        
    Dim i&
        
    If bActive <> m.bActive Then
        m.bActive = bActive
        If Not g.bStarting And Not g.bUnloading Then
            For i = 0 To m.aQbCells.Size - 1
                If Not m.aQbCells(i) Is Nothing Then m.aQbCells(i).QbCellAlertUpdate Me, -1
            Next
            
            If FormIsLoaded("frmAlertMessages") Then frmAlertMessages.AlertActiveChanged Me
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAlert.Active.Let"
    
End Property

Public Property Get Deactivate() As Boolean
    Deactivate = m.bDeactivate
End Property

Public Property Let Deactivate(ByVal bDeactivate As Boolean)
    If m.nAlertType = eGDAlertType_AutoTrade Or m.nAlertType = eGDAlertType_Status Then
        'never let these types of alerts get deactivated after first triggered
        m.bDeactivate = False
    Else
        m.bDeactivate = bDeactivate
    End If
End Property

Public Property Get AsTradingUnits() As Boolean
    AsTradingUnits = m.bAsTradingUnits
End Property
Public Property Let AsTradingUnits(ByVal bTU As Boolean)
    m.bAsTradingUnits = bTU
End Property

Public Property Get ActionString(ByVal nIndex As eGDAlertAction) As String
    ActionString = m.astrActions(nIndex)
End Property
Public Property Let ActionString(ByVal nIndex As eGDAlertAction, ByVal strAction As String)
    m.astrActions(nIndex) = strAction
End Property

Public Property Get CriteriaID() As String
    CriteriaID = m.strCriteriaID
End Property
Public Property Let CriteriaID(ByVal strCriteriaID As String)
    m.strCriteriaID = strCriteriaID
End Property

Public Property Get OrderAlertType() As eGDOrderAlertType
    OrderAlertType = m.nOrderAlert
End Property
Public Property Let OrderAlertType(ByVal nOrderAlert As eGDOrderAlertType)
    m.nOrderAlert = nOrderAlert
End Property

Public Property Get TradeItemID() As Long
    TradeItemID = m.lTradeItemID
End Property
Public Property Let TradeItemID(ByVal lTradeItemID As Long)
    m.lTradeItemID = lTradeItemID
End Property

Public Property Get AutoTradeCondition() As String
    AutoTradeCondition = m.strAutoTradeCondition
End Property
Public Property Let AutoTradeCondition(ByVal strAutoTradeCondition As String)
    m.strAutoTradeCondition = strAutoTradeCondition
End Property

Public Property Get StatusItem() As String
    StatusItem = m.strStatusItem
End Property
Public Property Let StatusItem(ByVal strStatusItem As String)
    m.strStatusItem = strStatusItem
End Property

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property
Public Property Let Broker(ByVal nBroker As eTT_AccountType)
    m.nBroker = nBroker
End Property

Public Property Get Weekday() As Boolean
    Weekday = m.bWeekday
End Property
Public Property Let Weekday(ByVal bWeekday As Boolean)
    m.bWeekday = bWeekday
End Property

Public Property Get WeekdayMask() As String
    WeekdayMask = m.strDayMask
End Property
Public Property Let WeekdayMask(ByVal strWeekdayMask As String)
    m.strDayMask = strWeekdayMask
End Property

Public Property Get TriggerTime() As Double
    TriggerTime = m.dTriggerDateTime
End Property
Public Property Let TriggerTime(ByVal dTriggerTime As Double)
    m.dTriggerDateTime = dTriggerTime
End Property

Public Property Get TimeZone() As String
    TimeZone = m.strTimeZone
End Property
Public Property Let TimeZone(ByVal strTimeZone As String)
    m.strTimeZone = strTimeZone
End Property

Public Property Get UseGetsDownTo() As Boolean
    UseGetsDownTo = m.bUseGetsDownTo
End Property
Public Property Let UseGetsDownTo(ByVal bUseGetsDownTo As Boolean)
    m.bUseGetsDownTo = bUseGetsDownTo
End Property

Public Property Get GetsDownToPrice() As Double
    GetsDownToPrice = m.dGetsDownTo
End Property
Public Property Let GetsDownToPrice(ByVal dGetsDownToPrice As Double)
    m.dGetsDownTo = dGetsDownToPrice
End Property

Public Property Get UseGetsUpTo() As Boolean
    UseGetsUpTo = m.bUseGetsUpTo
End Property
Public Property Let UseGetsUpTo(ByVal bUseGetsUpTo As Boolean)
    m.bUseGetsUpTo = bUseGetsUpTo
End Property

Public Property Get GetsUpToPrice() As Double
    GetsUpToPrice = m.dGetsUpTo
End Property
Public Property Let GetsUpToPrice(ByVal dGetsUpToPrice As Double)
    m.dGetsUpTo = dGetsUpToPrice
End Property

Public Property Get PriceCondition() As String
    PriceCondition = m.strPriceCondition
End Property
Public Property Let PriceCondition(ByVal strPriceCondition As String)
    m.strPriceCondition = strPriceCondition

    m.strCodedText = ""
    m.strCodedText = CodedText
End Property

Public Property Get ChartCondition() As String
    ChartCondition = m.strChartCondition
End Property
Public Property Let ChartCondition(ByVal strCondition As String)
    m.strChartCondition = strCondition
End Property

Public Property Get Annotation() As cAnnotation
    Set Annotation = m.Annot
End Property
Public Property Let Annotation(Annot As cAnnotation)
    Set m.Annot = Annot
End Property

Public Property Get Indicator() As cIndicator
    Set Indicator = m.Ind
End Property
Public Property Let Indicator(Ind As cIndicator)
    Set m.Ind = Ind
End Property

Public Property Get ChartAlertId() As Double
    ChartAlertId = m.dChartAlertId
End Property
Public Property Let ChartAlertId(ByVal dId As Double)
    m.dChartAlertId = dId
End Property

Public Property Let AlertKey(ByVal strKey As String)
    m.strKey = strKey
End Property

Public Property Get AlertKey() As String
    AlertKey = m.strKey
End Property

Public Property Get RepeatPlay() As Boolean
    RepeatPlay = m.bRepeatPlay
End Property

Public Property Let RepeatPlay(ByVal bRepeat As Boolean)
    m.bRepeatPlay = bRepeat
End Property

Public Property Get AutoDetect() As Boolean
    AutoDetect = m.bAutoDetect
End Property
Public Property Let AutoDetect(ByVal bAutoDetect As Boolean)
    m.bAutoDetect = bAutoDetect
End Property

Public Property Get NumBarsCalc() As Long
    NumBarsCalc = m.lNumBarsCalc
End Property
Public Property Let NumBarsCalc(ByVal lNumBarsCalc As Long)
    m.lNumBarsCalc = lNumBarsCalc
End Property

Public Property Get NumBarsOver() As Long
    NumBarsOver = m.lNumBarsOver
End Property
Public Property Let NumBarsOver(ByVal lNumBarsOver As Long)
    m.lNumBarsOver = lNumBarsOver
End Property

Public Property Get NumBars() As Long
    If (m.bAutoDetect = True) Then
        NumBars = m.lNumBarsCalc
    Else
        NumBars = m.lNumBarsOver
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToFileString
'' Description: Get the string ready to put into an output file
'' Inputs:      None
'' Returns:     String to store into an output file
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToFileString() As String
On Error GoTo ErrSection:

    Dim strSymbolOrSymbolID As String   ' Symbol or symbol id to save to the string
    Dim aSoundParams As New cGdArray

    aSoundParams.SplitFields m.astrActions(eAA_PlaySound), ","
    aSoundParams(2) = Abs(m.bRepeatPlay)
    m.astrActions(eAA_PlaySound) = aSoundParams.JoinFields(",")
    
    Select Case m.nAlertType
        Case eGDAlertType_QuoteBoard
            If Len(m.strTabName) > 0 Then
                strSymbolOrSymbolID = ""
                m.strPeriod = ""
            ElseIf m.lSymbolID = 0 Then
                strSymbolOrSymbolID = m.strSymbol
            Else
                strSymbolOrSymbolID = Str(m.lSymbolID)
            End If
        
            ToFileString = Str(m.nAlertType) & "|" & Str(m.bSymbol) & "|" _
                & strSymbolOrSymbolID & ";" & m.strPeriod & "|" & m.strField & "|" & m.strTabName & "|" _
                & m.strOperator & "|" & Str(m.dValue) & "|" & Str(m.bActive) & "|" & Str(m.bDeactivate) & "|" _
                & Str(m.bAsTradingUnits) & "|" & m.strCriteriaID & "|" & m.astrActions.JoinFields("~")
                        
        Case eGDAlertType_AutoTrade
            ToFileString = Str(m.nAlertType) & "|" & Str(m.nOrderAlert) & "|" & Str(m.lTradeItemID) & "|" & _
                m.strAutoTradeCondition & "|" & Str(m.bActive) & "|" & Str(m.bDeactivate) & "|" & _
                m.astrActions.JoinFields("~")
    
        Case eGDAlertType_Status
            ToFileString = Str(m.nAlertType) & "|" & m.strStatusItem & ";" & Str(m.nBroker) & "|" _
                & Str(m.bActive) & "|" & Str(m.bDeactivate) & "|" & m.astrActions.JoinFields("~")
                
        Case eGDAlertType_Time
            ToFileString = Str(m.nAlertType) & "|" & Str(m.bWeekday) & "|" & m.strDayMask & "|" _
                & Str(m.dTriggerDateTime) & "|" & m.strTimeZone & "|" _
                & Str(m.bActive) & "|" & Str(m.bDeactivate) & "|" & m.astrActions.JoinFields("~")
                
        Case eGDAlertType_Price
            If m.lSymbolID = 0 Then strSymbolOrSymbolID = m.strSymbol Else strSymbolOrSymbolID = Str(m.lSymbolID)
            
            ToFileString = Str(m.nAlertType) & "|" & strSymbolOrSymbolID & "|" & m.strPeriod & "|" _
                & Str(m.bUseGetsDownTo) & "|" & Str(m.dGetsDownTo) & "|" & Str(m.bUseGetsUpTo) & "|" _
                & Str(m.dGetsUpTo) & "|" & Str(m.bActive) & "|" & Str(m.bDeactivate) & "|" _
                & m.astrActions.JoinFields("~")
        
        Case eGDAlertType_TradeSense
            If m.lSymbolID = 0 Then strSymbolOrSymbolID = m.strSymbol Else strSymbolOrSymbolID = Str(m.lSymbolID)
            
            ToFileString = Str(m.nAlertType) & "|" & strSymbolOrSymbolID & "|" & m.strPeriod & "|" _
                & m.strPriceCondition & "|" & Str(m.bActive) & "|" & Str(m.bDeactivate) & "|" _
                & m.astrActions.JoinFields("~")
    
        Case eGDAlertType_Annot, eGDAlertType_Chart
            ToFileString = Str(m.nAlertType) & "|" & m.strSymbol & "|" & m.lSymbolID & "|" _
                & m.strChartCondition & "|" & m.strOperator & "|" & m.strField & "|" _
                & Str(m.dValue) & "|" & Str(m.bActive) & "|" & Str(m.bDeactivate) & "|" _
                & m.astrActions.JoinFields("~")
    End Select

    If m.astrQBSymTrue.Size > 0 Then
        ToFileString = ToFileString & "|" & m.astrQBSymTrue.JoinFields("!")     'list of symbols that were true when last checked for QB tab
    Else
        ToFileString = ToFileString & "|" & Abs(m.bLastCheckedStatus)           'last checked status (0 or 1)
    End If
    
    'last checked time and time zone of last checked time
    'this is to prevent alert from going to TRUE when user exit TN with realtime on then restart TN and data time is less than reatime data
    ToFileString = ToFileString & "|" & Str(m.dLastCheckedTime) & "," & m.strLastCheckedTimeZone
        
    Select Case m.nAlertType
        Case eGDAlertType_Price
            'flag of whether to compare alert value against close for price alert
            ToFileString = ToFileString & "|" & Str(ValidateBarFlag(m.eAlertBarFlag))
            'flag of whether to show alert on charts
            ToFileString = ToFileString & "|" & Abs(m.bShowOnCharts)
        
        Case eGDAlertType_Annot
            'flag of whether to compare alert value against close for price alert
            ToFileString = ToFileString & "|" & Str(ValidateBarFlag(m.eAlertBarFlag))
            
        Case eGDAlertType_TradeSense
            ToFileString = ToFileString & "|" & Str(m.bAutoDetect) & "," & Str(m.lNumBarsCalc) & "," _
                & Str(m.lNumBarsOver)
    End Select
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.ToFileString", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromFileString
'' Description: Load the class members from a string brought in from an input
''              file
'' Inputs:      String read in from input file
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromFileString(ByVal strFileString As String)
On Error GoTo ErrSection:

    Dim strTemp As String               ' Temporary string
    Dim astrAlert As New cGdArray       ' Array of alert properties
    Dim bValid As Boolean
    Dim astrAutoDetect As New cGdArray  ' Auto detect information for TradeSense alert

    ' TLB 9/20/2012: rename "LW Sentiment" functions
    strFileString = Replace(strFileString, "LW Sentiment", "TN Consensus")

    astrAlert.SplitFields strFileString, "|"

    m.nAlertType = CLng(Val(astrAlert(0)))
    Select Case m.nAlertType
        Case eGDAlertType_QuoteBoard
            bValid = True       'reset
            m.bSymbol = CBool(astrAlert(1))
            strTemp = Parse(astrAlert(2), ";", 1)
            If IsDigit(strTemp) Then
                m.lSymbolID = Val(strTemp)
                m.strSymbol = g.SymbolPool.SymbolForID(m.lSymbolID)         '40942 does not yield a valid symbol (HU)
            Else
                m.strSymbol = strTemp
                m.lSymbolID = g.SymbolPool.SymbolIDforSymbol(m.strSymbol)   'ZE-057 does not yield a valid symbolID
            End If
            
            If m.bSymbol Then
                If Len(m.strSymbol) = 0 Or m.lSymbolID = 0 Then
                    bValid = False
                End If
            End If
            
            If bValid Then
                m.strPeriod = Parse(astrAlert(2), ";", 2)
                If Len(m.strPeriod) = 0 Then m.strPeriod = "Daily"
                m.strField = astrAlert(3)
                m.strTabName = astrAlert(4)
                m.strOperator = astrAlert(5)
                Select Case UCase(m.strOperator)
                    Case "EQUAL"
                        m.strOperator = "="
                    Case "LESS THAN"
                        m.strOperator = "<"
                    Case "GREATER THAN"
                        m.strOperator = ">"
                    Case "LESS THAN OR EQUAL"
                        m.strOperator = "<="
                    Case "GREATER THAN OR EQUAL"
                        m.strOperator = ">="
                End Select
                m.dValue = Val(astrAlert(6))
                If Len(astrAlert(7)) > 0 Then m.bActive = CBool(astrAlert(7)) Else m.bActive = True
                If Len(astrAlert(8)) > 0 Then m.bDeactivate = CBool(astrAlert(8)) Else m.bDeactivate = True
                If Len(astrAlert(9)) > 0 Then m.bAsTradingUnits = CBool(astrAlert(9)) Else m.bAsTradingUnits = True
                m.strCriteriaID = astrAlert(10)
                m.astrActions.SplitFields astrAlert(11), "~"
                If astrAlert.Size > 12 Then
                    If m.bSymbol Then
                        m.bLastCheckedStatus = astrAlert(12) * -1
                    Else
                        m.astrQBSymTrue.SplitFields astrAlert(12), "!"
                    End If
                End If
                If astrAlert.Size > 13 Then
                    m.dLastCheckedTime = ValOfText(Parse(astrAlert(13), ",", 1))
                    m.strLastCheckedTimeZone = Parse(astrAlert(13), ",", 2)
                End If
            End If
            
        Case eGDAlertType_AutoTrade
            If Len(astrAlert(1)) > 0 Then m.nOrderAlert = Abs(CLng(Val(astrAlert(1)))) Else m.nOrderAlert = eGDOrderAlertType_AllOrders
            m.lTradeItemID = CLng(Val(astrAlert(2)))
            m.strAutoTradeCondition = astrAlert(3)
            If Len(astrAlert(4)) > 0 Then m.bActive = CBool(astrAlert(4)) Else m.bActive = True
            If Len(astrAlert(5)) > 0 Then m.bDeactivate = CBool(astrAlert(5)) Else m.bDeactivate = True
            m.astrActions.SplitFields astrAlert(6), "~"
            If astrAlert.Size > 7 Then
                m.bLastCheckedStatus = astrAlert(7) * -1
            End If
            If astrAlert.Size > 8 Then
                m.dLastCheckedTime = ValOfText(Parse(astrAlert(8), ",", 1))
                m.strLastCheckedTimeZone = Parse(astrAlert(8), ",", 2)
            End If
            
        Case eGDAlertType_Status
            m.strStatusItem = Parse(astrAlert(1), ";", 1)
            m.nBroker = CLng(Val(Parse(astrAlert(1), ";", 2)))
            If Len(astrAlert(2)) > 0 Then m.bActive = CBool(astrAlert(2)) Else m.bActive = True
            If Len(astrAlert(3)) > 0 Then m.bDeactivate = CBool(astrAlert(3)) Else m.bDeactivate = True
            m.astrActions.SplitFields astrAlert(4), "~"
            If astrAlert.Size > 5 Then
                If InStr(UCase(m.strStatusItem), "STREAMING") > 0 Then
                    m.bLastCheckedStatus = False
                Else
                    m.bLastCheckedStatus = astrAlert(5) * -1
                End If
            End If
            If astrAlert.Size > 6 Then
                If InStr(UCase(m.strStatusItem), "STREAMING") > 0 Then
                    m.dLastCheckedTime = kNullData
                Else
                    m.dLastCheckedTime = ValOfText(Parse(astrAlert(6), ",", 1))
                End If
                m.strLastCheckedTimeZone = Parse(astrAlert(6), ",", 2)
            End If
            
        Case eGDAlertType_Time
            If Len(astrAlert(1)) > 0 Then m.bWeekday = CBool(astrAlert(1)) Else m.bWeekday = True
            m.strDayMask = astrAlert(2)
            m.dTriggerDateTime = Val(astrAlert(3))
            m.strTimeZone = astrAlert(4)
            If Len(astrAlert(5)) > 0 Then m.bActive = CBool(astrAlert(5)) Else m.bActive = True
            If Len(astrAlert(6)) > 0 Then m.bDeactivate = CBool(astrAlert(6)) Else m.bDeactivate = True
            m.astrActions.SplitFields astrAlert(7), "~"
            CalcNextTriggerTime
            If astrAlert.Size > 8 Then
                m.bLastCheckedStatus = astrAlert(8) * -1
            End If
            If astrAlert.Size > 9 Then
                m.dLastCheckedTime = ValOfText(Parse(astrAlert(9), ",", 1))
                m.strLastCheckedTimeZone = Parse(astrAlert(9), ",", 2)
            End If
            
        Case eGDAlertType_Price
            m.strSymbol = astrAlert(1)
            m.lSymbolID = Val(m.strSymbol)
            If m.lSymbolID <> 0 Then m.strSymbol = GetSymbol(m.lSymbolID)
            m.strPeriod = astrAlert(2)
            If Len(m.strPeriod) = 0 Then m.strPeriod = "Daily"
            If Len(astrAlert(3)) > 0 Then m.bUseGetsDownTo = CBool(astrAlert(3)) Else m.bUseGetsDownTo = False
            m.dGetsDownTo = Val(astrAlert(4))
            If Len(astrAlert(5)) > 0 Then m.bUseGetsUpTo = CBool(astrAlert(5)) Else m.bUseGetsUpTo = False
            m.dGetsUpTo = Val(astrAlert(6))
            If Len(astrAlert(7)) > 0 Then m.bActive = CBool(astrAlert(7)) Else m.bActive = True
            If Len(astrAlert(8)) > 0 Then m.bDeactivate = CBool(astrAlert(8)) Else m.bDeactivate = True
            m.astrActions.SplitFields astrAlert(9), "~"
            If astrAlert.Size > 10 Then
                m.bLastCheckedStatus = Val(astrAlert(10)) * -1
            End If
            If astrAlert.Size > 11 Then
                m.dLastCheckedTime = ValOfText(Parse(astrAlert(11), ",", 1))
                m.strLastCheckedTimeZone = Parse(astrAlert(11), ",", 2)
            End If
            If astrAlert.Size > 12 Then
                m.eAlertBarFlag = ValidateBarFlag(ValOfText(Parse(astrAlert(12), "|", 1)))
            End If
            If astrAlert.Size > 13 Then
                m.bShowOnCharts = Val(astrAlert(13)) * -1
            End If
        
        Case eGDAlertType_TradeSense
            m.strSymbol = astrAlert(1)
            m.lSymbolID = Val(m.strSymbol)
            If m.lSymbolID <> 0 Then m.strSymbol = GetSymbol(m.lSymbolID)
            m.strPeriod = astrAlert(2)
            If Len(m.strPeriod) = 0 Then m.strPeriod = "Daily"
            m.strPriceCondition = astrAlert(3)
            If Len(astrAlert(4)) > 0 Then m.bActive = CBool(astrAlert(4)) Else m.bActive = True
            If Len(astrAlert(5)) > 0 Then m.bDeactivate = CBool(astrAlert(5)) Else m.bDeactivate = True
            m.astrActions.SplitFields astrAlert(6), "~"
            If astrAlert.Size > 7 Then
                m.bLastCheckedStatus = astrAlert(7) * -1
            End If
            If astrAlert.Size > 8 Then
                m.dLastCheckedTime = ValOfText(Parse(astrAlert(8), ",", 1))
                m.strLastCheckedTimeZone = Parse(astrAlert(8), ",", 2)
            End If
            If astrAlert.Size > 9 Then
                astrAutoDetect.SplitFields astrAlert(9), ","
                If Len(astrAutoDetect(0)) > 0 Then m.bAutoDetect = CBool(astrAutoDetect(0)) Else m.bAutoDetect = True
                If Len(astrAutoDetect(1)) > 0 Then m.lNumBarsCalc = CLng(Val(astrAutoDetect(1))) Else m.lNumBarsCalc = -1&
                If Len(astrAutoDetect(2)) > 0 Then m.lNumBarsOver = CLng(Val(astrAutoDetect(2))) Else m.lNumBarsOver = -1&
            End If
            
            m.strCodedText = ""
            m.strCodedText = CodedText
            
        Case eGDAlertType_Annot, eGDAlertType_Chart
            m.strSymbol = astrAlert(1)
            m.lSymbolID = astrAlert(2)
            m.strChartCondition = astrAlert(3)
            m.strOperator = astrAlert(4)
            m.strField = astrAlert(5)
            m.dValue = astrAlert(6)
            m.bActive = astrAlert(7)
            m.bDeactivate = astrAlert(8)
            m.astrActions.SplitFields astrAlert(9), "~"
            If astrAlert.Size > 10 Then
                If astrAlert(10) = "1" Then
                    m.bLastCheckedStatus = True
                Else
                    m.bLastCheckedStatus = False
                End If
            End If
            If astrAlert.Size > 11 Then
                m.dLastCheckedTime = ValOfText(Parse(astrAlert(11), ",", 1))
                m.strLastCheckedTimeZone = Parse(astrAlert(11), ",", 2)
            End If
            If astrAlert.Size > 12 Then
                m.eAlertBarFlag = ValidateBarFlag(ValOfText(Parse(astrAlert(12), "|", 1)))
            End If
    
    End Select
           
    If ValOfText(Parse(m.astrActions(eAA_PlaySound), ",", 3)) = 1 Then
        m.bRepeatPlay = True
    Else
        m.bRepeatPlay = False
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAlert.FromFileString", eGDRaiseError_Raise
    
End Sub

Private Sub HandleAlertAction(ByVal bOnlyBackColor As Boolean, ByVal dAlertTime As Double, _
    ByVal strSymbol$, ByVal strPeriod$, ByVal strOrderInfo$, ByVal strCaption$, _
    Optional ByVal bConfirmOverride As Boolean = False, _
    Optional ByVal ePriceCondition As eGDPriceConditionTrue = eGDConditionTrue_None, _
    Optional ByVal bNewBar As Boolean = True, Optional ByVal StatusInfo As cStatusAlertInfo = Nothing)

    Dim lIndex As Long                  ' Index into a for loop
    
    Dim astrAction As New cGdArray      ' Array of properties for an alert action
    Dim strMsg As String                ' Message to display
    Dim strCustomMsg As String          ' custom message for message box
    Dim strTemp As String
    Dim fhLogFile As Integer            ' File handle to the log file

    Dim frm As Form                     ' Form to use to display message box
    
    Dim nComPort As Long                ' Com Port to use for paging
    Dim nWaitSeconds As Long            ' Number of seconds to wait
    Dim strPager As String              ' Pager phone number
    Dim strModemInit As String          ' Modem initialization string
    Dim strDate As String               ' Formatted date for the e-mail
    
    Dim bDoHistory As Boolean           ' Do the history message?
    Dim bChangeCellColor As Boolean     'was change-cell-color one of the action item?
    Dim dNow As Double                  ' Time right now in correct time zone
    Dim strAlertSymbol As String        ' Symbol to send to the alert message form
    
    Dim bRepeatInProg As Boolean
    
    
    If m.astrQBSymTrue.Size = 0 Then
        If m.nAlertType = eGDAlertType_AutoTrade Then
            If m.nOrderAlert = eGDOrderAlertType_ManualOrder Then
                bDoHistory = Not m.bLastCheckedStatus
            Else
                bDoHistory = True   'always perform alert action for orders
            End If
        Else
            bDoHistory = Not m.bLastCheckedStatus   'if last time checked was true then don't do history again
        End If
    Else
        bDoHistory = False
        m.astrQBSymTrue.Sort eGdSort_DeleteDuplicates
        strTemp = GetSymbolID(strSymbol) & ";" & strPeriod
        If m.astrQBSymTrue.BinarySearch(strTemp, lIndex) Then
            If lIndex >= 0 And lIndex < m.astrQBSymTrue.Size Then
                If m.astrQBSymTrue(lIndex) <> strTemp Then
                    bDoHistory = True
                End If
            End If
        Else
            bDoHistory = True
        End If
    End If
    
    For lIndex = 0 To m.astrActions.Size - 1
        astrAction.SplitFields m.astrActions(lIndex), ","
        If astrAction(0) <> "0" Then
            Select Case lIndex
                Case eAA_MessageBox
                    If bOnlyBackColor = False And m.bActive And bDoHistory And bNewBar Then
                        If Not FormIsLoaded("frmAlertMessages") Then frmAlertMessages.ShowMe
                        strCustomMsg = Trim(astrAction(1))
                    End If
                
                Case eAA_LogToFile
                    If bOnlyBackColor = False And m.bActive And bDoHistory And bNewBar Then
                        strMsg = Trim(astrAction(1))
                        If Len(strOrderInfo) > 0 Then
                            If Len(strMsg) = 0 Then
                                strMsg = Replace(strOrderInfo, ";", ", ")
                            Else
                                strMsg = Replace(strOrderInfo, ";", ", ") & ", Message : " & strMsg
                            End If
                        ElseIf Len(strMsg) = 0 Then
                            strMsg = EnglishString(False, strSymbol, strPeriod, ePriceCondition, StatusInfo)
                        End If
                        
                        strTemp = astrAction(2)
                        If Len(strTemp) > 0 Then
                            fhLogFile = FreeFile
                            Open strTemp For Append As #fhLogFile
                            Print #fhLogFile, DateFormat(Now, MM_DD_YYYY, HH_MM_SS, NO_AMPM) & ": " & strMsg
                            Close #fhLogFile
                        End If
                    End If
                
                Case eAA_ChangeBackColor
                    bChangeCellColor = True
                    If m.nAlertType = eGDAlertType_QuoteBoard Then
                        'If Len(m.strTabName) = 0 Or (frmQuotes.TabStr(eGDTabSettings_Name, frmQuotes.vsTab.CurrTab) = m.strTabName) Then
                        'JM - 02-09-2010: call this all the time so detached QB tabs can also be processed
                        frmQuotes.ColorCell strSymbol, strPeriod, m.strField, m.strTabName, Val(astrAction(1))
                    End If
                
                Case eAA_SendPage
                    If bOnlyBackColor = False And m.bActive And bDoHistory And bNewBar Then
                        strMsg = astrAction(1)
                        If Len(strOrderInfo) > 0 Then
                            If Len(strMsg) = 0 Then
                                strMsg = Replace(strOrderInfo, ";", ", ")
                            Else
                                strMsg = Replace(strOrderInfo, ";", ", ") & ", Message : " & strMsg
                            End If
                        ElseIf Len(strMsg) = 0 Then
                            strMsg = EnglishString(False, strSymbol, strPeriod, ePriceCondition, StatusInfo)
                        End If
                        
                        strPager = GetIniFileProperty("Pager", "", "PagerSettings", g.strIniFile)
                        nComPort = GetIniFileProperty("ComPort", 1, "PagerSettings", g.strIniFile)
                        nWaitSeconds = GetIniFileProperty("WaitSeconds", 10, "PagerSettings", g.strIniFile)
                        strModemInit = GetIniFileProperty("ModemInit", MODEM_INIT_DEFAULT, "PagerSettings", g.strIniFile)
                        If DialPager(nComPort, strPager, nWaitSeconds, strMsg, strModemInit, True) Then
                            strMsg = "Page sent:  " & strMsg
                        Else
                            strMsg = "ERROR sending page:  " & strMsg
                        End If
                        ' also show popup to indicate page was attempted
                        Set frm = New frmAlertPopup
                        ''frm.ShowMe m.strSymbol, strMsg
                        frm.ShowMe eGDAlertMode_QuoteBoardAlert, strCaption, strMsg
                    End If
                
                Case eAA_SendEmail
                    If bOnlyBackColor = False And m.bActive And bDoHistory And bNewBar Then
                        strDate = Format(Now, "d mmm yyyy") & " " & DateFormat(Now, NO_DATE, HH_MM_SS)
                        strMsg = astrAction(1)
                        If Len(strOrderInfo) > 0 Then
                            If Len(strMsg) = 0 Then
                                strMsg = Replace(strOrderInfo, ";", vbLf)
                            Else
                                strMsg = Replace(strOrderInfo, ";", vbLf) & vbLf & "Message : " & strMsg
                            End If
                        ElseIf Len(strMsg) = 0 Then
                            strMsg = EnglishString(False, strSymbol, strPeriod, ePriceCondition, StatusInfo)
                        End If
                        If Len(strMsg) > 0 Then SendEmail astrAction, strSymbol, strMsg, strDate
                    End If
                    
                Case eAA_PlaceOrder
                    If bOnlyBackColor = False And m.bActive And bDoHistory And bNewBar Then
                        Set frm = New frmAlertPopup
                        strTemp = Parse(astrAction(1), ",", 8)
                        If Len(strTemp) > 0 Then
                            strTemp = g.SymbolPool.SymbolForID(Val(strTemp))
                            If Len(strTemp) = 0 Then strTemp = strSymbol
                        Else
                            strTemp = strSymbol
                        End If
                        'do this right away to prevent double-ordering, which can occur in real-time with genesis simulated account
                        m.bLastCheckedStatus = True
                        'm.astrActions(eAA_PlaceOrder) = "0"      '4219
                        If Not m.bNeverDeactivate Then      'special flag for Pete
                            m.bActive = False
                            m.bDeactivate = True
                        End If
                        ToFileString
                        If Not m.Annot Is Nothing Then m.Annot.UpdateAlert 2
                        If astrAction.Size >= 3 And Not bConfirmOverride Then       '4307
                            frm.ShowMe eGDAlertMode_PlaceOrderAlert, strTemp, astrAction(1), 0, ValOfText(astrAction(2)), Me
                        Else
                            frm.ShowMe eGDAlertMode_PlaceOrderAlert, strTemp, astrAction(1), 0, , Me
                        End If
                    End If
                    
                Case eAA_PlaySound
                    If bOnlyBackColor = False And m.bActive And bDoHistory And bNewBar Then
                        If Len(astrAction(1)) > 0 Then
                            If m.bRepeatPlay Then
                                PlaySoundFile astrAction(1), False, True
                                bRepeatInProg = True
                            Else
                                frmMain.PlaySound astrAction(1)
                            End If
                        Else
                            Beep
                        End If
                    End If
                    
                Case eAA_MsgHistory
                    If bOnlyBackColor = False And bDoHistory And (m.bActive Or bConfirmOverride) And bNewBar Then
                        strAlertSymbol = strSymbol
                        
                        If Len(Trim(astrAction(1))) > 0 Then
                            strMsg = Trim(astrAction(1))
                        ElseIf Len(strCustomMsg) > 0 Then
                            strMsg = strCustomMsg
                        Else
                            strMsg = ""
                        End If
                        
                        If Len(strOrderInfo) > 0 Then
                            If Len(strMsg) = 0 Then
                                strMsg = Replace(strOrderInfo, ";", ", ")
                            Else
                                strMsg = Replace(strOrderInfo, ";", ", ") & ", Message : " & strMsg
                            End If
                            
                            strAlertSymbol = ExtractSymbolFromOrderInfo(strOrderInfo)
                        ElseIf Len(strMsg) = 0 Then
                            strMsg = EnglishString(False, strSymbol, strPeriod, ePriceCondition, StatusInfo)
                        End If
                        g.Alerts.AddAlertMessage dAlertTime, strMsg, Me, strAlertSymbol
                        If bRepeatInProg Then
                            geTaskBarNotify ActiveChart.hWnd, strMsg
                            If Not FormIsLoaded("frmAlertMessages") Then frmAlertMessages.ShowMe
                        End If
                    End If
                    
            End Select
        End If
    Next lIndex

    If m.nAlertType = eGDAlertType_QuoteBoard And Not bChangeCellColor Then
        'alert is true, but change cell color was not one of the currently selected action item
        'clear cell color in case it was previously colored by change-cell-color action
        frmQuotes.ColorCell strSymbol, strPeriod, m.strField, m.strTabName
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckAlert
'' Description: If the alert is true, do the given action
'' Inputs:      Only check back color?, Order Information
''              bConfirmOverride - set true to override confirm order at create/edit time (4307)
'' Returns:     True if the alert was hit, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckAlert(Optional ByVal bOnlyBackColor As Boolean = False, _
    Optional ByVal strOrderInfo As String = "", _
    Optional ByVal bChartAlert As Boolean = False, _
    Optional ByVal bConfirmOverride As Boolean = False, _
    Optional ByVal StatusInfo As cStatusAlertInfo = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim dCurrentValue As Double         ' Current value from the database
    Dim TradeItem As New cAutoTradeItem ' Temporary automated trading item
    Dim strCaption As String            ' Caption for popup message
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim anAlert As New cGdArray         ' Array of whether to alert or not
    Dim bAlert As Boolean               ' Return value from the function
    Dim astrSymbols As New cGdArray     ' List of symbols to check
    Dim strSymbol As String             ' Symbol
    Dim strPeriod As String             ' Period
    Dim Order As New cPtOrder           ' Temporary order object
    Dim Bars As New cGdBars             ' Temprary bars object
    Dim bUsedOrderInfo As Boolean       ' Did we use the order info passed in?
    Dim bGridStyle As Boolean           ' Is the quote board tab a grid style tab?
    Dim strOrderStatus As String        ' Order status from the order info
    
    Dim dAlertTime As Double            'time from bars if available, else feed time, else machine time converted to NY
    Dim dLastCheckedTime As Double      'time alert was last check converted to match time zone of bar if necessary
    Dim strAlertTimeZone As String      'time zone of dAlertTime
    Dim strNewQBSymTrue As String       'list of symbols that are true for QB tab alert, delimited by !
    
    Dim iBar As Long                    'last good data bar on chart
    Dim bWhatIfMode As Boolean
    Dim astrOrderInfo As cGdArray       ' Order information array
    Dim alLastBar As cGdArray           ' Last bar in the Bars
    Dim bNewBar As Boolean              ' Is this a new bar?
    Dim dNyTime As Double
    Dim lTradeItem As Long              ' Index into a for loop
    
    Dim bGameMode As Boolean            'true = this alert on a chart that is in Instant Replay mode
    
    Dim ePriceConditionTrue As eGDPriceConditionTrue
    
    If g.bStarting Or g.bUnloading Or m.bEditInProg Then Exit Function
    
    anAlert.Create eGDARRAY_TinyInts
    astrSymbols.Create eGDARRAY_Strings
    strNewQBSymTrue = ""
    dAlertTime = 0#
    
    Set alLastBar = New cGdArray
    alLastBar.Create eGDARRAY_Longs, 1, kNullData
    
    Select Case m.nAlertType
        Case eGDAlertType_QuoteBoard
            anAlert(0) = 0
            
            If m.bSymbol Then
                If m.lSymbolID = 0 Then
                    astrSymbols(0) = m.strSymbol & ";" & m.strPeriod
                Else
                    astrSymbols(0) = Str(m.lSymbolID) & ";" & m.strPeriod
                End If
            Else
                bGridStyle = False
                For lIndex2 = 0 To frmQuotes.TabRecords - 1
                    If frmQuotes.TabStr(eGDTabSettings_Name, lIndex2) = m.strTabName Then
                        astrSymbols.SplitFields frmQuotes.TabStr(eGDTabSettings_Symbols, lIndex2), ","      'symbolId;periodString
                        If Val(frmQuotes.TabStr(eGDTabSettings_Style, lIndex2)) = -1 Then bGridStyle = True
                        Exit For
                    End If
                Next lIndex2
                If Not bGridStyle Then GoTo ErrExit         'tab alerts not supported by box-style QB
            End If
            
            If astrSymbols.Size <> alLastBar.Size Then
                alLastBar.Size = astrSymbols.Size
            End If
            
            For lIndex2 = 0 To astrSymbols.Size - 1
                anAlert(lIndex2) = 0
                
                ' Get the current value from the grid
                strSymbol = ""
                If CLng(Val(Parse(astrSymbols(lIndex2), ";", 1))) = 0 Then
                    dCurrentValue = frmQuotes.SymbolData(Parse(astrSymbols(lIndex2), ";", 1), Parse(astrSymbols(lIndex2), ";", 2), m.strField)
                    If m.bSymbol Then
                        Set Bars = frmQuotes.GetBars(Parse(astrSymbols(lIndex2), ";", 1), Parse(astrSymbols(lIndex2), ";", 2))
                        If m.astrSymbols(lIndex2) <> astrSymbols(lIndex2) Then
                            alLastBar(lIndex2) = kNullData
                            m.astrSymbols(lIndex2) = astrSymbols(lIndex2)
                        Else
                            alLastBar(lIndex2) = Bars.Size - 1
                        End If
                    End If
                Else
                    dCurrentValue = frmQuotes.SymbolData(CLng(Val(Parse(astrSymbols(lIndex2), ";", 1))), Parse(astrSymbols(lIndex2), ";", 2), m.strField)
                    If m.bSymbol Then
                        Set Bars = frmQuotes.GetBars(CLng(Val(Parse(astrSymbols(lIndex2), ";", 1))), Parse(astrSymbols(lIndex2), ";", 2)) '
                        If m.astrSymbols(lIndex2) <> astrSymbols(lIndex2) Then
                            alLastBar(lIndex2) = kNullData
                            m.astrSymbols(lIndex2) = astrSymbols(lIndex2)
                        Else
                            alLastBar(lIndex2) = Bars.Size - 1
                        End If
                    End If
                End If
                If InStr(m.strField, "%") <> 0 Then dCurrentValue = dCurrentValue * 100
                If dCurrentValue <> kNullData Then
                    Select Case UCase(m.strOperator)
                        Case "<", "LESS THAN"
                            If dCurrentValue < m.dValue Then anAlert(lIndex2) = True
                        Case ">", "GREATER THAN"
                            If dCurrentValue > m.dValue Then anAlert(lIndex2) = True
                        Case "=", "EQUAL"
                            If dCurrentValue = m.dValue Then anAlert(lIndex2) = True
                        Case "<=", "LESS THAN OR EQUAL"
                            If dCurrentValue <= m.dValue Then anAlert(lIndex2) = True
                        Case ">=", "GREATER THAN OR EQUAL"
                            If dCurrentValue >= m.dValue Then anAlert(lIndex2) = True
                    End Select
                    
                    If Not m.bSymbol Then
                        'add symbolID & period to list of symbol that alert is true for when alert is set for entire QB Tab
                        If anAlert(lIndex2) = True Then
                            strNewQBSymTrue = strNewQBSymTrue & astrSymbols(lIndex2) & "!"
                        End If
                    End If
                End If
            Next lIndex2
            
        Case eGDAlertType_AutoTrade
            anAlert(0) = 0
            m.bDeactivate = False
            
            Select Case m.nOrderAlert
                Case eGDOrderAlertType_AutoTrade
                    Set TradeItem = g.TradingItems(m.lTradeItemID)
                    If Not TradeItem Is Nothing Then
                        If TradeItem.ParentID >= 0& Then
                            anAlert(0) = CheckAlertForTradeItem(TradeItem)
                        Else
                            anAlert(0) = 0
                            For lTradeItem = 1 To g.TradingItems.Count
                                If g.TradingItems(lTradeItem).ParentID = TradeItem.AutoTradeItemID Then
                                    If CheckAlertForTradeItem(g.TradingItems(lTradeItem)) Then
                                        anAlert(0) = 1
                                    End If
                                End If
                            Next lTradeItem
                        End If
                        
                        strCaption = TradeItem.Name
                    Else
                        m.bActive = False
                    End If
                
                Case eGDOrderAlertType_ManualOrder
                    If Order.Load(m.lTradeItemID) = True Then
                        Select Case UCase(m.strAutoTradeCondition)
                            Case "ORDER SUBMITTED"
                                anAlert(0) = Abs(HasBeenSent(Order.Status))
                            Case "ORDER PRICE HIT"
                                Set Bars = GetTradeBars(Order.SymbolOrSymbolID)
                                anAlert(0) = Abs(Order.OrderPriceHit(Bars(eBARS_Close, Bars.Size - 1)))
                            Case "ORDER FILLED"
                                anAlert(0) = Abs((Order.Status = eTT_OrderStatus_Filled) And (InStr(strOrderInfo, "Fill    : ") > 0))
                            Case "ORDER CANCELLED"
                                anAlert(0) = Abs((Order.Status = eTT_OrderStatus_Cancelled) Or (Order.Status = eTT_OrderStatus_BalCancelled))
                            Case "ORDER REJECTED"
                                anAlert(0) = Abs(Order.Status = eTT_OrderStatus_Rejected)
                        End Select
                        If Len(Order.BrokerID) > 0 Then
                            strCaption = Order.OrderText & " - " & Order.BrokerID
                        Else
                            strCaption = Order.OrderText
                        End If
                    Else
                        m.bActive = False
                    End If
                
                Case eGDOrderAlertType_AllOrders
                    strOrderStatus = UCase(Trim(Parse(Parse(strOrderInfo, ";", 3), ":", 2)))
                    Select Case UCase(m.strAutoTradeCondition)
                        Case "ORDER SUBMITTED"
                            anAlert(0) = (strOrderStatus = "SENT")
                            strCaption = "Order Submitted"
                        Case "ORDER FILLED"
                            anAlert(0) = ((strOrderStatus = "FILLED") And (InStr(strOrderInfo, "Fill    : ") > 0))
                            strCaption = "Order Filled"
                        Case "ORDER CANCELLED"
                            anAlert(0) = (strOrderStatus = "CANCELLED") Or (strOrderStatus = "EXPIRED")
                            strCaption = "Order Cancelled"
                        Case "ORDER REJECTED"
                            anAlert(0) = (strOrderStatus = "REJECTED")
                            strCaption = "Order Rejected"
                    End Select
                                    
            End Select
            
        Case eGDAlertType_Status
            anAlert(0) = 0
            m.bDeactivate = False
            
            Select Case UCase(m.strStatusItem)
                Case "ORDER STATUS CHANGE", UCase(kstrPositionMismatch), UCase(kstrAutoTradeItemDisabled)
                    anAlert(0) = Abs(True)
                    m.bLastCheckedStatus = False
                
                Case "GENESIS STREAMING", "E-SIGNAL STREAMING"
                    ' ignore this check on Sun mornings (NY time) since we do a normal disconnect then
                    dNyTime = ConvertTimeZone(Now, "", "NY")
                    If (mGenesis.Weekday(dNyTime) = vbSunday And Hour(dNyTime) < 12) = False Then
                        ' NOTE: 0 = connected, 1 = disconnected
                        anAlert(0) = Abs(Not g.RealTime.IsServerActive(True))
                        ' TLB 10/22/2012: check if user disconnected manually
                        ' (i.e. if not connected now but we were connected the last we checked)
                        If anAlert(0) = 1 And m.bLastCheckedStatus = False Then
                            If UCase(g.RealTime.DisconnectReason) = "USER TURNED OFF" Then
                                ' to skip the alert: pretend we're disconnected now
                                ' and were disconnected the last time we checked
                                anAlert(0) = 1
                                m.bLastCheckedStatus = True
                            End If
                        End If
                    End If
                
                Case Else
                    If m.nBroker > 0 Then
                        ' TLB 7/1/2012: ignore broker disconnects on weekends (NY time)
                        dNyTime = ConvertTimeZone(Now, "", "NY")
                        Select Case mGenesis.Weekday(dNyTime)
                        Case vbFriday
                            If Hour(dNyTime) > 17 Then
                                dNyTime = 0
                            End If
                        Case vbSaturday
                            dNyTime = 0
                        Case vbSunday
                            If Hour(dNyTime) < 12 Then
                                dNyTime = 0
                            End If
                        End Select
                        If dNyTime > 0 Then
                            ' TRIGGER for a Broker disconnect
                            'anAlert(0) = Abs(g.Broker.ConnectionStatusForBroker(m.nBroker) <> eGDConnectionStatus_Connected)
                            anAlert(0) = Abs(True)
                            m.bLastCheckedStatus = False
                        End If
                    End If
                    
            End Select
            strCaption = m.strStatusItem
            
        Case eGDAlertType_Time
            anAlert(0) = 0
            g.RealTime.FeedTime
            strCaption = DateFormat(m.dNextTrigger, MM_DD_YYYY, HH_MM, AMPM_LOWER)
            dAlertTime = ConvertTimeZone(Now, "", m.strTimeZone)
            strAlertTimeZone = m.strTimeZone
            If dAlertTime >= m.dNextTrigger Then
                anAlert(0) = 1
                CalcNextTriggerTime True
            End If
            
        Case eGDAlertType_Price
            anAlert(0) = 0
            
            Set Bars = g.Alerts.GetBars(SymbolOrSymbolID, m.strPeriod, NumBars)
            
            If (m.strSymbol & ";" & m.strPeriod) <> m.astrSymbols(0) Then
                alLastBar(0) = kNullData
                m.astrSymbols(0) = m.strSymbol & ";" & m.strPeriod
            Else
                alLastBar(0) = Bars.Size - 1
            End If
                
            If m.bUseGetsDownTo Then
                If m.eAlertBarFlag = eGDAlertBar1_DownToClose Or m.eAlertBarFlag = eGDAlertBar3_CloseThisBar Then
                    If Bars(eBARS_Close, Bars.Size - 1) <= m.dGetsDownTo Then
                        anAlert(0) = 1
                        ePriceConditionTrue = eGDConditionTrue_DownTo
                        If m.bUseGetsUpTo And m.ePriceConditionTrue <> eGDConditionTrue_DownTo Then
                            'both UpTo AND DownTo prices are selected; reset last checked flag so DownTo price can trigger
                            If m.dGetsDownTo <> m.dGetsUpTo Then m.bLastCheckedStatus = False
                        End If
                        m.ePriceConditionTrue = ePriceConditionTrue
                    End If
                ElseIf Bars(eBARS_Low, Bars.Size - 1) <= m.dGetsDownTo Then
                    anAlert(0) = 1
                    ePriceConditionTrue = eGDConditionTrue_DownTo
                    If m.bUseGetsUpTo And m.ePriceConditionTrue <> eGDConditionTrue_DownTo Then
                        'both UpTo AND DownTo prices are selected; reset last checked flag so DownTo price can trigger
                        If m.dGetsDownTo <> m.dGetsUpTo Then m.bLastCheckedStatus = False
                    End If
                    m.ePriceConditionTrue = ePriceConditionTrue
                End If
            End If
            If m.bUseGetsUpTo Then
                If m.eAlertBarFlag = eGDAlertBar2_UpToClose Or m.eAlertBarFlag = eGDAlertBar3_CloseThisBar Then
                    If Bars(eBARS_Close, Bars.Size - 1) >= m.dGetsUpTo Then
                        anAlert(0) = 1
                        ePriceConditionTrue = eGDConditionTrue_UpTo
                        If m.bUseGetsDownTo And m.ePriceConditionTrue <> eGDConditionTrue_UpTo Then
                            'both UpTo AND DownTo prices are selected; reset last checked flag so UpTo price can trigger
                            If m.dGetsDownTo <> m.dGetsUpTo Then m.bLastCheckedStatus = False
                        End If
                        m.ePriceConditionTrue = ePriceConditionTrue
                    End If
                ElseIf Bars(eBARS_High, Bars.Size - 1) >= m.dGetsUpTo Then
                    anAlert(0) = 1
                    ePriceConditionTrue = eGDConditionTrue_UpTo
                    If m.bUseGetsDownTo And m.ePriceConditionTrue <> eGDConditionTrue_UpTo Then
                        'both UpTo AND DownTo prices are selected; reset last checked flag so UpTo price can trigger
                        If m.dGetsDownTo <> m.dGetsUpTo Then m.bLastCheckedStatus = False
                    End If
                    m.ePriceConditionTrue = ePriceConditionTrue
                End If
            End If
        
        Case eGDAlertType_TradeSense
            anAlert(0) = 0
            
            Set Bars = g.Alerts.GetBars(SymbolOrSymbolID, m.strPeriod, NumBars)
            
            If (m.strSymbol & ";" & m.strPeriod) <> m.astrSymbols(0) Then
                alLastBar(0) = kNullData
                m.astrSymbols(0) = m.strSymbol & ";" & m.strPeriod
            Else
                alLastBar(0) = Bars.Size - 1
            End If
                
            If RunExpression = True Then
                anAlert(0) = 1
            End If
            
        Case eGDAlertType_Chart
            anAlert(0) = Abs(bChartAlert)
            If Not m.Ind Is Nothing Then
                Set Bars = m.Ind.Bars
                If Not Bars Is Nothing And Not m.Ind.IndChart Is Nothing Then
                    Set Bars = m.Ind.IndChart.Bars
                    iBar = m.Ind.IndChart.LastGoodDataBar(False, False)
                    bWhatIfMode = m.Ind.IndChart.IsInWhatIfMode
                
                    If (m.strSymbol & ";" & m.strPeriod) <> m.astrSymbols(0) Then
                        alLastBar(0) = kNullData
                        m.astrSymbols(0) = m.strSymbol & ";" & m.strPeriod
                    Else
                        alLastBar(0) = iBar
                    End If
                End If
            End If
        
        Case eGDAlertType_Annot
            anAlert(0) = Abs(bChartAlert)
            If Not m.Annot Is Nothing Then
                If Not m.Annot.AnnotChart Is Nothing Then
                    Set Bars = m.Annot.AnnotChart.Bars
                    iBar = m.Annot.AnnotChart.LastGoodDataBar(False, False)
                    bWhatIfMode = m.Annot.AnnotChart.IsInWhatIfMode
                
                    If (m.strSymbol & ";" & m.strPeriod) <> m.astrSymbols(0) Then
                        alLastBar(0) = kNullData
                        m.astrSymbols(0) = m.strSymbol & ";" & m.strPeriod
                    Else
                        alLastBar(0) = iBar
                    End If
                End If
            End If
        
    End Select
    
    'store alert time as Bars time if availaible, else use feed time which is NY time
    If Not Bars Is Nothing Then
        dAlertTime = Bars.Prop(eBARS_LastTickTime)
        If dAlertTime <> 0 Then
            If iBar <= 0 Then iBar = Bars.Size - 1
            dAlertTime = Int(Bars(eBARS_DateTime, iBar)) + dAlertTime / 1440#
            strAlertTimeZone = Bars.Prop(eBARS_ExchangeTimeZoneInf)
        End If
        If Bars.Size > 0 Then
            Set m.Bars = Bars
            m.strSymbol = Bars.Prop(eBARS_Symbol)
            m.lSymbolID = Bars.Prop(eBARS_SymbolID)
            m.strPeriod = Bars.Prop(eBARS_PeriodicityStr)
            strAlertTimeZone = Bars.Prop(eBARS_ExchangeTimeZoneInf)
        End If
    End If
    
    If dAlertTime <= 0 Then
        If g.RealTime.Active And g.RealTime.FeedTime() > 0 Then
            dAlertTime = g.RealTime.FeedTime()
        Else
            dAlertTime = ConvertTimeZone(CDbl(Now))
        End If
        strAlertTimeZone = "NY"
    End If
    
    'check for case when user had real time on then turned off and Bars time reverted to last QB refresh time
    If Not g.RealTime.Active Then
        If Len(m.strLastCheckedTimeZone) = 0 Or m.dLastCheckedTime = 0 Or m.strLastCheckedTimeZone = strAlertTimeZone Then
            dLastCheckedTime = m.dLastCheckedTime
        Else
            dLastCheckedTime = ConvertTimeZone(m.dLastCheckedTime, m.strLastCheckedTimeZone, strAlertTimeZone)
        End If
        If RoundNum(dLastCheckedTime, 10) >= RoundNum(dAlertTime, 10) Then
            'don't want to do this because confirm override = true must mean user just edited alert
            CheckAlert = m.bLastCheckedStatus
            
            'issue 5861
            If m.nAlertType = eGDAlertType_QuoteBoard Then
                If m.bLastCheckedStatus Then
                    frmQuotes.ColorCell m.strSymbol, m.strPeriod, m.strField, m.strTabName, ValOfText(Parse(m.astrActions(eAA_ChangeBackColor), ",", 2))
                Else
                    frmQuotes.ColorCell m.strSymbol, m.strPeriod, m.strField, m.strTabName      'clear cell color
                End If
            End If
            
            If m.nAlertType = eGDAlertType_Annot Or m.nAlertType = eGDAlertType_Chart Then
                If Not m.Annot Is Nothing Then
                    If m.Annot.AnnotChart Is Nothing Then
                        Exit Function
                    ElseIf m.Annot.AnnotChart.Form Is Nothing Then
                        Exit Function
                    ElseIf m.Annot.AnnotChart.Form.GameMode Is Nothing Then
                        Exit Function
                    Else
                        bGameMode = True        '6597
                    End If
                ElseIf Not m.Ind Is Nothing Then
                    If m.Ind.IndChart Is Nothing Then
                        Exit Function
                    ElseIf m.Ind.IndChart.Form Is Nothing Then
                        Exit Function
                    ElseIf m.Ind.IndChart.Form.GameMode Is Nothing Then
                        Exit Function
                    Else
                        bGameMode = True
                    End If
                Else
                    Exit Function
                End If
            Else
                Exit Function
            End If
        End If
    End If
        
    If m.dLastCheckedTime = kNullData Then
        'at startup, last checked time for streaming is set to knulldata to indicate streaming has not
        'only trigger "streaming was disconnected" alert after streaming has been started at least once
        If g.RealTime.IsServerActive(True) Then
            m.dLastCheckedTime = 0
        Else
            CheckAlert = False
            anAlert(0) = 0
        End If
    Else
        m.dLastCheckedTime = dAlertTime
    End If
    
    m.strLastCheckedTimeZone = strAlertTimeZone
    
    bAlert = False
    
    If m.alLastBarAlerted.Size <> anAlert.Size Then
        m.alLastBarAlerted.Size = anAlert.Size
    End If
       
    For lIndex2 = 0 To anAlert.Size - 1
        ' TLB 2/21/2012 -- added this check in case backing up during streaming replay:
        ' if last bar < last bar alerted, then must have backed up so just clear the last bar alerted
        If alLastBar(lIndex2) < m.alLastBarAlerted(lIndex2) Then
            m.alLastBarAlerted(lIndex2) = kNullData
        End If
        
        Select Case m.nAlertType
            Case eGDAlertType_QuoteBoard
                strSymbol = Parse(astrSymbols(lIndex2), ";", 1)
                If Val(strSymbol) > 0 Then strSymbol = GetSymbol(Val(strSymbol))
                strPeriod = Parse(astrSymbols(lIndex2), ";", 2)
                
            Case eGDAlertType_Annot, eGDAlertType_Chart, eGDAlertType_Price, eGDAlertType_TradeSense
                strSymbol = m.strSymbol
                strPeriod = m.strPeriod
                
            Case eGDAlertType_AutoTrade
                If Len(strOrderInfo) > 0 Then
                    strSymbol = Parse(strOrderInfo, ";", 1)
                    strSymbol = Trim(Parse(strSymbol, ":", 2))
                    
                    If UCase(m.strAutoTradeCondition) = "ORDER FILLED" Then
                        Set astrOrderInfo = New cGdArray
                        astrOrderInfo.SplitFields strOrderInfo, ";"
                        astrOrderInfo.Remove 2
                        astrOrderInfo.Remove 0
                        strOrderInfo = astrOrderInfo.JoinFields(";")
                    End If
                End If
            
            Case eGDAlertType_Status
                If Len(strOrderInfo) > 0 Then
                    strSymbol = Parse(strOrderInfo, ";", 1)
                    strSymbol = Trim(Parse(strSymbol, ":", 2))
                End If
        
        End Select
        
        If (anAlert(lIndex2) <> 0) Then
            bAlert = True
            
            If alLastBar(lIndex2) = kNullData Then
                bNewBar = True
            Else
                bNewBar = alLastBar(lIndex2) <> m.alLastBarAlerted(lIndex2)
                If bNewBar Then
                    m.alLastBarAlerted(lIndex2) = alLastBar(lIndex2)
                End If
            End If
            
            If bWhatIfMode Then
                HandleAlertAction bOnlyBackColor, dAlertTime, strSymbol, strPeriod, strOrderInfo, strCaption, True, ePriceConditionTrue, bNewBar, StatusInfo
            Else
                HandleAlertAction bOnlyBackColor, dAlertTime, strSymbol, strPeriod, strOrderInfo, strCaption, bConfirmOverride, ePriceConditionTrue, bNewBar, StatusInfo
            End If
        ElseIf m.nAlertType = eGDAlertType_QuoteBoard Then
            frmQuotes.ColorCell strSymbol, strPeriod, m.strField, m.strTabName        'alert is false, clear cell color
        End If
    Next lIndex2

    If ((m.bDeactivate = True) And (bAlert = True) And (bAlert <> m.bLastCheckedStatus)) Then
        m.bActive = False
        'need to update the alert management form if it is visible
        If bChartAlert Then
            If FormIsLoaded("frmAlertsSetup") Then
                If frmAlertsSetup.Visible Then frmAlertsSetup.LoadGrid
            End If
            If Not m.Annot Is Nothing Then
                m.Annot.geDrawAnn Nothing
            End If
        End If
    End If
    
    'reset & replace list of symbol that alert was true for with new list (this is for QB Tab alert)
    m.astrQBSymTrue.Size = 0
    If Len(strNewQBSymTrue) > 0 Then
        m.astrQBSymTrue.SplitFields strNewQBSymTrue, "!"
    End If
    
    m.bLastCheckedStatus = bAlert
    CheckAlert = bAlert
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.CheckAlert", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnglishString
'' Description: Make an english string to put into the alerts grid
'' Inputs:      None
'' Returns:     String to store into the grid
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnglishString(Optional ByVal bIncludeAction As Boolean = True, Optional ByVal strSymbol As String = "", _
    Optional ByVal strPeriod As String = "", Optional ByVal ePriceCondition As eGDPriceConditionTrue = eGDConditionTrue_None, _
    Optional ByVal StatusInfo As cStatusAlertInfo = Nothing) As String
On Error GoTo ErrSection:

    Dim astrActions As New cGdArray     ' String version of the action
    Dim Bars As New cGdBars             ' Temporary Bars object
    Dim TradeItem As New cAutoTradeItem ' Temporary automated trading item
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As New cPtOrder           ' Temporary order object
    Dim astrDaysOfWeek As New cGdArray  ' Array of days of week
    Dim strName As String               ' Account name
    
    If bIncludeAction Then
        EnglishString = "If "
    End If
    
    Select Case m.nAlertType
        Case eGDAlertType_QuoteBoard
            If Len(strSymbol) = 0 Then strSymbol = m.strSymbol
            If Len(strPeriod) = 0 Then strPeriod = m.strPeriod
            
            If (m.bSymbol = True) Or ((Len(strSymbol) > 0) And (Len(strPeriod) > 0)) Then
                EnglishString = EnglishString & m.strField & " of " & strSymbol & "(" & strPeriod
                EnglishString = EnglishString & ") " & m.strOperator
            Else
                EnglishString = EnglishString & m.strField & " of " & m.strTabName & " " & m.strOperator
            End If
            Select Case m.strField
                Case "Session"
                    EnglishString = EnglishString & " " & DateFormat(m.dValue)
                Case "Last Tick"
                    EnglishString = EnglishString & " " & DateFormat(m.dValue, , HH_MM, AMPM_LOWER)
                Case Else
                    If m.bAsTradingUnits Then
                        SetBarProperties Bars, strSymbol
                        EnglishString = EnglishString & " " & Bars.PriceDisplay(m.dValue, True)
                    ElseIf IsBooleanCriteria(m.strCriteriaID) Then
                        If m.dValue = 0 Then
                            EnglishString = EnglishString & " False"
                        Else
                            EnglishString = EnglishString & " True"
                        End If
                    Else
                        EnglishString = EnglishString & " " & CStr(m.dValue)
                    End If
            End Select
        
        Case eGDAlertType_AutoTrade
            Select Case m.nOrderAlert
                Case eGDOrderAlertType_AutoTrade
                    Set TradeItem = g.TradingItems(Str(m.lTradeItemID))
                    EnglishString = EnglishString & m.strAutoTradeCondition & " on "
                    If Not TradeItem Is Nothing Then
                        EnglishString = EnglishString & TradeItem.Name
                    Else
                        EnglishString = EnglishString & "<TRADING ITEM NO LONGER EXISTS>"
                    End If
                
                Case eGDOrderAlertType_ManualOrder
                    If Order.Load(m.lTradeItemID) = True Then
                        EnglishString = EnglishString & Order.OrderText
                    Else
                        EnglishString = EnglishString & "<ORDER NO LONGER EXISTS>"
                    End If
                    Select Case UCase(m.strAutoTradeCondition)
                        Case "ORDER SUBMITTED"
                            EnglishString = EnglishString & " is Submitted"
                        Case "ORDER PRICE HIT"
                            EnglishString = EnglishString & " has the Order Price Hit"
                        Case "ORDER FILLED"
                            EnglishString = EnglishString & " is Filled"
                        Case "ORDER CANCELLED"
                            EnglishString = EnglishString & " is Cancelled"
                        Case "ORDER REJECTED"
                            EnglishString = EnglishString & " is Rejected"
                    End Select
                
                Case eGDOrderAlertType_AllOrders
                    Select Case m.lTradeItemID
                        Case 0:
                            EnglishString = "Any Order"
                        
                        Case -1:
                            EnglishString = "Any Order in a Live Account"
                        
                        Case -2:
                            EnglishString = "Any Order in a Simulated Account"
                        
                        Case Else:
                            strName = g.Broker.AccountNameForID(m.lTradeItemID)
                            If Len(strName) > 0 Then
                                EnglishString = "Any Order in Account " & g.Broker.AccountNameForID(m.lTradeItemID)
                            Else
                                EnglishString = "Any Order in Account <ACCOUNT NO LONGER EXISTS>"
                            End If
                    
                    End Select
                    Select Case UCase(m.strAutoTradeCondition)
                        Case "ORDER SUBMITTED"
                            EnglishString = EnglishString & " is Submitted"
                        Case "ORDER PRICE HIT"
                            EnglishString = EnglishString & " has the Order Price Hit"
                        Case "ORDER FILLED"
                            EnglishString = EnglishString & " is Filled"
                        Case "ORDER CANCELLED"
                            EnglishString = EnglishString & " is Cancelled"
                        Case "ORDER REJECTED"
                            EnglishString = EnglishString & " is Rejected"
                    End Select
                
            End Select
            
        Case eGDAlertType_Status
            Select Case UCase(m.strStatusItem)
                Case "ORDER STATUS CHANGE"
                    EnglishString = EnglishString & "status has changed on any order"
                Case UCase(kstrPositionMismatch)
                    If StatusInfo Is Nothing Then
                        EnglishString = EnglishString & "position mismatch"
                    Else
                        EnglishString = EnglishString & "position mismatch on '" & StatusInfo.Symbol & "' in account '" & StatusInfo.Account & "'"
                    End If
                Case UCase(kstrAutoTradeItemDisabled)
                    If StatusInfo Is Nothing Then
                        EnglishString = EnglishString & "automated trading item disabled abnormally"
                    Else
                        EnglishString = EnglishString & "the automated trading item '" & StatusInfo.AutoTradeItem & "' running on symbol '" & StatusInfo.Symbol & "' in account '" & StatusInfo.Account & "' has been disabled abnormally because '" & StatusInfo.Reason & "'"
                    End If
                Case Else
                    EnglishString = EnglishString & m.strStatusItem & " was disconnected"
            End Select
            
        Case eGDAlertType_Time
            If m.bWeekday = True Then
                If Len(m.strTimeZone) > 0 Then
                    EnglishString = EnglishString & m.strTimeZone & " day of week is "
                Else
                    EnglishString = EnglishString & "Day of week is "
                End If
                
                If Mid(m.strDayMask, vbSunday, 1) = "1" Then astrDaysOfWeek.Add "Sunday"
                If Mid(m.strDayMask, vbMonday, 1) = "1" Then astrDaysOfWeek.Add "Monday"
                If Mid(m.strDayMask, vbTuesday, 1) = "1" Then astrDaysOfWeek.Add "Tuesday"
                If Mid(m.strDayMask, vbWednesday, 1) = "1" Then astrDaysOfWeek.Add "Wednesday"
                If Mid(m.strDayMask, vbThursday, 1) = "1" Then astrDaysOfWeek.Add "Thursday"
                If Mid(m.strDayMask, vbFriday, 1) = "1" Then astrDaysOfWeek.Add "Friday"
                If Mid(m.strDayMask, vbSaturday, 1) = "1" Then astrDaysOfWeek.Add "Saturday"
                
                If astrDaysOfWeek.Size = 1 Then
                    EnglishString = EnglishString & astrDaysOfWeek(0)
                ElseIf astrDaysOfWeek.Size = 2 Then
                    EnglishString = EnglishString & astrDaysOfWeek(0) & " or " & astrDaysOfWeek(1)
                Else
                    For lIndex = 0 To astrDaysOfWeek.Size - 2
                        EnglishString = EnglishString & astrDaysOfWeek(lIndex) & ", "
                    Next lIndex
                    EnglishString = EnglishString & " or " & astrDaysOfWeek(astrDaysOfWeek.Size - 1)
                End If
                
                If Len(m.strTimeZone) > 0 Then
                    EnglishString = EnglishString & " and " & m.strTimeZone & " time is "
                Else
                    EnglishString = EnglishString & " and time is "
                End If
                EnglishString = EnglishString & DateFormat(m.dTriggerDateTime, NO_DATE, HH_MM, AMPM_LOWER)
            Else
                If Len(m.strTimeZone) > 0 Then
                    EnglishString = EnglishString & m.strTimeZone & " time is "
                Else
                    EnglishString = EnglishString & "time is "
                End If
                EnglishString = EnglishString & DateFormat(m.dTriggerDateTime, MM_DD_YYYY, HH_MM, AMPM_LOWER)
            End If
            
        Case eGDAlertType_Price
            SetBarProperties Bars, m.strSymbol
            EnglishString = EnglishString & "price of " & m.strSymbol
            If m.bUseGetsDownTo Then
                If ePriceCondition = eGDConditionTrue_None Or ePriceCondition = eGDConditionTrue_DownTo Then
                    EnglishString = EnglishString & " is down to " & Bars.PriceDisplay(m.dGetsDownTo)
                End If
                If m.bUseGetsUpTo Then
                    If ePriceCondition = eGDConditionTrue_None Then
                        EnglishString = EnglishString & " or is up to " & Bars.PriceDisplay(m.dGetsUpTo)
                    ElseIf ePriceCondition = eGDConditionTrue_UpTo Then
                        EnglishString = EnglishString & " is up to " & Bars.PriceDisplay(m.dGetsUpTo)
                    End If
                End If
            ElseIf m.bUseGetsUpTo Then
                EnglishString = EnglishString & " is up to " & Bars.PriceDisplay(m.dGetsUpTo)
            End If
        
        Case eGDAlertType_TradeSense
            EnglishString = EnglishString & "(" & m.strPriceCondition & ") Of " & Chr(34) & _
                    m.strSymbol & ";" & m.strPeriod & Chr(34)
            
        Case eGDAlertType_Annot
            EnglishString = m.strSymbol & ":" & m.strChartCondition
    
        Case eGDAlertType_Chart
            EnglishString = m.strChartCondition
    
    End Select
    
    If bIncludeAction Then
        EnglishString = EnglishString & ", Then "
        
        For lIndex = 0 To m.astrActions.Size - 1
            If Parse(m.astrActions(lIndex), ",", 1) <> "0" Then
                Select Case lIndex
                    Case eAA_MessageBox
                        astrActions.Add "Pop Up a Message"
                    Case eAA_LogToFile
                        astrActions.Add "Log Message to a File"
                    Case eAA_ChangeBackColor
                        astrActions.Add "Change Cell Color"
                    Case eAA_SendPage
                        astrActions.Add "Send Page"
                    Case eAA_SendEmail
                        astrActions.Add "Send Email"
                    Case eAA_PlaceOrder
                        astrActions.Add "Place Order"
                    Case eAA_PlaySound
                        astrActions.Add "Play Sound"
                    Case eAA_MsgHistory
                        astrActions.Add "Add to Message History"
                End Select
            End If
        Next lIndex
        
        If astrActions.Size = 1 Then
            EnglishString = EnglishString & astrActions(0)
        ElseIf astrActions.Size = 2 Then
            EnglishString = EnglishString & astrActions(0) & " and " & astrActions(1)
        ElseIf astrActions.Size > 0 Then
            For lIndex = 0 To astrActions.Size - 2
                EnglishString = EnglishString & astrActions(lIndex) & ", "
            Next lIndex
            EnglishString = EnglishString & "and " & astrActions(astrActions.Size - 1)
        End If
    Else
        EnglishString = UCase(Left(EnglishString, 1)) & Mid(EnglishString, 2)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.EnglishString", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DataKindName
'' Description: Get the field name that the Bars structure is expecting for a
''              field name that is chosen
'' Inputs:      Field that has been chosen
'' Returns:     Field name for the Bars structure
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DataKindName(ByVal pstrField As String) As String
On Error GoTo ErrSection:

    Select Case UCase(pstrField)
        Case "OPEN"
            DataKindName = "EODOpen"
        Case "HIGH"
            DataKindName = "EODHigh"
        Case "LOW"
            DataKindName = "EODLow"
        Case "LAST"
            DataKindName = "EODClose"
        Case "VOLUME"
            DataKindName = "EODCVol"
        Case "OPEN INTEREST"
            DataKindName = "EODCOI"
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.DataKindName", eGDRaiseError_Raise
    
End Function

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderInEnglish
'' Description: Make an english string out of the given order string
'' Inputs:      None
'' Returns:     English String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderInEnglish() As String
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Order broken out into an array
    Dim strReturn As String             ' String to return from the function
    
    astrOrder.Create eGDARRAY_Strings
    astrOrder.SplitFields m.strOrderText, ","
    
    strReturn = astrOrder(0) & " " & Format(astrOrder(1), "#,##0") & " "
    strReturn = strReturn & m.strSymbol & " "
    Select Case ValOfText(astrOrder(2))
        Case eTT_OrderType_Market
            strReturn = strReturn & "At MARKET" & " "
        Case eTT_OrderType_Stop
            strReturn = strReturn & "At " & PriceDisplay(ValOfText(astrOrder(3)), m.strSymbol) & " STOP" & " "
        Case eTT_OrderType_Limit
            strReturn = strReturn & "At " & PriceDisplay(ValOfText(astrOrder(4)), m.strSymbol) & " LIMIT" & " "
        Case eTT_OrderType_StopWithLimit
            strReturn = strReturn & "At " & PriceDisplay(ValOfText(astrOrder(3)), m.strSymbol) & " STOP with " & PriceDisplay(ValOfText(astrOrder(4)), m.strSymbol) & " LIMIT" & " "
    End Select
    strReturn = strReturn & "In Account " & g.Broker.AccountNumberForID(astrOrder(5)) & " "
    If ValOfText(astrOrder(6)) = 0 Then
        strReturn = strReturn & "GTC"
    Else
        strReturn = strReturn & "Expires: " & Format(ValOfText(astrOrder(6)), DateFormat("Format", MM_DD_YYYY))
    End If
    
    OrderInEnglish = strReturn

ErrExit:
    Set astrOrder = Nothing
    Exit Function
    
ErrSection:
    Set astrOrder = Nothing
    RaiseError "cAlert.OrderInEnglish", eGDRaiseError_Raise
    
End Function
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    m.bActive = True
    m.eAlertBarFlag = eGDAlertBar_HiLowThisBar
    
    m.bNeverDeactivate = FileExist("NeverDisable.flg")

    Set m.astrActions = New cGdArray
    m.astrActions.Create eGDARRAY_Strings, eAA_NumActions
    
    Set m.astrQBSymTrue = New cGdArray
    m.astrQBSymTrue.Create eGDARRAY_Strings
    
    Set m.aQbCells = New cGdArray
    
    For lIndex = 0 To m.astrActions.Size - 1
        m.astrActions(lIndex) = "0"
    Next lIndex
    m.astrActions(eAA_MsgHistory) = "1,"
    m.astrActions(eAA_MessageBox) = "1,"
    
    m.bAutoDetect = True
    m.lNumBarsCalc = -1&
    m.lNumBarsOver = -1&
    
    m.nBroker = -1&
    
    Set m.alLastBarAlerted = New cGdArray
    m.alLastBarAlerted.Create eGDARRAY_Longs, 1, kNullData
    
    Set m.astrSymbols = New cGdArray
    m.astrSymbols.Create eGDARRAY_Strings, 1
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlert.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Dim i&
    Dim QbCell As cQuoteCell

    Set m.astrActions = Nothing
    Set m.astrQBSymTrue = Nothing
    Set m.Annot = Nothing
    Set m.Ind = Nothing
    Set m.Bars = Nothing
    
    If Not m.aCharts Is Nothing Then
        For i = 0 To m.aCharts.Size - 1
            Set m.aCharts(i) = Nothing
        Next
        Set m.aCharts = Nothing
    End If
    
    For i = 0 To m.aQbCells.Size - 1
        Set QbCell = m.aQbCells(i)
        If Not QbCell Is Nothing Then
            QbCell.QbCellAlertUpdate Me, 0
        End If
        Set m.aQbCells(i) = Nothing
    Next
    Set m.aQbCells = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAlert.Class_Terminate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AlertTypeText
'' Description: Returns text string for an alert type
'' Inputs:      None
'' Returns:     Alert Type Description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AlertTypeText() As String
On Error GoTo ErrSection:
 
    Dim strType As String               ' Alert type string to return
               
    Select Case m.nAlertType
        Case eGDAlertType_QuoteBoard
            strType = "QB"
        Case eGDAlertType_AutoTrade
            strType = "Order"
        Case eGDAlertType_Chart, eGDAlertType_Annot
            strType = "Chart"
        Case eGDAlertType_Price
            strType = "Price"
        Case eGDAlertType_Time
            strType = "Time"
        Case eGDAlertType_Status
            strType = "Status"
        Case eGDAlertType_TradeSense
            strType = "TradeSense"
        Case Else
            strType = "Unknown"
    End Select
    
    AlertTypeText = strType
    
ErrExit:
    Exit Function
 
ErrSection:
     RaiseError "cAlert.AlertTypeText"
 
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasColorAction
'' Description: Does this alert have a color action active?
'' Inputs:      None
'' Returns:     True if has active color action, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasColorAction() As Boolean
On Error GoTo ErrSection:

    HasColorAction = CBool(Val(Parse(m.astrActions(eAA_ChangeBackColor), ",", 1)))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.HasColorAction"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOrderChangeStatusAlert
'' Description: Is this an order change status alert?
'' Inputs:      None
'' Returns:     True if Order Change Status Alert, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsOrderChangeStatusAlert() As Boolean
On Error GoTo ErrSection:

    IsOrderChangeStatusAlert = ((m.nAlertType = eGDAlertType_Status) And (UCase(m.strStatusItem) = "ORDER STATUS CHANGE"))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.IsOrderChangeStatusAlert"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsPositionMismatchStatusAlert
'' Description: Is this a position mismatch status alert?
'' Inputs:      None
'' Returns:     True if Position Mismatch Status Alert, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsPositionMismatchStatusAlert() As Boolean
On Error GoTo ErrSection:

    IsPositionMismatchStatusAlert = ((m.nAlertType = eGDAlertType_Status) And (UCase(m.strStatusItem) = UCase(kstrPositionMismatch)))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.IsPositionMismatchStatusAlert"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsAutoTradeItemDisabledStatusAlert
'' Description: Is this a automated trading item disabled status alert?
'' Inputs:      None
'' Returns:     True if Automated Trading Item Disabled Status Alert, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsAutoTradeItemDisabledStatusAlert() As Boolean
On Error GoTo ErrSection:

    IsAutoTradeItemDisabledStatusAlert = ((m.nAlertType = eGDAlertType_Status) And (UCase(m.strStatusItem) = UCase(kstrAutoTradeItemDisabled)))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.IsAutoTradeItemDisabledStatusAlert"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsBrokerDisconnectAlert
'' Description: Is this a broker disconnect alert?
'' Inputs:      None
'' Returns:     True if Broker Disconnect Alert, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsBrokerDisconnectAlert() As Boolean
On Error GoTo ErrSection:

    IsBrokerDisconnectAlert = ((m.nAlertType = eGDAlertType_Status) And (m.nBroker > 0))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.IsBrokerDisconnectAlert"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcNextTriggerTime
'' Description: Calculate the next trigger time for a time alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CalcNextTriggerTime(Optional bForceBump As Boolean = False)
On Error GoTo ErrSection:

    Dim dNow As Double                  ' Date and time right now
    Dim dDateTime As Double             ' Date and time of check

    If m.bWeekday = True Then
        dNow = ConvertTimeZone(Now, "", m.strTimeZone)
        dDateTime = Int(dNow) + (m.dTriggerDateTime - Int(m.dTriggerDateTime))
        If dNow > dDateTime Or bForceBump Then dDateTime = dDateTime + 1
        Do While Mid(m.strDayMask, VBA.Weekday(dDateTime), 1) = "0"
            dDateTime = dDateTime + 1
        Loop
        m.dNextTrigger = dDateTime
    Else
        m.dNextTrigger = m.dTriggerDateTime
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlert.CalcNextTriggerTime"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CodedText
'' Description: Determine the coded text for the condition
'' Inputs:      None
'' Returns:     Coded Text string for the condition
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CodedText() As String
On Error GoTo ErrSection:

    Dim strCondition As String          ' Condition to pass to the validator
    Dim Expr As New cExpression         ' Expression to get coded text from
    Dim astrExpressions As New cGdArray ' Array of expressions
    Dim astrBarNames As New cGdArray    ' Array of bar names
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID for the secondary market
    Dim strSymbol As String             ' Symbol for the secondary market
    Dim strPeriod As String             ' Period for the secondary market
    Dim strCodedText As String          ' Coded text string to return from function
    Dim lNumBars As Long                ' Number of bars required
    
    If Len(m.strCodedText) = 0 Then
        With Expr
            .PortfolioNavigator = False
            .Functions = g.Functions
            strCondition = m.strPriceCondition
            .ValidateFunctionRule strCondition
            strCodedText = .CodedText
        End With
        
        If NumBars = -1& Then
            AutoDetectNumBars strCodedText
        End If
        lNumBars = NumBars
                
        frmTest2.AddList "cAlert.CodedText: " & m.strPriceCondition & " -->" & m.strCodedText & " (NumBars = " & Str(lNumBars) & ")"
        DebugLog "cAlert.CodedText: " & m.strPriceCondition & " -->" & m.strCodedText & " (NumBars = " & Str(lNumBars) & ")"
        astrExpressions(0) = strCodedText
        
        astrBarNames.Add "Market1"
        astrBarNames.Add "Daily"
        astrBarNames.Add "Weekly"

        MarketsInExpressions astrExpressions, 0, False, astrBarNames, Nothing, m.strPeriod
        For lIndex = 0 To astrBarNames.Size - 1
            Select Case lIndex
                Case 0:
                    g.Alerts.AddBars SymbolOrSymbolID, m.strPeriod, lNumBars
                
                Case 1:
                    g.Alerts.AddBars SymbolOrSymbolID, "Daily", lNumBars
                
                Case 2:
                    g.Alerts.AddBars SymbolOrSymbolID, "Weekly", lNumBars
                
                Case Else:
                    strSymbol = StripStr(Parse(astrBarNames(lIndex), ",", 1), Chr(34))
                    strPeriod = StripStr(Parse(astrBarNames(lIndex), ",", 2), Chr(34))
                    If Len(strPeriod) = 0 Then strPeriod = m.strPeriod
                    lSymbolID = GetSymbolID(strSymbol)
                    
                    If lSymbolID = 0 Then
                        g.Alerts.AddBars strSymbol, strPeriod, lNumBars
                    Else
                        g.Alerts.AddBars lSymbolID, strPeriod, lNumBars
                    End If
                    
            End Select
        
            frmTest2.AddList astrBarNames(lIndex)
            DebugLog astrBarNames(lIndex)
        Next lIndex
    End If
    
    CodedText = strCodedText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.CodedText"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunExpression
'' Description: Run the expression and determine whether it is true or not
'' Inputs:      None
'' Returns:     True if expression evaluates to true, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RunExpression() As Boolean
On Error GoTo ErrSection:

    Dim astrExpr As New cGdArray        ' Array of expressions to hand to engine
    Dim astrParms As New cGdArray       ' Array of parameters to hand to engine
    Dim astrBarNames As New cGdArray    ' Array of bar names to hand to engine
    Dim alResults As New cGdArray       ' Array of results back from engine
    Dim alBarHandles As New cGdArray    ' Array of bar handles to hand to engine
    Dim SecondaryMarkets As New cGdTree ' Secondary markets collection
    Dim Bars As New cGdBars             ' Bars object
    Dim strCodedText As String          ' Coded text expression
    Dim dStartDate As Double            ' Starting date to load data
    Dim hArray As Long                  ' Handle for an array
    Dim lBars As Long                   ' Index into a for loop
    Dim rc As Long                      ' Return code from a function call
    Dim dValue As Double                ' Value from the result array
    Dim lIndex As Long                  ' Index into a for loop
    
    If Len(m.strCodedText) > 0 Then
        ' Create the arrays...
        astrExpr.Create eGDARRAY_Strings
        astrParms.Create eGDARRAY_Strings
        astrBarNames.Create eGDARRAY_Strings
        alResults.Create eGDARRAY_Longs
        alBarHandles.Create eGDARRAY_Longs
        
        Set Bars = g.Alerts.GetBars(SymbolOrSymbolID, m.strPeriod, NumBars)

        If Bars.Size > 0 Then
            astrExpr.Add m.strCodedText
            dStartDate = Bars(eBARS_DateTime, 0)

            astrBarNames.Add "Market1"
            astrBarNames.Add "Daily"
            astrBarNames.Add "Weekly"

            MarketsInExpressions astrExpr, dStartDate, True, astrBarNames, Nothing, "Daily"

            hArray = gdCreateArray(eGDARRAY_Doubles, Bars.Size)
            alResults.Add hArray

            astrParms(0) = "Alert"
            If Not SetupExpressions(astrParms, astrBarNames, astrExpr) Then
                Active = False ' TLB 5/14/2014: deactivate now to keep error message from continually popping up
                InfBox "An error exists in the Alert Condition expression", "[]", , "Alert Condition"
                Exit Function
            End If
            
            FillSecondaryMarkets SecondaryMarkets, astrBarNames
            For lBars = 1 To SecondaryMarkets.Count
                alBarHandles.Num(lBars - 1) = SecondaryMarkets(lBars).BarsHandle
            Next lBars
            astrParms.Size = 1

            rc = RunExpressions(astrParms.ArrayHandle, astrBarNames.ArrayHandle, alBarHandles.ArrayHandle, alResults.ArrayHandle, ByVal 0&, ByVal 0&)
            If rc = 0 Then
                hArray = alResults.Num(0)
                dValue = gdGetNum(hArray, gdGetSize(hArray) - 1)
                If dValue <> gdNullValue(hArray) Then
                    ' if so, find first non-null item
                    For lIndex = gdGetSize(hArray) - 1 To 0 Step -1
                        dValue = gdGetNum(hArray, lIndex)
                        If dValue <> gdNullValue(hArray) Then
                            If dValue <> 0 Then RunExpression = True
                            Exit For
                        End If
                    Next
                End If
            End If
            
            ' destroy the result arrays that were created using gdCreateArray
            For lIndex = 0 To alResults.Size - 1
                gdDestroyArray alResults(lIndex)
            Next
        End If

        SetupExpressions astrParms
    End If

ErrExit:
    Set astrExpr = Nothing
    Set astrParms = Nothing
    Set astrBarNames = Nothing
    Set alResults = Nothing
    Set alBarHandles = Nothing
    Set SecondaryMarkets = Nothing
    Exit Function
    
ErrSection:
    Set astrExpr = Nothing
    Set astrParms = Nothing
    Set astrBarNames = Nothing
    Set alResults = Nothing
    Set alBarHandles = Nothing
    Set SecondaryMarkets = Nothing
    RaiseError "cAlert.RunExpression"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSecondaryMarkets
'' Description: Fill the secondary markets collection from the bar names
'' Inputs:      Bars Collection, Bar Names Array
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillSecondaryMarkets(SecondaryMarkets As cGdTree, ByVal astrBarNames As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol for the market
    Dim lSymbolID As Long               ' Symbol ID for the market
    Dim strPeriod As String             ' Period for the market
    
    SecondaryMarkets.Clear
    For lIndex = 0 To astrBarNames.Size - 1
        Select Case lIndex
            Case 0:
                strSymbol = m.strSymbol
                lSymbolID = m.lSymbolID
                strPeriod = m.strPeriod
            
            Case 1:
                strSymbol = m.strSymbol
                lSymbolID = m.lSymbolID
                strPeriod = "Daily"
            
            Case 2:
                strSymbol = m.strSymbol
                lSymbolID = m.lSymbolID
                strPeriod = "Weekly"
            
            Case Else:
                strSymbol = StripStr(Parse(astrBarNames(lIndex), ",", 1), Chr(34))
                strPeriod = StripStr(Parse(astrBarNames(lIndex), ",", 2), Chr(34))
                If Len(strSymbol) = 0 Then
                    strSymbol = m.strSymbol
                    lSymbolID = m.lSymbolID
                Else
                    lSymbolID = GetSymbolID(strSymbol)
                End If
                If Len(strPeriod) = 0 Then strPeriod = m.strPeriod
                
        End Select
        
        If m.lSymbolID = 0& Then
            Set SecondaryMarkets(lIndex + 1) = g.Alerts.GetBars(strSymbol, strPeriod, NumBars)
        Else
            Set SecondaryMarkets(lIndex + 1) = g.Alerts.GetBars(lSymbolID, strPeriod, NumBars)
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlert.FillSecondaryMarkets"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorCellOnly
'' Description: Are the only actions for this alert to color the quote board
''              cell and update message history?
'' Inputs:      None
'' Returns:     True if only color cell alert, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ColorCellOnly() As Boolean
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim bColorCell As Boolean           ' Color cell is turned on
    Dim bOther As Boolean               ' Other action is turned on
    
    bColorCell = False
    bOther = False
    
    For lIndex = 0 To eAA_NumActions - 1
        Select Case lIndex
            Case eAA_ChangeBackColor
                If Parse(m.astrActions(lIndex), ",", 1) <> "0" Then
                    bColorCell = True
                End If
                
            Case eAA_MsgHistory
                
            Case Else
                If Parse(m.astrActions(lIndex), ",", 1) <> "0" Then
                    bOther = True
                End If
        
        End Select
    Next lIndex
    
    ColorCellOnly = (bColorCell = True) And (bOther = False)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.ColorCellOnly"
    
End Function

Public Property Get LastCheckedFLag() As Boolean
    LastCheckedFLag = m.bLastCheckedStatus
End Property

Public Sub ResetLastChecked(Optional ByVal strOld As String = "")
On Error GoTo ErrSection:

    Dim strOldCondition As String       ' Old condition
    Dim strNewCondition As String       ' New condition
    Dim strEnglish As String            ' English string for the alert
    Dim bReset As Boolean               ' Reset the last checked variables?
    
    bReset = True
    If Len(strOld) > 0 Then
        strOldCondition = Trim(Parse(strOld, "Then", 1))
        strEnglish = EnglishString
        strNewCondition = Trim(Parse(strEnglish, "Then", 1))
        bReset = (strOldCondition <> strNewCondition)
    End If
    
    If bReset Then
        m.bLastCheckedStatus = False
        m.dLastCheckedTime = 0
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAlert.ResetLastChecked"
    
End Sub

Public Property Get LastCheckedTimeZone() As String
    LastCheckedTimeZone = m.strLastCheckedTimeZone
End Property

Public Property Get MyBars(nLastGoodBar As Long) As cGdBars
On Error GoTo ErrSection:
    
    Dim Chart As cChart
    
    If m.nAlertType = eGDAlertType_Annot Or m.nAlertType = eGDAlertType_Chart Then
        If Not m.Annot Is Nothing Then
            Set Chart = m.Annot.AnnotChart
        ElseIf Not m.Ind Is Nothing Then
            Set Chart = m.Ind.IndChart
        End If
        If Not Chart Is Nothing Then
            Set m.Bars = Chart.Bars
            nLastGoodBar = Chart.LastGoodDataBar(False, False)
        End If
    ElseIf Len(m.strSymbol) > 0 And Len(m.strPeriod) > 0 Then
        If m.Bars Is Nothing Then
            Set m.Bars = g.Alerts.GetBars(m.strSymbol, m.strPeriod, NumBars)
        ElseIf m.Bars.Size = 0 Or m.strSymbol <> m.Bars.Prop(eBARS_Symbol) Then
            Set m.Bars = g.Alerts.GetBars(m.strSymbol, m.strPeriod, NumBars)
        End If
    End If
    
    Set MyBars = m.Bars
    
    Exit Property

ErrSection:
    RaiseError "cAlert.MyBars"
    
End Property

Public Property Get MessageColor() As Long
On Error GoTo ErrSection:
    
    MessageColor = ValOfText(Parse(m.astrActions(eAA_MessageBox), ",", 3))
    
    Exit Property

ErrSection:
    RaiseError "cAlert.MessageColor"

End Property

Public Property Let OrderSubmitted(nOrderId As Long)
On Error GoTo ErrSection:

    Dim Order As New cPtOrder
    Dim strBrokerID$, strInfo$, strText$
    
'This is meant to attach an order number to the alert that triggered it.
'If all goes well, we will get a broker order id fairly quickly otherwise
'retry once then just show what info we have.

    strInfo = m.astrActions(eAA_PlaceOrder)   'default order info

    If Order.Load(nOrderId) Then
        strInfo = Order.OrderText
        strBrokerID = Order.BrokerID
        
        If Len(strBrokerID) = 0 Then
            strBrokerID = Parse(Order.BrokerInfoString, vbTab, 2)
        End If
        
        If Len(strBrokerID) = 0 Then
            DoEvents
            DoEvents
            DoEvents
            If Order.Load(nOrderId) Then
                strBrokerID = Order.BrokerID
                If Len(strBrokerID) = 0 Then
                    strBrokerID = Parse(Order.BrokerInfoString, vbTab, 1)
                End If
            End If
        End If
    End If
    
    strText = strInfo
    If Len(strBrokerID) > 0 Then strText = strText & "(" & strBrokerID & ")"
    
    g.Alerts.AddAlertMessage m.dLastCheckedTime, strText, Me
    
    Exit Property
    
ErrSection:
    RaiseError "cAlert.OrderSubmitted"

End Property

Public Sub QBTabAlertColorOnly(frm As Form)
On Error GoTo ErrSection:

'this sub is for coloring QB cells for QB Tab alert when switching tabs

    Dim lIndex2 As Long
    Dim astrSymbols As New cGdArray
    Dim anAlert As New cGdArray         ' Array of whether to alert or not
    Dim dCurrentValue As Double
    
    Dim lColor As Long
    Dim strSymbol As String
    
    If frm Is Nothing Then Exit Sub
    
    If g.RealTime.Active Or Not HasColorAction Then
        Exit Sub           'precautionary
    End If
    
    If Not TypeOf frm Is frmQuotes And Not TypeOf frm Is frmDetachedQBTab Then
        Exit Sub           'precautionary
    End If
    
    For lIndex2 = 0 To frmQuotes.TabRecords - 1
        If frmQuotes.TabStr(eGDTabSettings_Name, lIndex2) = m.strTabName Then
            astrSymbols.SplitFields frmQuotes.TabStr(eGDTabSettings_Symbols, lIndex2), ","      'symbolId;periodString
            Exit For
        End If
    Next
    
    If astrSymbols.Size = 0 Then
        Exit Sub
    ElseIf Not m.bSymbol Then
        If TypeOf frm Is frmQuotes And lIndex2 <> frmQuotes.CurrentTabNum(True) Then
            Exit Sub
        End If
    End If
    
    For lIndex2 = 0 To astrSymbols.Size - 1
        anAlert(lIndex2) = 0
        
        ' Get the current value from the grid
        If CLng(Val(Parse(astrSymbols(lIndex2), ";", 1))) = 0 Then
            dCurrentValue = frmQuotes.SymbolData(Parse(astrSymbols(lIndex2), ";", 1), Parse(astrSymbols(lIndex2), ";", 2), m.strField)
        Else
            dCurrentValue = frmQuotes.SymbolData(CLng(Val(Parse(astrSymbols(lIndex2), ";", 1))), Parse(astrSymbols(lIndex2), ";", 2), m.strField)
        End If
        If InStr(m.strField, "%") <> 0 Then dCurrentValue = dCurrentValue * 100
        If dCurrentValue <> kNullData Then
            Select Case UCase(m.strOperator)
                Case "<", "LESS THAN"
                    If dCurrentValue < m.dValue Then anAlert(lIndex2) = True
                Case ">", "GREATER THAN"
                    If dCurrentValue > m.dValue Then anAlert(lIndex2) = True
                Case "=", "EQUAL"
                    If dCurrentValue = m.dValue Then anAlert(lIndex2) = True
                Case "<=", "LESS THAN OR EQUAL"
                    If dCurrentValue <= m.dValue Then anAlert(lIndex2) = True
                Case ">=", "GREATER THAN OR EQUAL"
                    If dCurrentValue >= m.dValue Then anAlert(lIndex2) = True
            End Select
        End If
    Next lIndex2

    lColor = ValOfText(Parse(m.astrActions(eAA_ChangeBackColor), ",", 2))
    For lIndex2 = 0 To anAlert.Size - 1
        If (anAlert(lIndex2) <> 0) Then
            strSymbol = g.SymbolPool.SymbolForID(ValOfText(Parse(astrSymbols(lIndex2), ";", 1)))
            frm.ColorCell strSymbol, Parse(astrSymbols(lIndex2), ";", 2), m.strField, m.strTabName, lColor
        End If
    Next lIndex2
    
    Exit Sub

ErrSection:
    RaiseError "cAlert.QBTabAlertColorOnly"
    
End Sub

Public Sub ResetUseCloseOfBar(Optional Bars As cGdBars = Nothing, Optional ByVal nBarIdx As Long = -1, _
    Optional ByVal eNewAlertBarFlag As eGDAlertBarOption = eGDAlertBar9_Undefined)
On Error GoTo ErrSection:

    Dim CurrBars As cGdBars
    
    Dim nLastGoodBar&
    Dim dHigh#, dLow#, dMinMove#, dY#
        
    If Bars Is Nothing Then
        Set CurrBars = MyBars(nLastGoodBar)
    Else
        Set CurrBars = Bars
        nLastGoodBar = nBarIdx
    End If
    
    If CurrBars Is Nothing Then Exit Sub
    
    If nLastGoodBar <= 0 Then nLastGoodBar = CurrBars.Size - 1
    
    m.eAlertBarFlag = eGDAlertBar_HiLowThisBar      'reset
    
    dMinMove = CurrBars.MinMove(CurrBars(eBARS_DateTime, nLastGoodBar))
    dLow = RoundToMinMove(CurrBars(eBARS_Low, nLastGoodBar), dMinMove)
    dHigh = RoundToMinMove(CurrBars(eBARS_High, nLastGoodBar), dMinMove)
    
    If m.nAlertType = eGDAlertType_Price Then
        'rule for price alert:
        '   down to price:TRUE  / up to price:FALSE - use close of bar if alert price for DOWN to price is between high & low
        '   down to price:FALSE / up to price:TRUE  - use close of bar if alert price for UP to price is between high & low
        '   down to price:TRUE  / up to price:TRUE  - use close of bar if either alert prices are between high & low
        '   down to price:FALSE / up to price:FALSE - theoretically should never happen
        If UseGetsDownTo And UseGetsUpTo Then
            dY = RoundToMinMove(m.dGetsDownTo, dMinMove)
            If dY > dLow And dY < dHigh Then
                m.eAlertBarFlag = eGDAlertBar3_CloseThisBar
            Else
                dY = RoundToMinMove(m.dGetsUpTo, dMinMove)
                If dY > dLow And dY < dHigh Then m.eAlertBarFlag = eGDAlertBar3_CloseThisBar
            End If
        ElseIf UseGetsDownTo Then
            dY = RoundToMinMove(m.dGetsDownTo, dMinMove)
            If dY > dLow And dY < dHigh Then m.eAlertBarFlag = eGDAlertBar1_DownToClose
        ElseIf UseGetsUpTo Then
            dY = RoundToMinMove(m.dGetsUpTo, dMinMove)
            If dY > dLow And dY < dHigh Then m.eAlertBarFlag = eGDAlertBar2_UpToClose
        End If
    ElseIf m.nAlertType = eGDAlertType_Annot Then
        dY = RoundToMinMove(m.dValue, dMinMove)
        If dY >= dLow And dY <= dHigh Then
            If eNewAlertBarFlag = eGDAlertBar4_HiLowCompleteBar Or eNewAlertBarFlag = eGDAlertBar5_CloseCompleteBar Then
                m.eAlertBarFlag = eGDAlertBar5_CloseCompleteBar
            Else
                m.eAlertBarFlag = eGDAlertBar3_CloseThisBar
            End If
        ElseIf eNewAlertBarFlag = eGDAlertBar4_HiLowCompleteBar Or eNewAlertBarFlag = eGDAlertBar5_CloseCompleteBar Then
            m.eAlertBarFlag = eGDAlertBar4_HiLowCompleteBar
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "cAlert.ResetUseCloseOfBar"
    
End Sub

Public Property Get AlertBarFlag() As eGDAlertBarOption
    AlertBarFlag = ValidateBarFlag(m.eAlertBarFlag)
End Property

Private Function ValidateBarFlag(ByVal eFlag As eGDAlertBarOption, _
    Optional ByVal eAlertType As eGDAlertType = eGDAlertType_Undefined) As eGDAlertBarOption
On Error GoTo ErrSection:

    Dim eType As eGDAlertType
    Dim eReturnFlag As eGDAlertBarOption
    
    eType = eAlertType
    If eType = eGDAlertType_Undefined Then eType = m.nAlertType

    eReturnFlag = eGDAlertBar9_Undefined        'initialized return value
    
    If eType = eGDAlertType_Annot Then
        Select Case eFlag
            Case eGDAlertBar_HiLowThisBar, _
                 eGDAlertBar3_CloseThisBar, _
                 eGDAlertBar4_HiLowCompleteBar, _
                 eGDAlertBar5_CloseCompleteBar
                 
                 eReturnFlag = eFlag
            
            Case eGDAlertBar1_DownToClose, _
                 eGDAlertBar2_UpToClose
                 
                 eReturnFlag = eGDAlertBar3_CloseThisBar
            
            Case Else
                 eReturnFlag = eGDAlertBar_HiLowThisBar
                 
        End Select
    ElseIf eType = eGDAlertType_Price Then
        Select Case eFlag
            Case eGDAlertBar_HiLowThisBar, _
                 eGDAlertBar1_DownToClose, _
                 eGDAlertBar2_UpToClose, _
                 eGDAlertBar3_CloseThisBar
                 
                 eReturnFlag = eFlag
            
            Case eGDAlertBar4_HiLowCompleteBar
                 eReturnFlag = eGDAlertBar_HiLowThisBar
                 
            Case eGDAlertBar5_CloseCompleteBar
                 eReturnFlag = eGDAlertBar3_CloseThisBar
            
            Case Else
                 eReturnFlag = eGDAlertBar_HiLowThisBar
                 
        End Select
    End If
    
    ValidateBarFlag = eReturnFlag

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.ValidateBarFlag"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTheSame
'' Description: Is this the same alert as the string being passed in?
'' Inputs:      Alert String
'' Returns:     True if same alert, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsTheSame(ByVal strAlert As String) As Boolean
On Error GoTo ErrSection:

    Dim strThisAlert As String          ' Alert string for this alert
    Dim astrThisAlert As New cGdArray   ' This alert split out into an array
    Dim astrThatAlert As New cGdArray   ' The passed in alert split out into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim lLastChecked As Long            ' Index for the Last Checked value
    Dim lActive As Long                 ' Index for the Active value
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    strThisAlert = ToFileString
    astrThisAlert.SplitFields strThisAlert, "|"
    astrThatAlert.SplitFields strAlert, "|"
    
    If astrThisAlert.Size <> astrThatAlert.Size Then
        bReturn = False
    Else
        Select Case m.nAlertType
            Case eGDAlertType_QuoteBoard
                lLastChecked = 13
                lActive = 7
            Case eGDAlertType_AutoTrade
                lLastChecked = 8
                lActive = 4
            Case eGDAlertType_Status
                lLastChecked = 6
                lActive = 2
            Case eGDAlertType_Price
                lLastChecked = 9
                lActive = 7
            Case eGDAlertType_Time
                lLastChecked = 11
                lActive = 5
            Case eGDAlertType_Chart
                lLastChecked = 11
                lActive = 7
            Case eGDAlertType_Annot
                lLastChecked = 11
                lActive = 7
            Case eGDAlertType_TradeSense
                lLastChecked = 8
                lActive = 4
            Case Else
                lLastChecked = -1
                lActive = -1
        End Select
    
        For lIndex = 0 To astrThisAlert.Size - 1
            If (lIndex <> lLastChecked) And (lIndex <> lActive) Then
                If UCase(astrThisAlert(lIndex)) <> UCase(astrThatAlert(lIndex)) Then
                    bReturn = False
                    Exit For
                End If
            End If
        Next lIndex
    End If

    IsTheSame = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.IsTheSame"
    
End Function

Public Property Get ShowOnCharts() As Boolean
    ShowOnCharts = m.bShowOnCharts
End Property

Public Property Let ShowOnCharts(ByVal bShow As Boolean)
    m.bShowOnCharts = bShow
End Property

Public Sub RemoveChartObject(ChartIn As cChart)
On Error GoTo ErrSection:

    Dim i&

    If ChartIn Is Nothing Then Exit Sub
    If ChartIn.Form Is Nothing Then Exit Sub
    If m.aCharts Is Nothing Then Exit Sub

    For i = m.aCharts.Size - 1 To 0 Step -1
        If Not m.aCharts(i).Form Is Nothing Then
            If m.aCharts(i).Form.hWnd = ChartIn.Form.hWnd Then m.aCharts.Remove i
        End If
    Next
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAlert.RemoveChartObject"
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AddChartObject Design Note:
'Relevant sequence of events during Trade Navigator start up:
'   1. frmQuotes load list of alerts from QuoteList.ALR file (this does not call cAlerts.Add)
'   2. mMain RestoreCharts (triggers chart load & turn on frmChart timers when done)
'   3. frmChart timer call this if not already done
'
'When cAlerts.Add is called to add new price alert, it will call this subroutine with nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddChartObject(Optional ChartIn As cChart = Nothing)

    Static bInProgress As Boolean
    
    Dim i&, Chart As cChart, bFound As Boolean

    If bInProgress Or g.bUnloading Then Exit Sub
    
    bInProgress = True
    
    If m.aCharts Is Nothing Then Set m.aCharts = New cGdArray
    
    If ChartIn Is Nothing Then
        'only want to go though list of loaded charts one time when new alert gets added
        'any chart loaded after startup will add itself directly by calling this function
        If m.aCharts.Size = 0 Then
            If Not m.bFormsChecked Then
                For i = 0 To Forms.Count - 1
                    If IsFrmChart(Forms(i)) Then
                        m.bFormsChecked = True
                        Set Chart = Forms(i).Chart
                        If Not Chart Is Nothing Then
                            If Chart.Symbol = m.strSymbol Then
                                If Chart.PriceAlertAdd(Me) Then
                                    m.aCharts.Add Chart
                                    Chart.GenerateChart eRedo1_Scrolled
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
    ElseIf ChartIn.Symbol = m.strSymbol Then
        If Not ChartIn.Form Is Nothing Then
            For i = 0 To m.aCharts.Size - 1
                If Not m.aCharts(i).Form Is Nothing Then
                    If m.aCharts(i).Form.hWnd = ChartIn.Form.hWnd Then
                        bFound = True
                        Exit For
                    End If
                End If
            Next
        End If
        If Not bFound Then m.aCharts.Add ChartIn
    End If
    
    bInProgress = False

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'UpdateChartObject:
'1. called from frmAlerts when user edits a price alert
'2. called from frmAlertsSetup when user removes a price alert
'3. called from frmChart when user moves alert bell on chart
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateChartObject(ByVal bRemove As Boolean)

    Dim Chart As cChart, i&
    
    If m.nAlertType <> eGDAlertType_Price Then Exit Sub         'precautionary
    If m.aCharts Is Nothing Then Exit Sub                       'alert not set to show on chart just exit
    
    For i = 0 To m.aCharts.Size - 1
        Set Chart = m.aCharts(i)
        If Not Chart Is Nothing Then
            If Chart.PriceAlertUpdate(Me, bRemove) Then
                Chart.GenerateChart eRedo1_Scrolled
                If ShowOnCharts Then
                    SetFocus Chart.Form.pbChart.hWnd            'so Ctrl-Z works right away
                End If
            End If
        End If
    Next
    

End Sub

Public Sub WebMailStatus(ByVal strInfo$)
On Error GoTo ErrSection:
    
    Dim strMsg$
    
    If InStr(UCase(strInfo), "SUCCESS") = 0 Then
        If Len(strInfo) <= 128 Then
            strMsg = strInfo
        Else
            strMsg = Left(strInfo, 128)
        End If
        'InfBox strMsg, "I", , "Attempt to send email for alert failed."
        ShowAlertPopup strMsg, "Attempt to send email for alert failed", vbCenter
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAlert.WebMailStatus"

End Sub

Private Sub SendEmail(ByRef astrAction As cGdArray, ByVal strCaption$, ByVal strMsg$, ByVal strDate$)
On Error GoTo ErrSection:

    Static bInProgress As Boolean

    Dim aAddresses As New cGdArray
    Dim astrEmailText As New cGdArray   ' Text of the body of the e-mail
    Dim strUserID$, strPassWd$, strWeb$
    Dim i&, iEnd&
    Dim lExitCode&                      ' Exit code from sending the e-mail
    Dim strIgnore$, bIgnore As Boolean
    
    Dim frm As Form                     ' Form to use to display message box
    
    If bInProgress Then GoTo ErrExit
    
    bInProgress = True
    
    If Len(astrAction(2)) = 0 Then
        bIgnore = True
    Else
        ' TLB 6/21/2016: check to see if we should ignore sending this alert (if matches one of the pipe-delimited IgnoreAlert's)
        bIgnore = False
        strIgnore = GetIniFileProperty("IgnoreAlert", "", "General", AddSlash(App.Path) & "Provided\Provided.INI")
        aAddresses.SplitFields strIgnore, "|"
        For i = 0 To aAddresses.Size - 1
            strIgnore = Trim(UCase(aAddresses(i)))
            If Len(strIgnore) > 0 Then
                If InStr(UCase(strMsg), strIgnore) > 0 Then
                    bIgnore = True
                    Exit For
                End If
            End If
        Next
        aAddresses.Size = 0
    End If
    
    If Not bIgnore Then
        aAddresses.SplitFields astrAction(2), ";"
        
        If aAddresses.Size > 3 Then
            astrAction(2) = aAddresses(0) & ";" & aAddresses(1) & ";" & aAddresses(2)   'precautionary; theoretically should never get here
            iEnd = 2
        Else
            iEnd = aAddresses.Size - 1
        End If
        
        If Len(astrAction(3)) > 0 And Len(astrAction(4)) > 0 Then
            For i = 0 To iEnd
                astrEmailText.Size = 0
                astrEmailText.Add "To: " & aAddresses(i)
                astrEmailText.Add "From: " & astrAction(3)
                astrEmailText.Add "Date: " & strDate
                If ExtremeCharts >= 1 Then
                    astrEmailText.Add "Subject: Extreme Charts Alert for " & strCaption
                Else
                    astrEmailText.Add "Subject: Trade Navigator Alert for " & strCaption
                End If
        
                astrEmailText.Add ""
                astrEmailText.Add strMsg
                astrEmailText.ToFile "AlertEmail.txt"
        
                RunProcess App.Path & "\smtp.exe", Chr(34) & App.Path & "\AlertEmail.txt" & Chr(34) & " " & Chr(34) & astrAction(4) & Chr(34), True, vbMinimizedNoFocus, lExitCode
                If lExitCode <> 0 Then
                    Set frm = New frmAlertPopup
                    frm.ShowMe eGDAlertMode_QuoteBoardAlert, strCaption, _
                        "Email to " & aAddresses(i) & " failed. Error(" & Str(lExitCode) & ")"
                    Beep
                End If
            Next
        Else
            strMsg = "Time    : " & strDate & vbLf & strMsg
            strWeb = "http://www.TradeNavigator.com/alerts/index.aspx?U=*&P=*"
            strWeb = FixURL(GetProvidedProperty("AlertsWeb", strWeb))
            
            strMsg = Replace(strMsg, vbLf, vbCrLf)
            
            If ExtremeCharts >= 1 Then
                strWeb = strWeb & "&A=" & UrlEncodeField(astrAction(2)) & "&S=" & UrlEncodeField("Extreme Charts Alert for " & strCaption) & "&C=" & UrlEncodeField(strMsg)
            Else
                strWeb = strWeb & "&A=" & UrlEncodeField(astrAction(2)) & "&S=" & UrlEncodeField("Trade Navigator Alert for " & strCaption) & "&C=" & UrlEncodeField(strMsg)
            End If
            If Len(strWeb) >= 1024 Then strWeb = Left(strWeb, 1022)  'cannot send more than 1024 chars in URL
            If 0 Then
                ' if using a simple Internet Transfer control
                strMsg = GetWebPageData(strWeb)
                WebMailStatus strMsg
            Else
                ' if using a full-fledged Browser control
                Set frm = New frmWebReport
                frm.ShowMe "", , strWeb, , Me
                Set frm = Nothing
            End If
        End If
    End If

    bInProgress = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAlert.SendEmail"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoDetectNumBars
'' Description: Auto detect the number of bars needed for the expression
'' Inputs:      Coded Text Expression
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AutoDetectNumBars(Optional ByVal strCodedText As String = "")
On Error GoTo ErrSection:

    Dim AutoDetect As New cAutoDetect   ' Object to assist with the auto detect
    
    If Len(strCodedText) = 0 Then
        m.lNumBarsCalc = AutoDetect.AutoDetect(m.strCodedText, SymbolOrSymbolID, m.strPeriod)
    Else
        m.lNumBarsCalc = AutoDetect.AutoDetect(strCodedText, SymbolOrSymbolID, m.strPeriod)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlert.AutoDetectNumBars"
    
End Sub

Public Sub AddQbCell(QbCell As cQuoteCell)
On Error GoTo ErrSection:

    Dim i&
    Dim bFound As Boolean

    If m.aQbCells Is Nothing Or QbCell Is Nothing Then Exit Sub
    
    For i = 0 To m.aQbCells.Size - 1
        If Not m.aQbCells(i) Is Nothing Then
            If m.aQbCells(i) Is QbCell Then
                bFound = True
                Exit For
            End If
        End If
    Next
    
    If Not bFound Then m.aQbCells.Add QbCell

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlert.AddQbCell"
    
End Sub

Public Sub RemoveQbCell(QbCell As cQuoteCell)
On Error GoTo ErrSection:

    Dim i&

    If m.aQbCells Is Nothing Or QbCell Is Nothing Then Exit Sub
    If m.aQbCells.Size = 0 Then Exit Sub

    For i = m.aQbCells.Size - 1 To 0
        If Not m.aQbCells(i) Is Nothing Then
            If m.aQbCells(i) Is QbCell Then
                m.aQbCells.Remove i
                Exit For
            End If
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlert.RemoveQbCell"
    
End Sub

Public Property Let EditInProg(ByVal bUerEditing As Boolean)

    m.bEditInProg = bUerEditing
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetLastBarAlerted
'' Description: Reset the last bar alerted array
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ResetLastBarAlerted()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 0 To m.alLastBarAlerted.Size - 1
        m.alLastBarAlerted(lIndex) = kNullData
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlert.ResetLastBarAlerted"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtractSymbolFromOrderInfo
'' Description: Extract the symbol from the given order information string
'' Inputs:      Order Information
'' Returns:     Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ExtractSymbolFromOrderInfo(ByVal strOrderInfo As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lStart As Long                  ' Starting location for the symbol
    Dim lEnd As Long                    ' Ending location for the symbol

    strReturn = ""
    If Len(strOrderInfo) > 0 Then
        lStart = InStr(strOrderInfo, "'")
        If lStart > 0 Then
            lEnd = InStr(lStart + 1, strOrderInfo, "'")
            If lEnd > 0 Then
                strReturn = Mid(strOrderInfo, lStart + 1, lEnd - lStart - 1)
            End If
        End If
    End If

    ExtractSymbolFromOrderInfo = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.ExtractSymbolFromOrderInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckFromCheckAlerts
'' Description: Determine if this alert is allowed to be checked from CheckAlerts
'' Inputs:      None
'' Returns:     True if allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckFromCheckAlerts() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    If IsOrderChangeStatusAlert Then
        bReturn = False
    ElseIf IsBrokerDisconnectAlert Then
        bReturn = False
    ElseIf IsPositionMismatchStatusAlert Then
        bReturn = False
    ElseIf IsAutoTradeItemDisabledStatusAlert Then
        bReturn = False
    Else
        bReturn = True
    End If
    
    CheckFromCheckAlerts = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.CheckFromCheckAlerts"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckAlertForTradeItem
'' Description: Check the alert for the given automated trading item
'' Inputs:      Automated trading item
'' Returns:     0 if false, 1 if true
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckAlertForTradeItem(ByVal TradeItem As cAutoTradeItem) As Byte
On Error GoTo ErrSection:

    Dim nReturn As Byte                 ' Return value for the function

    If Not TradeItem Is Nothing Then
        Select Case UCase(m.strAutoTradeCondition)
            Case "ORDER SUBMITTED"
                nReturn = Abs(TradeItem.OrderSubmitted)
                TradeItem.OrderSubmitted = False
            Case "ORDER PRICE HIT"
                nReturn = Abs(TradeItem.OrderPriceHit)
                TradeItem.OrderPriceHit = False
            Case "ORDER FILLED"
                nReturn = Abs(TradeItem.OrderFilled)
                TradeItem.OrderFilled = False
            Case "ORDER CANCELLED"
                nReturn = Abs(TradeItem.OrderCancelled)
                TradeItem.OrderCancelled = False
            Case "ORDER REJECTED"
                nReturn = Abs(TradeItem.OrderRejected)
                TradeItem.OrderRejected = False
        End Select
    End If
    
    CheckAlertForTradeItem = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.CheckAlertForTradeItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidAccountForAutoTradeAlert
'' Description: Is the given account ID valid for this alert?
'' Inputs:      Account ID, Account Type
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidAccountForAutoTradeAlert(ByVal lAccountID As Long, ByVal nAccountType As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If m.nAlertType = eGDAlertType_AutoTrade Then
        Select Case m.lTradeItemID
            Case 0:
                bReturn = True
                
            Case -1:
                bReturn = g.Broker.IsLiveAccount(nAccountType)
            
            Case -2:
                bReturn = Not g.Broker.IsLiveAccount(nAccountType)
            
            Case Else:
                bReturn = (m.lTradeItemID = lAccountID)
                
        End Select
    End If
    
    ValidAccountForAutoTradeAlert = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlert.ValidAccountForAutoTradeAlert"
    
End Function
